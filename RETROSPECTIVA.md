<!--
  Versão sanitizada para publicação. Nomes de cliente, fornecedor, credenciais,
  PIDs e usuário de SO foram trocados por rótulos genéricos. As automações que
  monitoram equipamentos e as credenciais de acesso ficam fora deste repositório
  (ver "O que não está aqui" no README) — este documento descreve a arquitetura
  e o processo de investigação, não expõe o sistema real.
-->

# Retrospectiva — incidente "o dashboard volta sozinho para uma versão antiga"

Análise de um incidente investigado ao longo de um dia: o dashboard, servido a
partir de uma máquina-monitor, era periodicamente sobrescrito por uma versão
antiga da própria página. Reúne o que aconteceu, os erros de arquitetura que
tornaram o bug possível, os erros cometidos durante a própria investigação, as
soluções aplicadas e os caminhos que teriam sido melhores.

---

## 1. Contexto

O sistema monitora dois conjuntos de equipamentos (aqui **Campo A** e
**Campo B**), grava dados de conectividade em planilhas `.xlsx`, dispara reset
em equipamentos/chips offline e gera um dashboard HTML estático
(`visualizador/index.html`) servido em `localhost:2000` e exposto por um túnel.
Vários processos Python rodam de forma contínua:

| Processo | Papel |
|----------|-------|
| `observador_a.py` | Observa o Campo A, grava dados |
| `observador_b.py` | Observa o Campo B, grava dados |
| `reset_chips.py` | Reseta chips M2M via API da operadora (a cada ~1h) |
| `atualizar_site.py` | Regenera o `index.html` (a cada 5 min) |
| `serve_no_cache.py` | Servidor HTTP da porta 2000 |
| Tarefas agendadas `Vigia*` | Watchdogs que religam alguns processos |

## 2. Sintoma relatado

"Toda vez que recarrego a página, parece que carrega a versão antiga e depois
mascara para a versão nova." E, de forma intermitente, o dashboard aparecia
totalmente na versão antiga (layout sem os gráficos de rosca / donut).

Duas causas independentes foram encontradas por trás do mesmo sintoma.

## 3. Causa A — "pisca e mascara para a versão nova" (latência de CDN)

O gerador carregava Chart.js e a lib de planilha de um CDN externo. Medido: a
página levava **~5 segundos** para baixar essas bibliotecas. Nesse intervalo o
HTML e os dados (locais) já apareciam, mas os gráficos ficavam vazios até o CDN
responder — dando a impressão de "carregou velho e depois virou novo". Os
centros vazios dos donuts em um dos screenshots eram exatamente esse quadro
transitório.

**Não era** cache de navegador nem do túnel, e **não era** bug de código de
renderização (verificado com Playwright: os centros dos donuts renderizam
corretamente quando a página termina de carregar).

### Solução aplicada
No sistema de produção, as bibliotecas foram baixadas para
`visualizador/vendor/` e os `<script>` passaram a apontar para o caminho local.
Resultado medido: tempo de carga caiu de **5,03 s → 1,65 s**, com zero
requisição externa. (O `gerar.py` deste repositório de exemplo ainda usa CDN —
é um scaffold, não a máquina-monitor.)

## 4. Causa B — "volta para a versão antiga" (código velho preso na memória)

Este foi o problema difícil. O `index.html` só é escrito em um lugar:
`gerar.main()`. O `gerar.py` em disco tem **um único template** e produz
**somente** a versão nova (donut + libs locais) — não existe um "template
antigo" que ele pudesse emitir por engano.

Ainda assim, o arquivo era sobrescrito periodicamente pela versão antiga (com
CDN, sem os donuts). A única explicação compatível com todas as evidências:
**um processo Python de vida longa importou o `gerar` no passado (antes das
edições de layout) e mantém essa versão antiga do código congelada na memória,
regenerando o site com ela.**

Quando você faz `from visualizador import gerar` e o processo fica rodando por
dias, o Python compila e guarda aquela versão do módulo em memória. Editar o
`gerar.py` no disco depois **não** afeta o processo já rodando — ele continua
executando a versão antiga até ser reiniciado.

Isto explica por que os "consertos" pareciam não pegar: cada vez que o site era
regenerado à mão para a versão nova, o processo fantasma o sobrescrevia de volta
para a versão antiga no ciclo seguinte.

### Como foi diagnosticado
Depois de várias tentativas de deduzir o culpado por horário de início de
processo e data de modificação de arquivo (abordagem que se mostrou lenta e
enganosa), a identificação veio de uma **armadilha empírica**: um monitor que
checa o `index.html` a cada poucos segundos e, no instante em que ele reverte
para a versão antiga, tira uma foto de todos os processos Python. As reversões
aconteciam a cada ~7 minutos.

Um **instrumento temporário** foi adicionado ao `gerar.main()` para registrar
quem (PID/PPID) gerava o site a cada escrita. A lógica: a versão atual do
`gerar` loga; um escritor com o `gerar` antigo congelado na memória **não**
loga. Nas reversões, nenhuma geração aparecia no log no instante do write velho
— provando que o reverter executava código antigo em memória, não o `gerar.py`
do disco.

O primeiro palpite (o `reset_chips.py`, por ser o único de vida longa que
importava `gerar`) estava **errado**: reiniciá-lo não parou as reversões. Isso
é o erro 5.4 na prática. A resposta certa veio ao notar que a geração do site
**foi removida do observador num refactor** — o próprio comentário no
`observador_a.py` diz "o site roda separado agora, ver atualizar_site.py". Mas o
processo do observador do Campo A estava rodando **de antes desse refactor**,
ainda executando a versão antiga que gerava o site. A cada ciclo dele, o
`index.html` era sobrescrito com o layout velho.

### Solução aplicada
- O `observador_a.py` foi reiniciado a partir do código atual, que **não** gera
  mais o site. Um monitor de confirmação rodou por 15 minutos, atravessando mais
  de duas janelas de reversão: **0 reversões**. Causa e culpado confirmados.
- O `atualizar_site.py` já havia sido reescrito para rodar o `gerar.py` como
  **subprocesso** (`python gerar.py`) em vez de `import gerar` uma vez só — assim
  ele relê o arquivo do disco a cada ciclo e nunca executa código velho. Esse é
  o padrão correto que os demais geradores deveriam seguir.
- O instrumento temporário em `gerar.main()` foi **removido** após o diagnóstico.

### A lição central deste caso
O refactor moveu a geração do site para fora do observador, mas **o processo do
observador nunca foi reiniciado** desde então. Editar/refatorar código não tem
efeito sobre um processo Python de vida longa que já o carregou na memória. Todo
refactor que muda o comportamento de um processo contínuo exige reiniciar esse
processo — senão a versão antiga continua rodando, invisível, por dias.

---

## 5. Erros cometidos durante a investigação (processo de debug)

Registrados para não repetir.

**5.1 — Declarar "resolvido" cedo demais.**
O conserto foi anunciado como pronto duas vezes antes de reverter de novo. Um
bug cujo sintoma é "volta a cada X minutos" **não** se confirma com uma
verificação única; exige observar estabilidade ao longo de vários ciclos.

**5.2 — Deduzir em vez de instrumentar.**
Muito tempo foi gasto teorizando "quem escreve o arquivo" a partir de horário de
processo e mtime de arquivo — e a teoria contradizia a evidência. O certo era
instrumentar/armar a armadilha logo no primeiro sintoma. Dado bruto mata
suposição.

**5.3 — Assinatura de diagnóstico não-única (falso positivo).**
A primeira checagem de "versão antiga" usou um texto que existe **nas duas**
versões, levando a uma conclusão errada. O marcador de discriminação tem que ser
único da versão nova.

**5.4 — Agir sobre teoria não confirmada.**
Um processo foi morto por suposição; o sintoma voltou, provando a teoria errada.
Confirmar antes de agir — ou usar uma ação que seja simultaneamente teste e
correção reversível.

---

## 6. Erros de planejamento do projeto (arquitetura)

Do mais grave para o menor. São estes que tornaram o bug possível.

**6.1 — Código compartilhado importado por processos de vida longa.**
`from visualizador import gerar` dentro de scripts que rodam por dias congela o
código na memória. É a **raiz** do "volta para a versão antiga".
*Melhor caminho:* nenhum processo de longa duração deve importar código que é
editado com frequência e gerar a partir dele. Quem precisa regenerar deve rodar
o gerador como subprocesso/script fresco (relê o disco).

**6.2 — Vários escritores no mesmo arquivo de saída, sem dono único.**
Mais de um processo capaz de gerar o `index.html` cria uma corrida (um grava
novo, outro grava velho).
*Melhor caminho:* um único processo dono da geração do site. Todo o resto só lê.

**6.3 — Sem deploy/reinício ao mudar código e sem carimbo de versão.**
Editar um `.py` com o processo rodando não tem efeito, e não havia como saber
qual versão estava no ar.
*Melhor caminho:* (a) reiniciar o processo dono ao alterar seu código; (b)
carimbar o HTML gerado com um identificador de build (`<!-- build: hash -->`),
tornando "velho vs novo" um `grep` inequívoco.

**6.4 — Processos soltos, iniciados à mão.**
`Popen`/Task Scheduler misturados, sem gerenciador único. Resultado: PIDs
órfãos, risco de duplicata, ninguém sabe quantos rodam.
*Melhor caminho:* um gerenciador de serviços com um processo nomeado por papel,
reinício automático e um lugar só para ver o que está no ar.

**6.5 — Watchdogs com cobertura desigual.**
Os vigias religam observador e reset, mas **não** cobrem o `atualizar_site` nem
o servidor. Se estes caem, ninguém levanta.
*Melhor caminho:* cobertura uniforme — todo processo essencial vigiado, de forma
consistente.

**6.6 — Planilha `.xlsx` como banco de dados.**
Abrir o `.xlsx` mestre trava a gravação, o que forçou o hack do "arquivo view"
(cópia ordenada).
*Melhor caminho:* um banco leve (SQLite) elimina o lock. **Ressalva:** isto
conserta a dor de arquivo, mas **não** conserta o bug deste incidente — um
processo com código velho gravaria layout velho no banco do mesmo jeito. Banco é
sobre *dado*; o bug era sobre *código desatualizado rodando*.

**6.7 — Dependência de CDN externo para bibliotecas.**
Numa máquina-monitor, buscar libs de fora adiciona latência visível.
*Melhor caminho:* servir as libs localmente desde o início.

**6.8 — Acoplamento por caminho relativo entre muitos scripts.**
Tudo se referencia por `..\..\`; qualquer mudança de estrutura quebra em
silêncio.
*Melhor caminho:* um módulo único de configuração de caminhos, ou empacotar o
projeto de forma que os imports não dependam da posição no disco.

**6.9 — Sem log central / observabilidade.**
Não existia registro de "quem gerou o site e quando", por isso foi preciso armar
uma armadilha em vez de ler um log.
*Melhor caminho:* um log estruturado por processo.

**6.10 — Estado apenas local, numa máquina.**
Ponto único de falha. (Decisão consciente, mas registrada como risco.)

---

## 7. Prioridade recomendada

Se for arrumar, atacar nesta ordem — os três primeiros matam a classe inteira do
bug deste incidente:

1. **6.1** — nenhum processo de vida longa gera a partir de `import` de código
   editável; usar subprocesso.
2. **6.2** — um dono único da geração do site.
3. **6.3** — carimbo de versão no HTML + reinício ao mudar código.
4. **6.4 / 6.5** — gerenciador de serviços e watchdog uniforme.
5. **6.6** — SQLite no lugar do `.xlsx` como fonte de dados (resolve lock, não o
   bug deste incidente).

O fio condutor de tudo: **muito processo solto + código compartilhado importado
+ zero deploy/versionamento**. Não é um problema de arquivo-versus-banco; migrar
para banco sem resolver 6.1–6.3 apenas reproduz o mesmo caos com um banco novo.

---

## 8. Padrões recorrentes (incidentes seguintes)

Nas semanas seguintes, outros incidentes no mesmo sistema confirmaram os itens
6.1–6.5 repetidamente. Os padrões, destilados:

**Processo iniciado à mão, fora do vigia padrão, é ponto cego.** Um observador
travou por dois dias acumulando dezenas de popups de erro porque não existia
task agendada para o watchdog dele — só para os outros dois. *Regra:* antes de
declarar um processo "coberto", conferir o agendador de verdade, não só a
existência do script de vigia no repo.

**"Religa se travar" só cobre falha que se anuncia como exceção.** Vazamento de
memória de sessão longa do navegador degradava o processo aos poucos sem nunca
lançar erro — passava batido pelo `except` e pelo watchdog. Chegou ao ponto de a
RAM ficar tão crítica que nem o restart do vigia subia. *Fix:* reciclagem
proativa por tempo (fecha e reabre o navegador a cada N horas), reaproveitando a
mesma função já usada no caminho de crash. Manutenção preventiva por tempo, não
só reativa por sintoma. O intervalo (8 h) é um chute sem medição de MB/hora
vazados — ajustável numa constante.

**Heartbeat tem que vir da fonte incondicional, não de um proxy.** Duas vezes na
mesma investigação, um watchdog e depois uma métrica do dashboard usaram o
`.log` de eventos como sinal de vida. Mas o log só ganha linha quando há algo a
reportar — com a frota 100% saudável ele fica quieto, e o "quieto" era lido como
"travado". Um watchdog matou um processo perfeitamente vivo por isso; a métrica
"tempo perdido" inflou para 74% contando silêncio saudável como script caído.
*Fix:* checar o `.xlsx` (que grava um divisor a cada ~60 s, sempre) como
heartbeat real. *Regra do projeto:* heartbeat sempre do divisor incondicional,
nunca do log condicional.

**Fix aplicado na raiz, mas não nos callers vizinhos.** Um cache de token
compartilhado entre dois processos (um rápido e frequente, outro lento e raro)
sofria corrida: o rápido renovava o token no meio do loop do lento, revogando o
antigo no servidor, e o resto do lote tomava `403` registrado como "falha"
comum. O tratamento certo (no primeiro `401/403`, forçar login novo e reenviar)
já existia — só num dos dois consumidores. Mesma classe: raiz corrigida num
lugar, gap num caller que ninguém checou.

**Rotação de nome de arquivo tem que atualizar todo leitor.** Logs passaram de
`nome.ext` para `nome_AAAA-MM.ext`; quem escrevia foi atualizado, um dos
leitores não — e passou a calcular métricas a partir de um snapshot morto na
data da rotação.

**Comparar timestamp como string quebra na virada de mês.** `"01-08" < "31-07"`
é verdadeiro por comparação de string. Esse bug foi corrigido e comentado numa
função — e reapareceu dois commits depois, em código novo escrito minutos antes,
porque o comentário existente não foi conferido. *Regra:* qualquer
ordenação/comparação de timestamp `"DD-MM-AAAA HH:MM:SS"` passa por
`datetime.strptime` primeiro, mesmo "só copiando o padrão do lado".

**Auto-incidente: self-test escreveu em dado de produção.** Rodar
`registro_excel.py check` para validar sintaxe gravou linhas de teste no
histórico real, porque `_check()` reatribuía uma global (`XLSX_PATH`) que já
estava **vinculada como valor default de parâmetro na definição da função**
(resolvida no import, não na chamada). As linhas de teste com datas não-parseáveis
puseram um observador em crash-loop de ~15 min por reinício. *Lição:* um
self-test que muda estado global compartilhado com um default de parâmetro preso
a esse mesmo global é uma armadilha clássica de Python — e um `check` que grava
arquivo merece o mesmo cuidado de qualquer script de teste: confirmar isolamento
do dado real antes de rodar, não assumir pelo nome.

**Séries temporais de clocks independentes não se combinam por timestamp exato.**
Dois observadores com ciclos próprios e sem sincronismo: o timeline combinado
recontava por correspondência exata de timestamp, então um equipamento offline
há dias "caía para 0" em todo tick do outro site (onde não havia leitura nova
dele) — flapping puramente artefato de amostragem. *Fix:* forward-fill por site
antes de somar (cada site mantém seu último valor conhecido até o próximo tick
dele).
