# Observador de Anomalias

Dashboard estático (sem backend, sem build step) com duas visões operacionais:

1. **Comunicação dos equipamentos** — quais equipamentos ficaram fora do ar, por quanto
   tempo, quantos ficaram offline ao mesmo tempo, e quantos resets foram enviados.
2. **Chips e consumo** — status de conexão e consumo de franquia de dados dos SIM cards
   M2M/IoT, puxado ao vivo da [API DATATEM](https://www.datatem.com.br/).

Um único script Python lê os dados (planilha de monitoramento + API de SIM cards) e gera
um `index.html` autocontido: todo o payload vira JSON embutido na página, e o Chart.js
(via CDN) desenha os gráficos no navegador. Não precisa de servidor rodando.

## Ver sem configurar nada

```
python gerar_exemplo.py
```

Gera `exemplo/index.html` com dados 100% fictícios (não chama nenhuma API), só pra
visualizar o layout e as interações.

## Funcionalidades

**Aba de equipamentos** (tema *dark tech*, inspirado em Linear/Vercel):
- Ranking de equipamentos por tempo offline
- Linha do tempo de quantos ficaram offline simultaneamente
- Distribuição da simultaneidade
- Tabela com ciclos offline e resets enviados por equipamento

**Aba de chips** (tema *dark glassmorphism*):
- 4 KPIs no topo: chips ativos, consumo do mês (GB), chips críticos (sem comunicação
  há mais de 48h), % de franquia consumida
- 4 gráficos lado a lado, **clicáveis** — clicar numa fatia/barra filtra a tabela
  inteira (status de conexão, cliente, operadora, faixa de consumo). Clicar de novo
  remove o filtro.
- Barra de filtros ativos com botão para remover individualmente ou limpar tudo
- Busca por texto + seletores de cliente/operadora
- Chips de status rápido (Todos / Atualizados / Atrasados / Sem comunicação)
- Tabela com cabeçalho fixo (sticky) ao rolar, ordenação tri-estado em qualquer coluna
  (ascendente → descendente → neutro, nulos sempre por último)
- Seletor de colunas (exibir/ocultar)
- Exportar CSV dos dados filtrados na tela
- Ação rápida por linha: copiar ICCID, ver detalhes completos do chip
- Linhas de chips críticos destacadas visualmente

Status de comunicação de cada chip é calculado pelo tempo desde a última troca de
dados reportada pela operadora: **Atualizado** (≤ 2h), **Atrasado** (≤ 24h),
**Sem comunicação** (> 24h ou nunca reportou).

## Configuração

```
pip install -r requirements.txt
cp .env.example .env   # preencha com suas credenciais DATATEM
```

Coloque seus dados de monitoramento de equipamentos na raiz do projeto:
- `registro_offline_view.xlsx` — planilha com colunas `codigo`, `data_desativação`,
  `campo`, `registrada` (timestamp do ciclo) e `reset` (sim/não). Linhas divisórias no
  formato `---- DD-MM-AAAA HH:MM:SS ----` marcam cada ciclo de verificação.
- `monitor_fundo.log` — log do processo de monitoramento, usado só para calcular
  há quanto tempo o sistema está observando e se houve interrupções.

```
python gerar.py
```

Gera `index.html` na raiz e imprime um resumo no terminal. Rode de novo sempre que
quiser atualizar os dados.

## Integração com a API DATATEM

Login (`POST /authorization/api/v1/auth/login`, senha em base64) → token Bearer →
`GET /simcard/connection/v1/connections` (paginado) traz, por chip: status de conexão,
franquia contratada, bytes consumidos no ciclo, cliente final, cidade, operadora e
timestamp da última comunicação. O script pagina até a API retornar uma página vazia.

## Stack

- Python 3 (`requests`, `python-dotenv`, `openpyxl`) só para ler os dados e montar o JSON
- Chart.js 4 via CDN para os gráficos
- HTML/CSS/JS puro no restante — sem framework, sem bundler

## O que não está aqui

Este repositório é só o dashboard. As automações que efetivamente monitoram
equipamentos e credenciais de acesso a sistemas de clientes ficam fora do controle de
versão público, por segurança.
