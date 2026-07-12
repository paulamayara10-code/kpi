# App de KPIs Executivos — First Medical

## Arquivos do projeto
- `app.py`: aplicação Streamlit.
- `BASE BI.xlsx`: base comercial padrão.
- `rev2026 Base bi.xlsx`: base financeira padrão.
- `requirements.txt`: dependências.
- `runtime.txt`: versão do Python recomendada para o Streamlit Cloud.
- `.streamlit/config.toml`: tema visual.

## Publicação no Streamlit Cloud
1. Crie ou utilize um repositório privado no GitHub.
2. Envie todos os arquivos desta pasta, mantendo os nomes das duas bases.
3. No Streamlit Cloud, selecione o repositório e informe `app.py` como arquivo principal.
4. O app lê automaticamente as bases do repositório. Também é possível substituí-las temporariamente pelo menu lateral.

## Regra da REV2026
A aplicação verifica o estado das abas com `openpyxl` e processa apenas as abas marcadas como visíveis. Abas ocultas são ignoradas automaticamente.

Para os indicadores financeiros, o app utiliza:
- `Resumo Despesas` para o total oficial de saídas e classificação gerencial por natureza.
- `Centro de Custos` para receitas operacionais e abertura por centro de custo/rateio.
- `Resumo Receitas`, `Caixa Operacional` e `Performance Recebimento` para caixa e recebimentos.

## Indicadores
- Faturamento e atingimento da meta
- Margem de contribuição estimada
- EBITDA gerencial de caixa e margem EBITDA
- Recebimentos e performance de recebimento
- Geração de caixa e ponto de equilíbrio
- Rankings comerciais e concentração de clientes
- Custos por natureza, centro de custo e rateio

## Observação metodológica
O EBITDA atual é gerencial em regime de caixa. Para EBITDA contábil oficial, integrar balancete/DRE por competência, CMV, depreciação, amortização e despesas financeiras.
