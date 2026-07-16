# First Intelligence | Business Performance

Aplicativo Streamlit com visão consolidada para a diretoria e acessos restritos por gestor de linha.

## Bases utilizadas

- `BASE BI.xlsx`: faturamento, clientes, produtos, gerentes, vendedores e metas.
- `rev2026 Base bi.xlsx`: receitas e despesas realizadas por centro de custos. Somente abas visíveis são processadas.
- `base_crm_cobranca.csv`: inadimplência exportada do CRM de cobrança, com cliente, vendedor, gerente, vencimento e saldo atual. O formato Excel anterior permanece compatível.

## Visões disponíveis

- Dashboard integrado de performance.
- Desempenho e metas.
- Comparação das linhas de negócio para a diretoria.
- Recebimentos e inadimplência.
- Clientes.
- Produtos.
- Centro de custos, com receitas e despesas operacionais e não operacionais.

Os módulos de preços e estoque permanecem fora desta versão para amadurecimento separado.

## Regra de centro de custos

Toda visão por departamento e por linha utiliza exclusivamente a coluna `CENTRO DE CUSTOS` da aba visível `Centro de Custos`.

A coluna `CENTRO DE CUSTOS RATEAO` é removida durante o carregamento e não participa de cálculos, filtros ou visualizações.

A página **Centro de custos** permite filtrar:

- receitas e despesas;
- movimentos operacionais e não operacionais;
- departamento;
- natureza;
- fornecedor ou cliente;
- período.

## Regra da equipe comercial

Os rankings, gráficos e tabelas de desempenho exibem somente vendedores ou representantes com participação efetiva e percentual superior a zero no período selecionado.

## Perfis

- Diretoria: visão consolidada e comparação entre linhas.
- Celso: Microtech.
- Renato: Vendas.
- Amauri: Locação.
- Ronaldo: Endoscopia.

Cada gestor visualiza apenas a própria linha, determinada pelo centro de custos direto correspondente.

## Publicação

1. Envie o conteúdo desta pasta para um repositório privado.
2. No Streamlit Cloud, selecione `app.py` como arquivo principal.
3. Copie o conteúdo de `SECRETS_STREAMLIT_PRONTO.toml` para **Manage app → Settings → Secrets**.
4. Reinicie o aplicativo.

## Atualização das bases

A diretoria pode substituir as três bases na seção **Fontes de dados** da barra lateral. O CSV do CRM deve utilizar separador `;` e os campos do arquivo modelo incluído no pacote.
