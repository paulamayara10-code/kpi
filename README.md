# First Intelligence — KPIs de Caixa v8

Painel Streamlit em **regime de caixa**, com visão consolidada para a diretoria e acesso restrito por gestor/linha.

## Correções desta versão

- Removida a dependência `rapidfuzz`, evitando o erro `ModuleNotFoundError` no Streamlit Cloud.
- A similaridade de clientes agora usa apenas bibliotecas nativas do Python.
- O relatório do CRM de cobrança foi incluído como fonte padrão: `relatorio_cobranca_gerente.xlsx`.
- O app reconhece automaticamente as abas `Resumo` e `Titulos detalhados`.
- Datas numéricas do Excel são convertidas corretamente para o padrão brasileiro.

## Gestores e linhas

A atribuição do faturamento usa prioritariamente a coluna **GERENTE** da BASE BI:

- Celso → Microtech
- Renato → Vendas
- Amauri → Locação
- Ronaldo → Endoscopia

Registros de outros gerentes usam a classificação complementar por produto, segmento e tipo de receita.

## Inadimplência do CRM

O arquivo atual possui:

- `Resumo`: total por gerente, clientes, títulos, valor e maior atraso.
- `Titulos detalhados`: cliente, título, vencimento, valor corrigido, atraso, histórico e telefones.

Como a aba detalhada não possui a coluna Gerente, o app relaciona cada cliente ao gerente dominante da BASE BI. A diretoria visualiza também o resumo original exportado pelo CRM para conferência.

Campos priorizados no cálculo:

- Cliente: `Nome`
- Título: `Prf-Numero Parcela`
- Vencimento: `Vencto Real`
- Saldo vencido: `Tit Vencidos Valor Corrigido`
- Dias de atraso: `Dias Atraso`

Para máxima precisão futura, recomenda-se incluir a coluna **Gerente** também na aba de títulos detalhados do relatório do CRM.

## Nova análise de produtos

A página **Produtos** apresenta, dentro do escopo autorizado:

- faturamento por produto;
- quantidade faturada ou número de registros, quando não houver quantidade;
- preço médio;
- clientes e notas por produto;
- produto líder;
- concentração dos dez maiores produtos;
- ranking por faturamento e volume;
- exportação da análise em Excel.

Cada gestor visualiza somente os produtos associados ao próprio gerente/linha.

## Controle de acesso

O arquivo `.streamlit/secrets.toml.example` já contém os hashes das senhas informadas para:

- Celso / Microtech
- Renato / Vendas
- Amauri / Locação
- Ronaldo / Endoscopia

Para publicar:

1. Abra o app no Streamlit Cloud.
2. Acesse **Manage app > Settings > Secrets**.
3. Cole o conteúdo de `.streamlit/secrets.toml.example`.
4. O hash da senha da diretoria já está preenchido para o usuário Paula.

Sem Secrets configurados, o app permanece em modo demonstração para validação.

Acesso da diretoria: usuário `paula` ou e-mail `paulamayara10@gmail.com`. A senha definida pela administradora foi convertida para SHA-256 e não aparece em texto aberto nos arquivos de configuração.

## Fontes do projeto

- `BASE BI.xlsx`
- `rev2026 Base bi.xlsx`
- `relatorio_cobranca_gerente.xlsx`

Na REV2026, somente abas visíveis são processadas. Custos compartilhados e rateios não são exibidos aos gestores.
