# Painel de Vendas

Cadastro de vendas com cálculo de frete e dashboard. Cada venda registrada é gravada na planilha `vendas.xlsx`.

## Arquivos
- `index.html`: estrutura da página
- `style.css`: visual
- `script.js`: formulário, cálculo e dashboard
- `servidor.py`: entrega a página e grava as vendas no Excel
- `vendas.xlsx`: planilha (aba "Vendas", com as tabelas de preços e de frete ao lado)

## Como usar
1. Instale o Python 3 (python.org) e, no terminal, dentro desta pasta: `pip install openpyxl`
2. Rode: `python servidor.py` (em alguns computadores: `python3 servidor.py`)
3. O navegador abre em http://localhost:8000. Para parar, Ctrl+C no terminal.

## Observações
- Feche `vendas.xlsx` no Excel antes de registrar ou remover vendas. Com o arquivo aberto, o Excel o bloqueia e a página avisa.
- A coluna Total da planilha é fórmula: Quantidade × (Valor Unitário + Frete). O frete é por unidade.
- Para mudar preços ou fretes, edite `PRECOS` e `FRETE` no início de `servidor.py`.
- Abrir `index.html` direto no navegador não funciona: a página precisa do `servidor.py` rodando.
