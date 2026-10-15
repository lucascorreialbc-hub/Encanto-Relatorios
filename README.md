# Relatórios Encanto (Trinks)

Scripts que geram os relatórios do salão Encanto Progressivas a partir da API da Trinks.

- `run.py manha` — fechamento do dia anterior e plano do dia (na segunda inclui a semana anterior)
- `run.py tarde` — parcial do dia atual

A chave da API **não** fica neste repositório: é passada pela variável de ambiente `TRINKS_KEY`.
Saídas (ignoradas pelo git): `out/` com a planilha Excel, `email.json` e `painel.html`.
