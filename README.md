# Comparativo de Faturamento

Aplicação web para comparar notas fiscais entre três fontes de dados:

- **Canal A** — exportação em `.xlsx`
- **Canal B** — exportação em `.xls`
- **Contabilidade** — exportação em `.xlsx`

## Funcionalidades

- Upload de arquivos por canal
- Filtros por situação, natureza, série e faixa de NF
- Resumo por natureza de operação com totais consolidados
- Relatório de notas com diferença (presença/ausência entre sistemas)
- Gráfico comparativo por natureza
- Download dos resultados em Excel

## Como executar

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Requisitos

Ver `requirements.txt`.
