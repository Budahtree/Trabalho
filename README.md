# Trabalho — Avaliação de Investimentos

## Modelo_Avaliacao_Investimentos.xlsx

Modelo-padrão em Excel para avaliar a rentabilidade e a viabilidade de projetos de turismo e imobiliário (hotelaria, com ou sem venda de unidades). Projeção anual até 20 anos, com três cenários calculados em simultâneo.

| Folha | Para quê |
|---|---|
| Início | Instruções, legenda de cores e lógica do modelo |
| Dashboard | Veredicto (Viável / Com reservas / Não viável), KPI, comparação de cenários, fontes e aplicações, gráficos |
| Pressupostos | Todos os inputs: calendário, cenários, operação, custos, vendas, fiscalidade, financiamento, saída, critérios |
| Capex | Rubricas de investimento por classe, período e n.º de períodos |
| Pessoal | Quadro de pessoal da unidade e shared services do grupo (alocados por chave) |
| FSE | Fornecimentos e serviços externos (parte fixa + variável com a ocupação) |
| Relatório | Exploração, resultados, dívida, saída e cash flows ano a ano do cenário ativo |
| Checks | Verificações automáticas de coerência |
| M_Base / M_Pessimista / M_Otimista | Motores de cálculo (não editar) |

Os valores por defeito são os do projeto **Cocoon**, usado como caso de teste: a receita e o GOP reproduzem o modelo original.

### Regenerar o ficheiro

```bash
python3 ferramentas/gerar_modelo.py Modelo_Avaliacao_Investimentos.xlsx
```
