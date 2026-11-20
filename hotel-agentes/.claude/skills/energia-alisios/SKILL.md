---
name: energia-alisios
description: Atualiza o painel "Energia · Hotel Alísios" com o relatório mais recente da Bluenergy e assinala desvios de consumo. Usar quando chegar um relatório Bluenergy ou quando o gestor pedir o ponto de situação da energia.
---

# Agente de energia do Hotel Alísios

## O que faz
1. Procura no Outlook os emails de `@blue-energy.pt` com relatórios novos (assuntos "Relatório do consumo energético", "Relatório de Métricas", "Relatório … Produção Solar", "Alerta de Consumos").
2. Lê os PDFs anexos e tira:
   - consumo por fonte (bombas de calor, lavandaria, chiller, receção, cozinhas, restaurante, osmose, piscina, outros, total) do mês e do acumulado, 2026 contra 2025;
   - ocupação do mês;
   - consumo, custo (€ sem IVA) e quartos ocupados por mês (relatório de métricas, trimestral);
   - produção solar.
3. Lê no SharePoint a versão mais recente do orçamento "Classe 6" do Alísios (contas 6241 eletricidade, 6242 gás, 6243 água): real do ano e orçamento por mês. Não existe orçamento de 2026 por especialidade; o painel compara com o orçamento de 2027.
4. Atualiza os ficheiros de dados do painel (mensal, fontes, solar, custos, acoes) e confirma que a soma das fontes dá o total do relatório.
5. Assinala como desvio:
   - uma fonte que sobe mais de 10% face ao homólogo quando a ocupação sobe menos do que isso;
   - consumo acima do esperado em meses de ocupação baixa (menos de 50%);
   - produção solar a descer com radiação igual;
   - custo do mês acima do orçamento.
6. Atualiza a lista "O que está por fazer" com o que os emails mostrarem (proposta aceite, manutenção feita, etc.).
7. Se houver desvios, prepara um **rascunho** de email para a manutenção, no tom de `modelos/tom-de-voz.md`. **Nunca envia.**

## Painel
- Endereço: https://claude.ai/artifact/KicPAETi8cTzh7XqR3wBXW (privado ao gestor).
- Os números ficam no painel, não neste repositório (regra 5: dados financeiros privados).

## Regras
- Só leitura no email; rascunhos apenas.
- Se um número do relatório não bater com a soma das fontes, avisa o gestor em vez de corrigir.
- Indicar sempre de que relatório e data vem cada número.
