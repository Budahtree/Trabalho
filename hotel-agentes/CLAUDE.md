# Projeto: Agentes de IA para gestão do hotel

## Quem sou eu
- Gestor/a de hotel com experiência em operações e gestão. **Não sei programar.**
- Explica tudo em **português de Portugal**, de forma simples, com instruções concretas sempre que eu tiver de fazer alguma coisa.
- Uso **Windows**. Dá-me sempre os passos exatos (onde clicar, o que escrever).

## Objetivo
Criar um conjunto de "agentes" (skills + tarefas agendadas + ligações às ferramentas do hotel) que façam o trabalho repetitivo de gestão antes de eu começar o dia, e um **painel** com a operação do hotel atualizada ao longo do dia.

Inspiração: um post de um hoteleiro com 21 agentes (triagem de email, rascunhos de resposta, horas de chegada, concierge, preços dinâmicos, upgrades, plano de limpezas, lista de compras F&B, arquivo de faturas, relatórios, CRM). Não é para copiar os 21: é para adaptar ao meu hotel, começando pelos que poupam mais tempo.

## Regras que nunca mudam
1. **Nenhum email é enviado sem a minha aprovação.** Os agentes só criam rascunhos.
2. **Preços, restrições (MLOS) e upgrades são propostas.** Eu aprovo antes de qualquer alteração.
3. **Não é um sistema de vigilância da equipa.** Organiza tarefas, não controla pessoas.
4. "Tempo real" = atualização periódica (por exemplo, de hora a hora), não ao segundo.
5. Dados de hóspedes e financeiros ficam privados. Nada de credenciais escritas em ficheiros do projeto.

## Arquitetura prevista (Claude Code)
```
hotel-agentes/
├── CLAUDE.md                 ← este ficheiro
├── .claude/
│   ├── skills/               ← um agente por pasta, cada um com SKILL.md
│   │   ├── morning-routine/
│   │   ├── email-triage/
│   │   ├── email-draft/
│   │   ├── booking-inquiry/
│   │   ├── arrival-time-check/
│   │   ├── cleaning-schedule/
│   │   ├── dynamic-pricing/
│   │   └── finance-filing/
│   └── agents/               ← subagentes, se forem precisos
├── dados/                    ← exportações do PMS (CSV/Excel) quando não houver API
├── modelos/                  ← tom de voz, respostas-tipo, contactos de confiança (concierge)
└── painel/                   ← painel da operação (ocupação, chegadas, limpezas, tarefas)
```

## Agentes por área (lista de partida)
| Área | Agente | O que faz |
|---|---|---|
| Rotina | morning-routine | Corre os agentes da manhã e entrega um resumo do dia |
| Email | email-triage | Classifica e prioriza as caixas de email |
| Email | email-draft / booking-inquiry | Rascunhos no meu tom; respostas calorosas a pedidos de reserva |
| Email | arrival-time-check | Lista hóspedes sem hora de chegada e prepara o email |
| Hóspedes | concierge | Sugestões de restaurantes, tours e transfers (só contactos aprovados por mim) |
| Hóspedes | pre-arrival-outreach | Mensagem de boas-vindas para reservas premium |
| Receita | dynamic-pricing | Propõe tarifas e MLOS para os próximos 30 dias |
| Receita | room-upgrade | Identifica upgrades possíveis; só avança com aprovação |
| Operação | cleaning-schedule | Plano diário de limpezas a partir das reservas |
| Operação | fb-shopping-list | Lista de compras a partir das contagens de hóspedes e stock |
| Finanças | finance-filing | Lê faturas, arquiva no Drive/SharePoint e regista numa folha |
| Finanças | relatórios | Relatório semanal de ocupação, receita e marketing |
| Vendas | crm-pipeline | Acompanha pedidos do primeiro contacto ao pós-estadia |

## Fases
1. **Fundação:** ligar email e PMS; definir o meu tom de voz e respostas-tipo.
2. **Primeiros agentes (2–3):** os das áreas que me tiram mais tempo.
3. **Rotina da manhã + painel.**
4. **Receita e finanças** (preços, upgrades, faturas, relatórios).
5. Ligação opcional à minha app de P&L/Budget/Forecast (projeto separado).

## Decisões tomadas (13/11/2026)
- **PMS:** Newhotel. Começamos com exportações (Excel/CSV) para a pasta `dados/`; ver mais tarde se há acesso a API.
- **Email:** Outlook / Microsoft 365 (conector já ligado e a funcionar: leitura, rascunhos, OneDrive/SharePoint, Teams).
- **Áreas prioritárias:** Email, Operação, Finanças.
- **Dimensão:** mais de 60 quartos, equipa grande com vários departamentos.
- Os dados exportados do PMS **não** vão para o GitHub (ver `.gitignore`).

## Como trabalhar comigo
- Antes de construir algo grande, mostra-me o plano em poucas linhas e pergunta.
- Constrói um agente de cada vez, testa com dados reais e só depois passa ao seguinte.
- No fim de cada passo, diz-me em 2–3 frases o que ficou feito e o que preciso de fazer.
