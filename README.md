# HistoricoBase
HistoricoBase é uma aplicação em C# feita para atualizar o histórico dos lançamentos contábeis das bases no SAP B1. Ela consiste em conectar e executar uma Query no banco Hana via DIAPI, que retorna os lançamentos e suas devidas novas observações, armazena esses valores numa lista e atualiza-os.

O objetivo foi poupar um trabalho que anteriormente era feito via DTW e de forma manual, feito a cada 30 dias, o que demorava muito pois a quantidade de lançamentos se tornava maior e de base em base. Nessa aplicação uma Tarefa agendada foi configurada para executar o .exe todos os dias, diminuindo a quantidade de LC's e não deixando-os ficar obsoletos.

# Tecnologias Utilizadas
## C#
## SQL Hana
