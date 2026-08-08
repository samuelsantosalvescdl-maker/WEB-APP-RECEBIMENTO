# Agenda do Google Calendar gerenciada pela planilha

O arquivo `AgendaCalendar.gs` adiciona o menu **Agenda** à planilha, com os comandos
**Atualizar agenda** e **Gerar PDF**.

## Intervalos nomeados

Crie estes intervalos na planilha:

| Nome | Conteúdo |
| --- | --- |
| `ID` | Uma única célula com o ID da agenda do Google Calendar. |
| `MÊS` | Uma única célula com um número de `1` a `12`. O PDF sempre usa o ano atual. |
| `TAREFAS` | Um intervalo com pelo menos cinco colunas, no formato descrito abaixo. |

As colunas de `TAREFAS` são:

1. título da tarefa;
2. recorrência (`Semanal`, `Quinzenal`, `Mensal`, `Bimestral`, `Trimestral`,
   `Quadrimestral`, `Semestral` ou `Anual`);
3. data da primeira ocorrência;
4. texto acrescentado ao título depois de ` - ` e exibido da mesma forma no PDF;
5. descrição do evento, inclusive os links aplicados a partes do texto.

Uma linha somente é processada quando sua segunda coluna contém uma recorrência.
O script cria eventos de dia inteiro individuais durante dois anos a partir da data
inicial. Quando a quarta coluna estiver preenchida, o título terá o formato
`Nome da tarefa - valor da quarta coluna`.

## Instalação e autorização

1. Associe o projeto do Apps Script à planilha e copie os arquivos do projeto.
2. Confirme em **Serviços**, no editor do Apps Script, que a **Google Calendar API**
   está habilitada. O manifesto já declara o serviço avançado `Calendar` versão `v3`.
3. Reabra a planilha para executar `onOpen` e mostrar o menu **Agenda**.
4. Na primeira utilização, autorize o acesso à planilha, ao Google Calendar, ao Drive
   e ao Google Docs.

**Atualizar agenda** remove todos os eventos da agenda indicada em `ID` antes de criar
as novas ocorrências. Use uma agenda exclusiva para esta automação se houver eventos
que devam ser preservados. O usuário que executar o comando precisa ter permissão
para alterar eventos nessa agenda; ele não precisa ser proprietário da planilha.

Os PDFs são armazenados na pasta **Calendários de tarefas**, criada na raiz do
Google Drive quando necessário. O documento temporário usado na conversão é enviado
à lixeira depois que o PDF é criado.
