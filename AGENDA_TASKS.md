# Agenda de tarefas da planilha

O arquivo `AgendaTasks.gs` adiciona o menu **Agenda** à planilha, com os comandos
**Atualizar agenda** e **Gerar PDF**.

## Intervalos nomeados

Crie estes intervalos na planilha:

| Nome | Conteúdo |
| --- | --- |
| `ID` | Uma única célula com o ID da lista do Google Tasks. |
| `MÊS` | Uma única célula com um número de `1` a `12`. O PDF sempre usa o ano atual. |
| `TAREFAS` | Um intervalo com pelo menos cinco colunas, no formato descrito abaixo. |

As colunas de `TAREFAS` são:

1. título da tarefa;
2. recorrência (`Semanal`, `Quinzenal`, `Mensal`, `Bimestral`, `Trimestral`,
   `Quadrimestral`, `Semestral` ou `Anual`);
3. data da primeira ocorrência;
4. primeira linha das observações e texto exibido depois do título no PDF;
5. terceira linha das observações, inclusive os links aplicados a partes do texto.

Uma linha somente é processada quando sua segunda coluna contém uma recorrência.
O script cria ocorrências individuais durante dois anos a partir da data inicial.
Isso é necessário porque a API do Google Tasks não oferece recorrência para tarefas
criadas pela integração.

## Instalação e autorização

1. Associe o projeto do Apps Script à planilha e copie os arquivos do projeto.
2. Confirme em **Serviços**, no editor do Apps Script, que a **Google Tasks API**
   está habilitada. O manifesto já declara o serviço avançado `Tasks` versão `v1`.
3. Reabra a planilha para executar `onOpen` e mostrar o menu **Agenda**.
4. Na primeira utilização, autorize o acesso à planilha, ao Google Tasks, ao Drive
   e ao Google Docs.

**Atualizar agenda** remove inclusive tarefas concluídas e ocultas da lista indicada
em `ID` antes de criar as novas ocorrências. Use uma lista exclusiva para esta
automação se houver tarefas que devam ser preservadas.

Os PDFs são armazenados na pasta **Calendários de tarefas**, criada na raiz do
Google Drive quando necessário. O documento temporário usado na conversão é enviado
à lixeira depois que o PDF é criado.
