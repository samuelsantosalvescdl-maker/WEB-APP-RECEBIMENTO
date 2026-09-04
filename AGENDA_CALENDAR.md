# Agenda do Google Calendar gerenciada pela planilha

O arquivo `AgendaCalendar.gs` adiciona o menu **Agenda** à planilha, com os comandos
**Atualizar agenda**, **Gerar PDF**, **Diagnosticar atualização** e **Cancelar
atualização travada**.

## Intervalos nomeados

Crie estes intervalos na planilha:

| Nome | Conteúdo |
| --- | --- |
| `ID` | Uma única célula com o ID da agenda do Google Calendar. |
| `MÊS` | Uma única célula com um número de `1` a `12`. O PDF sempre usa o ano atual. |
| `PDF_TEMPLATE_ID` | ID de uma apresentação Google Slides vazia configurada como A3 horizontal. |
| `TAREFAS` | Um intervalo com pelo menos cinco colunas, no formato descrito abaixo. |

As colunas de `TAREFAS` são:

1. título da tarefa;
2. recorrência (`Semanal`, `Quinzenal`, `Mensal`, `Bimestral`, `Trimestral`,
   `Quadrimestral`, `Semestral` ou `Anual`);
3. data da primeira ocorrência;
4. texto acrescentado ao título depois de ` - ` e exibido da mesma forma no PDF;
5. descrição do evento, inclusive os links aplicados a partes do texto.

Uma linha somente é processada quando contém recorrência, título e uma data inicial
válida. Linhas sem título ou sem data são ignoradas e não impedem a criação das
outras tarefas válidas. O script cria uma única série de evento de dia inteiro para
cada linha, contendo todas as datas dos dois anos a partir da data inicial. Quando a quarta coluna estiver preenchida, o
título terá o formato `Nome da tarefa - valor da quarta coluna`.

## Atualização em blocos

Para não exceder o tempo máximo do Apps Script, **Atualizar agenda** remove e cria
no máximo 10 séries por execução (ou trabalha por até um minuto). Antes
de iniciar cada bloco, o script deixa preparado um gatilho de segurança para continuar
automaticamente aproximadamente um minuto depois. Assim, mesmo uma interrupção
abrupta por limite de tempo não deixa o processo sem continuação.

O progresso é salvo depois de cada série criada. Cada linha recebe um ID determinístico
e todas as suas datas ficam em uma única série; portanto, repetir um bloco interrompido
não cria uma cópia duplicada e não exige centenas de chamadas à API.
A exclusão também tolera resultados antigos mantidos temporariamente pela API do
Calendar. Uma notificação aparece na planilha quando todo o processo termina.

A criação somente começa depois que a API retorna a agenda sem eventos ativos em três execuções
consecutivas e uma quarta consulta final também confirma que não existe nenhum evento ativo.
A consulta não usa limite de data: eventos passados, futuros, avulsos e séries
recorrentes são removidos. Essa confirmação adicional pode acrescentar alguns minutos
à primeira sincronização, mas impede que a inserção comece sobre uma agenda que ainda
contenha eventos antigos.

Registros com `status=cancelled` são tombstones de eventos ou exceções recorrentes
já removidas. Eles são percorridos pela paginação, mas não são removidos novamente e
não impedem a agenda de ser considerada visualmente vazia. A listagem usa páginas de
250 registros, separadas do limite de 10 exclusões reais por lote.

Todas as listagens de exclusão e verificação percorrem `nextPageToken`, inclusive
quando uma página intermediária não contém itens. Uma página vazia somente confirma o
fim da agenda quando também não existe uma próxima página.

Cada atualização recebe um `runId` novo. Os IDs dos eventos incluem esse `runId`, a
planilha e a linha de origem: uma retomada da mesma atualização reutiliza o ID com
segurança, mas uma atualização futura não reutiliza IDs de eventos apagados. Um conflito
409 somente é aceito depois de consultar o evento e validar sua origem, execução,
conteúdo e recorrência.

O progresso é compartilhado pelo projeto e protegido por um bloqueio. Enquanto uma
atualização estiver ativa, outro clique ou outro usuário não poderá iniciar uma
segunda atualização concorrente. Isso evita que um processo apague os eventos que o
outro acabou de criar. Aguarde a mensagem de conclusão antes de iniciar novamente.

O identificador recebido no callback do gatilho é usado apenas para diagnóstico. A
continuação é controlada pelo bloqueio e pelo estado ativo, evitando que diferenças
entre representações numéricas e textuais de `triggerUid` interrompam um lote válido.
O estado diferencia `lastHeartbeatAt` (um callback executou) de `lastProgressAt`
(houve exclusão real, confirmação vazia, transição, criação ou conclusão). Agendar
um gatilho não renova o progresso. Um estado ativo sem seu gatilho correspondente, ou
sem progresso real por mais de 15 minutos, é considerado
órfão e não bloqueia uma nova sincronização.

Se for necessário interromper administrativamente um fluxo, use **Agenda > Cancelar
atualização travada**. Essa opção cancela os gatilhos e marca o estado como cancelado,
mas não exclui nem cria eventos. Depois dela, **Atualizar agenda** pode ser iniciado
novamente.

Use **Agenda > Diagnosticar atualização** para consultar, sem alterar a agenda, o
estado persistido, o gatilho correspondente, gatilhos residuais e as contagens de
eventos ativos, cancelados, instâncias e mestres recorrentes. Um resumo aparece em um
toast e os detalhes, inclusive uma amostra limitada de eventos, ficam nos logs de
execução do Apps Script.

Não altere, insira ou remova linhas de `TAREFAS` enquanto uma atualização estiver
em andamento. Se isso acontecer, o processo é interrompido para evitar uma agenda
parcial; depois das alterações, execute **Atualizar agenda** novamente.

## Instalação e autorização

1. Associe o projeto do Apps Script à planilha e copie os arquivos do projeto.
2. Confirme em **Serviços**, no editor do Apps Script, que a **Google Calendar API**
   está habilitada. O manifesto já declara o serviço avançado `Calendar` versão `v3`.
3. Reabra a planilha para executar `onOpen` e mostrar o menu **Agenda**.
4. Na primeira utilização, autorize o acesso à planilha, ao Google Calendar, ao Drive,
   ao Google Docs e à interface da planilha. O acesso à interface permite que o
   comando **Gerar PDF** mostre a janela com o link para download.

Se o script já tiver sido autorizado antes da inclusão dessas permissões, execute
`gerarPDF` uma vez diretamente pelo editor do Apps Script e aceite a nova solicitação
de autorização. Depois disso, o comando pode ser usado normalmente pelo menu da
planilha.

**Atualizar agenda** remove todos os eventos da agenda indicada em `ID` antes de criar
as novas ocorrências. Use uma agenda exclusiva para esta automação se houver eventos
que devam ser preservados. O usuário que executar o comando precisa ter permissão
para alterar eventos nessa agenda; ele não precisa ser proprietário da planilha.

Os PDFs são armazenados na pasta **Calendários de tarefas**, criada na raiz do
Google Drive quando necessário. Configure uma apresentação vazia com página A3
horizontal (420 × 297 mm), deixe pelo menos um slide e informe o ID do arquivo no
intervalo nomeado `PDF_TEMPLATE_ID`. O script copia esse template, mantém e limpa
exatamente um slide e nunca modifica o arquivo original. Por isso, o PDF resultante
possui exatamente uma página. A cópia temporária usada na conversão é enviada à
lixeira depois que o PDF é criado, inclusive se a renderização ou a conversão falhar.

No PDF, a quarta coluna de `TAREFAS` é tratada como responsável. Cada responsável
recebe uma cor consistente no realce claro aplicado atrás do texto das tarefas e da
legenda. Linhas sem responsável usam o marcador neutro **SEM RESPONSÁVEL**. As tarefas usam fonte fixa de
8 pontos e uma única linha: títulos longos são abreviados visualmente com reticências,
sem alterar os dados originais. Cada dia possui uma única coluna com no máximo dez
linhas visuais. Quando houver mais de dez tarefas, são mostradas as nove primeiras e
a última linha informa `+N tarefas`. O template é rejeitado com um erro detalhado se
não tiver altura suficiente para as dez linhas, sem reduzir automaticamente a fonte.
As cores são derivadas do nome normalizado do responsável, com resolução de colisões
entre as pessoas presentes no mesmo PDF.

Durante a geração, uma amostra limitada registra nos logs a linha de origem, o título
original, o título visualmente truncado, os comprimentos e o responsável. Títulos que
começam com `SLIMPEZA` são registrados explicitamente como contendo o caractere extra
na própria fonte de dados; o gerador não faz correções ortográficas silenciosas.
