
//Essas são as variaves globais que usamos para inicar uma lógica no google planilhas
var mailApp=MailApp;
var app=SpreadsheetApp;
var spreadsheet=app.getActiveSpreadsheet();
var sheet=spreadsheet.getSheetByName("nome da aba de sua planilha");

var cabecalho = "<img src=\"https://g-workplace.com/application/files/3716/0829/6797/google-apps-script_logo-removebg-preview_1.png\" width=\"200px\"><p></p>"
//Essa varialvel foi criada pra colocar uma url de uma imagem para ir no corpo do e-mail acima o texto. Você pode colocar logo que representa sua empresa ou outra imagem que queira inserir.


var img ="https://michigan.it.umich.edu/news/wp-content/uploads/2019/02/google-apps-script-1.png"
//Essa varialvel foi criada para inserir uma imagem na assinatuta do e-mail

var imagem ="https://michigan.it.umich.edu/news/wp-content/uploads/2019/02/google-apps-script-1.png>"
//Essa varialvel foi criada para inserir uma imagem na assinatuta do e-mail



//Essa função ela cria uma menu - interativo na barra de tarefa do google sheet, através dela vc pode declarar funções para serem executadas. Nesse script sera usado para enviar e-mail.
function onOpen()
{
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('CRACHÁ PROVISÓRIO')
  .addSubMenu(ui.createMenu('Notificação de Empréstimo')
  .addItem("Enviar E-mail crachá provisório esquecimento ","sendMail_esquecimento")
  .addItem("Enviar E-mail crachá provisório perda ","sendMail_perda")
  .addItem("Enviar E-mail crachá provisório confecção","sendMail_confeccao"))

  //.addSubMenu(ui.createMenu('Alerta de Devolução')
  //.addItem("Enviar E-mail de alarta para devolução do Crachá antes do vencimento","sendMail_vence_hoje"))
  
  //.addSubMenu(ui.createMenu('Solicitação de Devolução')
  //.addItem("Enviar E-mail solicitando devolução do crachá","sendMail_prazo_vencido"))

  //.addSubMenu(ui.createMenu('Solicitação de bloqueio')
  //.addItem("Enviar E-mail para o time de Segfis realizar o bloqueio do crachá vencido ","sendMail_bloqueio_de_acesso"))

  .addSubMenu(ui.createMenu('Solicitação de desbloqueio')
  .addItem("Enviar E-mail para o time de Segfis realizar o desbloqueio ","sendMail_desbloqueio_de_acesso"))

  //.addSubMenu(ui.createMenu('Código de Rastreio')
  //.addItem("Enviar E-mail com o código de rastreio ","sendMailCorrreio"))

  .addToUi();       
}