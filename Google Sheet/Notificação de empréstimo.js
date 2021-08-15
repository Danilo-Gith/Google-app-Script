

//====================================================================================================================================
//Essa função encaminha o E-mail para coloborador notificando que ele retitou um crachá provisório por motivo de ESQUECIMENTO.
//====================================================================================================================================
function sendMail_esquecimento(rng) {

    var ui = SpreadsheetApp.getUi();
    
     var result = ui.alert(
        'Empréstimo de crachá provisório! ',
        'Prosseguir com notificação de empréstimo por ESQUECIMENTO?',
         ui.ButtonSet.YES_NO);
         
     if (result == ui.Button.YES) {
       
     } else {
     
     }
     
     var values=sheet.getDataRange().getValues(); 
   
     for (var row=0; row < values.length; row++){
   
       SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange('N2:N').clearContent();
       //Essa função é para apagar automaticamente a caixa de seleção que está  na coluna N. Você vai selecionar a caixa de seleção que corresponde a linha da pessoa que contem as informações para mandar o e-mail informando sobre o empréstimo e para solicitar o desblqueio do crachá quando o colaborador não devolver na data determinada. 
   
         var corpo ='<p>Olá '+values[row][1]+values[row][2]+', tudo bem? Espero que sim!<p><p>Estamos notificando que você retirou na data de hoje, '+ values[row][9]+' um crachá provisório de número '+values[row][0]+ ' no prédio da '+values[row][5]+' por motivo de <b>(ESQUECIMENTO). </b><p><p>Esse crachá tem funcionalidade de apenas 1 dia conforme informado no momento da retirada. Caso ele não seja entregue na data de devolução, o acesso será bloqueado por questões de segurança.<p><p>Obrigado!</p></p><p><p>Atenciosamente,</p></p><br> '+img+ '<br><br><font color="#1da74b" face="Arial, sans-serif"><b>Assinatura</b></font><br><br><span style="color:rgb(117,123,128);font-family:Arial,sans-serif;font-size:9pt">Departamento&nbsp;&nbsp;</span><br><font color="#757b80" face="Arial,sans-serif"><span style="font-size:12px">(11) 94532-4025</span></font><br>teste@com.br</br></br></br></br></br></br</br></br></br><p/></p></p></p></p>'
   
       if(row > 0 && values[row][17] == "S"){
         mailApp.sendEmail({to: values[row][7],
           cc: 'teste@gmail.com',
           subject: '[Lembrete] Crachá Provisório  ',
           htmlBody: cabecalho + corpo             
          
           
         });
      
   
         }            
        
         }
            
       }
   
     
   //======================================================================================================================
   //Essa função encaminha o E-mail para coloborador notificando que ele retitou um crachá provisório por motivo de PERDA.
   //=======================================================================================================================
   function sendMail_perda(rng) {
   
     var ui = SpreadsheetApp.getUi();
        
     var result = ui.alert(
        'Empréstimo de crachá provisório! ',
        'Prosseguir com notificação de empréstimo por PERDA?',
         ui.ButtonSet.YES_NO);
         
     if (result == ui.Button.YES) {
       
     } else {
   
     }
   
     var values=sheet.getDataRange().getValues(); 
   
     for (var row=0; row < values.length; row++){
   
       SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange('N2:N').clearContent();
       //Essa função é para apagar automaticamente a liha da  Coluna N onde a caixa de seleção está selecionada para mandar o e-mail informando sobre o empréstimo e para solicitar o desblqueio do crachá quando o colaborador não devolver na data determinada.
   
         var corpo ='<p>Olá '+values[row][1]+values[row][2]+', tudo bem? Espero que sim!<p><p>Estamos notificando que você retirou na data de hoje, '+ values[row][9]+' um crachá provisório de número '+values[row][0]+ ' no prédio da '+values[row][5]+' por motivo de <b>(PERDA). </b><p><p>Esse crachá tem funcionalidade de apenas 3 dias corridos conforme informado no momento da retirada.Caso ele não seja entregue na data de devolução, o acesso será bloqueado por questões de segurança.<p><p>Obrigado!</p></p><p><p>Atenciosamente,</p></p><br> '+img+ '<br><br><font color="#1da74b" face="Arial, sans-serif"><b>Assinatura</b></font><br><br><span style="color:rgb(117,123,128);font-family:Arial,sans-serif;font-size:9pt">Departamento&nbsp;&nbsp;</span><br><font color="#757b80" face="Arial,sans-serif"><span style="font-size:12px">(11) 94532-4025</span></font><br>teste@com.br</br></br></br></br></br></br</br></br></br><p/></p></p></p></p>'
   
       if(row > 0 && values[row][17] == "S"){
         mailApp.sendEmail({to: values[row][7],
            cc:'teste@gmail.com',
            subject: '[Lembrete] Crachá Provisório  ',
           htmlBody: cabecalho + corpo       
           
         });
   
         }            
   
         }
         // ui.alert('Notificação enviada com sucesso');    
       }
   
   //===========================================================================================================================
   //Essa função encaminha o E-mail para coloborador notificando que ele retitou um crachá provisório por motivo de CONFECCÇÃO.
   //===========================================================================================================================
   function sendMail_confeccao(rng) {
   
     var ui = SpreadsheetApp.getUi();
     //Essa variavel é a função que pergunta ao usuário como uma confirmação se ele realmente quer mandar o e-mail notificando o empréstimo do crachá
   
     var result = ui.alert(
        'Empréstimo de crachá provisório! ',
        'Prosseguir com notificação de empréstimo por CONFECÇÃO?',
         ui.ButtonSet.YES_NO);
   
         //SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange('Q2:Q').clearContent();
   
     if (result == ui.Button.YES) {
       
     } else {
   
     }
   
     var values=sheet.getDataRange().getValues(); 
   
     for (var row=0; row < values.length; row++){
   
       SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange('N2:N').clearContent();
       //Essa função é para apagar automaticamente a liha da  Coluna N onde a caixa de seleção está selecionada para mandar o e-mail informando sobre o empréstimo e para solicitar o desblqueio do crachá quando o colaborador não devolver na data determinada. 
   
         var corpo ='<p>Olá '+values[row][1]+values[row][2]+', tudo bem? Espero que sim!<p><p>Estamos notificando que você retirou na data de hoje, '+ values[row][9]+' um crachá provisório de número '+values[row][0]+ ' no prédio da '+values[row][5]+' por motivo que o seu está em fase de <b>(CONFECÇÃO). </b><p><p>Esse crachá tem funcionalidade de apenas 3 dias corridos conforme informado no momento da retirada.Caso ele não seja entregue na data de devolução, o acesso será bloqueado por questões de segurança.<p><p>Obrigado!</p></p><p><p>Atenciosamente,</p></p><br> '+img+ '<br><br><font color="#1da74b" face="Arial, sans-serif"><b>Assinatura</b></font><br><br><span style="color:rgb(117,123,128);font-family:Arial,sans-serif;font-size:9pt">Departamento&nbsp;&nbsp;</span><br><font color="#757b80" face="Arial,sans-serif"><span style="font-size:12px">(11) 94532-4025</span></font><br>teste@com.br</br></br></br></br></br></br</br></br></br><p/></p></p></p></p>'
   
       if(row > 0 && values[row][17] == "S"){
         mailApp.sendEmail({to: values[row][7],
           cc: 'teste@gmail.com',
           subject: '[Lembrete] Crachá Provisório  ',
           htmlBody: cabecalho + corpo       
           
         });
   
         }            
   
         }
              
       }