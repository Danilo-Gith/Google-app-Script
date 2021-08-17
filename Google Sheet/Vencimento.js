function sendMail_prazo_vencido(rng) {
  
    var values=sheet.getDataRange().getValues();
 
   for (var row=0; row < values.length; row++){ 
 
   //var cabecalho = cabecalho + "<img src=\"https://michigan.it.umich.edu/news/wp-content/uploads/2019/02/google-apps-script-1.png/" width=\"200px\"><p></p>"  
 
   var corpo = '<p> Olá '+values[row][1]+values[row][2]+', tudo bem? Espero que sim!<p> Estamos notificando o prazo de vencimento.<p>Obrigado!</p><p>Atenciosamente, <br><br><br> ' +imagem+ '<br><br><font color="#1da74b" face="Arial, sans-serif"><b>Assinatura</b></font><br><br><span style="color:rgb(117,123,128);font-family:Arial,sans-serif;font-size:9pt">Departamento&nbsp;&nbsp;</span><br><font color="#757b80" face="Arial,sans-serif"><span style="font-size:12px">(11) 94532-4025</span></font><br>teste@com.br</br></br></br></br></br></br</br></br></br><p/></p></p></p></p>'
 
     if(row > 0 && values[row][20] == "S"){
       mailApp.sendEmail({to: values[row][7],
         cc: 'teste@gmail.com',
         subject: '[LEMBRETE] Vencimento  ',
         htmlBody: cabecalho + corpo     
         
       });
 
       }            
 
       }
            
     }
   