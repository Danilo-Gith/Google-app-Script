function sendMail_prazo_vencido(rng) {
  
    var values=sheet.getDataRange().getValues();
 
   for (var row=0; row < values.length; row++){ 
 
   //var cabecalho = cabecalho + "<img src=\"https://upload.wikimedia.org/wikipedia/commons/c/c9/Stone_pagamentos.png\" width=\"200px\"><p></p>"  
 
   var corpo = '<p> Olá '+values[row][1]+values[row][2]+', tudo bem? Espero que sim!<p> Identificamos que você não realizou a devoluçao do crachá provisório. Por favor, solicitamos que devolva na recepção o quanto antes.<p>Obrigado!</p><p>Atenciosamente, <br><br><br> ' +imagem+ '<br><br><br><font color="#1da74b" face="Arial, sans-serif"><b>Segfís | Segurança da Informação</b></font><br><br><span style="background-color:transparent;color:rgb(112,173,71);font-family:Arial;font-size:8pt;white-space:pre-wrap">Security Team Office</span><br><span style="color:rgb(153,153,153);font-family:&quot;Segoe UI&quot;,sans-serif;font-size:10.6667px">Rua Gomes de Carvalho, 1609 - 8°andar | São Paulo - SP</span><br><font color="#757b80" face="Arial,sans-serif"><span style="color:rgb(117,123,128);font-family:Arial,sans-serif;font-size:9pt">Email: segfis@stone.com.br&nbsp;&nbsp;</span></p></p></p></p></p></p></p></br></br></br></br></br></br</br></br></br></br></br></br>'
 
     if(row > 0 && values[row][20] == "S"){
       mailApp.sendEmail({to: values[row][7],
         cc: 'teste@gmail.com',
         subject: '[LEMBRETE] Crachá Provisório  ',
         htmlBody: cabecalho + corpo     
         
       });
 
       }            
 
       }
            
     }
   