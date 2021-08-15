//=========================================================================================================
//Essa função encaminha um  E-mail de alerta para coloborador no dia que ele precisa devolver o crachá
//=========================================================================================================

//Esse alerta de devolução vc pode crirar utilizando os acionadores 

function sendMail_vence_hoje(rng) {

    var values=sheet.getDataRange().getValues(); 
 
   for (var row=0; row < values.length; row++){
 
       var corpo = '<p>Olá '+values[row][1]+values[row][2]+', tudo bem? Espero que sim!<p><p>Passando para lembrar que o prazo de devolução do crachá provisório vence hoje. <p><p> Pedimos que o devolva até o final do dia na recepção. <p><p>Obrigado! <p><p>Atenciosamente, <br><br><br> ' +imagem+ '<br><br><font color="#1da74b" face="Arial, sans-serif"><b>Assinatura</b></font><br><br><span style="color:rgb(117,123,128);font-family:Arial,sans-serif;font-size:9pt">Departamento&nbsp;&nbsp;</span><br><font color="#757b80" face="Arial,sans-serif"><span style="font-size:12px">(11) 94532-4025</span></font><br>teste@com.br</br></br></br></br></br></br</br></br></br><p/></p></p></p></p>'      
 
     if(row > 0 && values[row][19] == "H"){
       mailApp.sendEmail({to: values[row][7],
       //cc: 'teste@gmail.com',
       subject: '[Lembrete] Crachá Provisório  ',
         htmlBody: cabecalho + corpo       
         
       });
 
       }            
 
       }
            
     }