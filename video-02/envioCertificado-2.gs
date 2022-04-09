function carregarDados() {
 
  const respForm = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Respostas ao formulário 1');
  const rangeValues = respForm.getRange('A2:D').getValues()
  
  rangeValues.forEach((row, rowId)=>{
    //Verifica se existe registro -> row[0] com algum dado, e se não foi enviado row[4] 'Enviado em' sem data
    if(row[0] && !row[3]){
      const email = row[1];
      const nome = row[2];

      enviaEmail(email, nome);
      respForm.getRange('D' + (rowId+2)).setValue(new Date());
    }
  });
}

function enviaEmail(email, nome){
  const corpo = `  Olá ${nome} ,

  É com muita satisfação que agradecemos a sua participação no Workshop Aprendendo Google Script.

  O seu certificado está anexado aqui.

  Muito obrigado!
  
  Canal do Luiz`

  // ID do PDF que tem o certificado.
  const certificadoId = "1XDg5LfMqzhfS0UYXHyj_LFUy_1SnuZX2";
  // Enviar e-mail com certificado
  const certificado = geraCertificadoNominal(certificadoId, nome);

  GmailApp.sendEmail(email, `Certificado do Workshop Aprendendo Google Script`, corpo, {
      attachments: certificado.getAs(MimeType.PDF),
      name: 'Canal do Luiz'
  });
}


function geraCertificadoNominal(certificadoId, nome){
  
  const diretorio = "1LQ1LMYJMv9OJJKlWPjK4Le4LKeKOL4ZU";

  const certificadoOriginal = DriveApp.getFileById(certificadoId);
  const novoArquivo = certificadoOriginal.makeCopy(nome, DriveApp.getFolderById(diretorio));

  const certificado = SlidesApp.openById(novoArquivo.getId());

  //Colocando o nome no certificado
  // no slide 0 tem uma caixa(shape) escrito {nome}, vamos achar ela:
  certificado.getSlides()[0].getShapes().forEach(caixa =>{
    if(caixa.getText().asString().trim().toLowerCase() == "{nome}"){
      caixa.getText().setText(nome);
    }
  });

  //Vamos Salvar e retornar o arquivo de Slides
  certificado.saveAndClose();

  return  DriveApp.getFileById(certificado.getId());
}