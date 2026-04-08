function doGet() {

return HtmlService.createHtmlOutputFromFile('Index')
  .setTitle("NOME_TITULO_ARQUIVO")
  .setSandboxMode(HtmlService.SandboxMode.IFRAME)
  .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
  .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function processarFormulario(dados) {
  var planilha = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Captura a data e hora atual e formata no padrão brasileiro
  var dataHoraAtual = new Date();
  var dataFormatada = Utilities.formatDate(dataHoraAtual, "America/Sao_Paulo", "dd/MM/yyyy 'às' HH:mm");
  
  planilha.appendRow([
    dataHoraAtual, dados.proposta, dados.nomeContato, dados.emailContato, 
    dados.razaoContratante, dados.enderecoContratante, 
    dados.vendedorNome, dados.observacoes, JSON.stringify(dados.solicitantes)
  ]);

  var templateId = "TEMPLATE_ID_DOC";
  var arquivoCopia = DriveApp.getFileById(templateId).makeCopy("NOME_DO_ARQUIVO" + dados.proposta);
  var doc = DocumentApp.openById(arquivoCopia.getId());
  var body = doc.getBody();

  body.replaceText('{{PROPOSTA}}', dados.proposta);
  body.replaceText('{{NOME_CONTATO}}', dados.nomeContato);
  body.replaceText('{{EMAIL_CONTATO}}', dados.emailContato);
  body.replaceText('{{RAZAO_CONTRATANTE}}', dados.razaoContratante);
  body.replaceText('{{ENDERECO_CONTRATANTE}}', dados.enderecoContratante);
  body.replaceText('{{VENDEDOR_NOME}}', dados.vendedorNome);
  body.replaceText('{{OBSERVACOES}}', dados.observacoes || "Nenhuma observação.");
  
  // Substitui a tag da data/hora (Lembre de colocar {{DATA_HORA}} no seu Google Docs)
  body.replaceText('{{DATA_HORA}}', dataFormatada);

  // Lógica para desenhar os Blocos de Solicitantes e as Tabelas
  var placeholder = body.findText('{{BLOCOS_SOLICITANTES}}');
  if (placeholder) {
    var elementoPai = placeholder.getElement().getParent();
    var index = body.getChildIndex(elementoPai);

    dados.solicitantes.forEach(function(solicitante, i) {
      var titulo = body.insertParagraph(index++, "Solicitante " + (i+1) + " (Consumidor Final)");
      titulo.setHeading(DocumentApp.ParagraphHeading.HEADING3);
      body.insertParagraph(index++, "Razão Social: " + (solicitante.razao || "Mesmo que o Contratante"));
      body.insertParagraph(index++, "Endereço Completo: " + (solicitante.endereco || "Mesmo que o Contratante"));
      body.insertParagraph(index++, "");

      var table = body.insertTable(index++);
      var headerRow = table.appendTableRow();
      headerRow.appendTableCell("Capacidade").setBackgroundColor('#f4f4f9');
      headerRow.appendTableCell("Tag/Identificação").setBackgroundColor('#f4f4f9');
      headerRow.appendTableCell("Identificação Conjunto").setBackgroundColor('#f4f4f9');

      solicitante.pecas.forEach(function(peca) {
        var row = table.appendTableRow();
        row.appendTableCell(peca.capacidade);
        row.appendTableCell(peca.tag);
        row.appendTableCell(peca.conjunto);
      });

      body.insertParagraph(index++, "");
    });
    
    elementoPai.removeFromParent();
  }

  doc.saveAndClose();

  var pdf = arquivoCopia.getAs(MimeType.PDF);

  // --- 1. E-MAIL INTERNO PROFISSIONAL (Fixo + Vendedor) ---
  var emailFixoEmpresa = "EMAIL_FIXO_DA_EMPRESA";
  var emailVendedor = dados.vendedorEmail;
  var destinatariosInternos = emailFixoEmpresa + "," + emailVendedor;

  // Montagem do e-mail interno com visual de sistema
  var emailInternoHtml = 
    "<div style='font-family: Arial, sans-serif; color: #333; max-width: 600px; border: 1px solid #ddd; border-radius: 8px; overflow: hidden;'>" +
      "<div style='background-color: #1f497d; color: white; padding: 15px; text-align: center;'>" +
        "<h2 style='margin: 0;'>Formulário de Dados para Certificado(s)</h2>" +
      "</div>" +
      "<div style='padding: 20px;'>" +
        "<p>Olá Equipe,</p>" +
        "<p>Os dados para a <b>Proposta " + dados.proposta + "</b> foram preenchidos com sucesso pelo portal.</p>" +
        "<table style='width: 100%; border-collapse: collapse; margin-top: 15px; margin-bottom: 15px;'>" +
          "<tr><td style='padding: 8px; border-bottom: 1px solid #eee; width: 30%;'><b>Cliente:</b></td><td style='padding: 8px; border-bottom: 1px solid #eee;'>" + dados.razaoContratante + " (" + dados.nomeContato + ")</td></tr>" +
          "<tr><td style='padding: 8px; border-bottom: 1px solid #eee;'><b>Vendedor:</b></td><td style='padding: 8px; border-bottom: 1px solid #eee;'>" + dados.vendedorNome + "</td></tr>" +
          "<tr><td style='padding: 8px; border-bottom: 1px solid #eee;'><b>Preenchido em:</b></td><td style='padding: 8px; border-bottom: 1px solid #eee;'>" + dataFormatada + "</td></tr>" +
        "</table>" +
        "<p>O arquivo PDF com o detalhamento completo está em <b>anexo</b> neste e-mail para darmos andamento.</p>" +
      "</div>" +
    "</div>";

  MailApp.sendEmail({
    to: destinatariosInternos,
    subject: 'NOVA SOLICITAÇÃO - Proposta: ' + dados.proposta + ' (' + dados.razaoContratante + ')',
    htmlBody: emailInternoHtml,
    attachments: [pdf],
    name: "Portal WL Pesos Padrão"
  });

  arquivoCopia.setTrashed(true);

  // --- 2. E-MAIL DE CONFIRMAÇÃO PARA O CLIENTE ---
  var assuntoCliente = "Confirmação de Dados - Proposta: " + dados.proposta;
  var mensagemHtmlCliente = 
    "<div style='font-family: Arial, sans-serif; color: #333; max-width: 600px; margin: 0 auto; border: 1px solid #ddd; border-radius: 8px; overflow: hidden;'>" +
      "<div style='background-color: #c82333; color: white; padding: 20px; text-align: center;'>" +
        "<h2 style='margin: 0;'>Dados Recebidos com Sucesso!</h2>" +
      "</div>" +
      "<div style='padding: 20px;'>" +
        "<p>Olá, <b>" + dados.nomeContato + "</b>!</p>" +
        "<p>Gostaríamos de confirmar que recebemos as informações referentes à <b>Proposta " + dados.proposta + "</b> em " + dataFormatada + ".</p>" +
        "<p>Para sua conferência, enviamos em <b>anexo um PDF</b> com o resumo de todos os dados preenchidos.</p>" +
        "<p>Seu vendedor responsável (<b>" + dados.vendedorNome + "</b>) já foi notificado e dará andamento ao seu pedido e emissão do(s) certificado(s).</p>" +
        "<p>Qualquer dúvida, basta responder este e-mail ou nos chamar no WhatsApp (11) 3641-5974.</p>" +
        "<br><p>Atenciosamente,<br><b>Equipe WL Pesos Padrão</b><br><a href="site_da_empresa">"link_do_site"</a></p>" +
      "</div>" +
    "</div>";

  try {
    MailApp.sendEmail({
      to: dados.emailContato,
      subject: assuntoCliente,
      htmlBody: mensagemHtmlCliente,
      attachments: [pdf],
      name: "WL Pesos Padrão"
    });
  } catch(e) {
    Logger.log("Erro ao enviar e-mail: " + e);
  }

  return "Sucesso";
}
