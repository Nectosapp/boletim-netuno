/**
 * Google Apps Script — Confirmação de Descadastramento do Boletim Grupo Netuno
 *
 * COMO IMPLANTAR:
 * 1. Acesse https://script.google.com → Novo Projeto
 * 2. Cole este código no editor
 * 3. Clique em "Implantar" → "Nova implantação"
 * 4. Tipo: "App da Web"
 * 5. Executar como: "Eu" (sua conta)
 * 6. Quem tem acesso: "Qualquer pessoa"
 * 7. Copie a URL gerada e cole na variável APPS_SCRIPT_URL do arquivo unsubscribe.html
 */

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var email = data.email;

    if (!email || !email.match(/^[^\s@]+@[^\s@]+\.[^\s@]+$/)) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false, message: "Email inválido"
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // Enviar email de confirmação ao destinatário
    var subject = "Confirmação de descadastramento — Boletim Grupo Netuno";
    var htmlBody = getConfirmationEmailHtml(email);

    GmailApp.sendEmail(email, subject, "", {
      htmlBody: htmlBody,
      name: "Grupo Netuno | R2F Capital",
      replyTo: "gustavoportugalhamer@gmail.com"
    });

    // Notificar o admin
    GmailApp.sendEmail("gustavoportugalhamer@gmail.com",
      "[Boletim] Descadastramento: " + email,
      "O email " + email + " solicitou descadastramento do Boletim Diário de Mercado.\n\n" +
      "Data: " + new Date().toLocaleString("pt-BR") + "\n\n" +
      "Ação necessária: adicionar este email ao arquivo unsubscribed.txt no repositório."
    );

    return ContentService.createTextOutput(JSON.stringify({
      success: true, message: "Email de confirmação enviado"
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false, message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    status: "ok", service: "Boletim Netuno Unsubscribe"
  })).setMimeType(ContentService.MimeType.JSON);
}

function getConfirmationEmailHtml(email) {
  return '<!DOCTYPE html>' +
  '<html lang="pt-BR"><head><meta charset="UTF-8"></head>' +
  '<body style="margin:0; padding:0; background-color:#e5e7eb;">' +
  '<table width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color:#e5e7eb;">' +
  '<tr><td align="center" style="padding:30px 10px;">' +
  '<table width="560" cellpadding="0" cellspacing="0" border="0" style="background-color:#ffffff; border-radius:12px; overflow:hidden; box-shadow:0 2px 12px rgba(0,0,0,0.08);">' +

  // Header
  '<tr><td style="background-color:#1a2744; padding:30px; text-align:center;">' +
  '<table cellpadding="0" cellspacing="0" border="0" style="margin:0 auto;"><tr>' +
  '<td style="background-color:#00acad; border-radius:8px; padding:10px 12px;">' +
  '<span style="font-family:Arial,sans-serif; font-size:20px; font-weight:700; color:#ffffff;">N</span></td>' +
  '<td style="padding-left:12px;">' +
  '<span style="font-family:Arial,sans-serif; font-size:18px; font-weight:700; color:#ffffff; letter-spacing:0.5px;">GRUPO NETUNO</span></td>' +
  '</tr></table></td></tr>' +

  // Body
  '<tr><td style="padding:35px 30px;">' +
  '<table width="100%" cellpadding="0" cellspacing="0" border="0">' +

  // Check icon
  '<tr><td align="center" style="padding-bottom:20px;">' +
  '<table cellpadding="0" cellspacing="0" border="0"><tr>' +
  '<td style="width:56px; height:56px; border-radius:50%; background-color:rgba(0,172,173,0.1); border:2px solid #00acad; text-align:center; vertical-align:middle; font-size:24px; color:#00acad;">&#10003;</td>' +
  '</tr></table></td></tr>' +

  // Title
  '<tr><td style="text-align:center; padding-bottom:16px;">' +
  '<p style="margin:0; font-family:Arial,sans-serif; font-size:20px; font-weight:700; color:#1a2744;">Inscrição cancelada com sucesso</p>' +
  '</td></tr>' +

  // Message
  '<tr><td style="text-align:center; padding-bottom:20px;">' +
  '<p style="margin:0; font-family:Arial,sans-serif; font-size:14px; color:#555555; line-height:1.7;">' +
  'Confirmamos que o e-mail abaixo foi removido da nossa lista de distribuição do <strong>Boletim Diário de Mercado</strong> do Grupo Netuno.</p>' +
  '</td></tr>' +

  // Email box
  '<tr><td align="center" style="padding-bottom:24px;">' +
  '<table cellpadding="0" cellspacing="0" border="0"><tr>' +
  '<td style="background-color:rgba(0,172,173,0.08); border:1px solid rgba(0,172,173,0.25); border-radius:6px; padding:10px 20px;">' +
  '<span style="font-family:Arial,sans-serif; font-size:13px; color:#00acad; font-weight:600;">' + email + '</span>' +
  '</td></tr></table></td></tr>' +

  // Info text
  '<tr><td style="text-align:center; padding-bottom:10px;">' +
  '<p style="margin:0; font-family:Arial,sans-serif; font-size:13px; color:#888888; line-height:1.6;">' +
  'A partir de agora, você não receberá mais nossos comunicados por e-mail.<br>' +
  'Caso mude de ideia, entre em contato conosco a qualquer momento.</p>' +
  '</td></tr>' +

  '</table></td></tr>' +

  // Divider
  '<tr><td style="padding:0 30px;"><table width="100%" cellpadding="0" cellspacing="0" border="0"><tr>' +
  '<td style="border-top:1px solid #eeeeee;"></td></tr></table></td></tr>' +

  // Footer
  '<tr><td style="padding:20px 30px; text-align:center;">' +
  '<p style="margin:0 0 6px 0; font-family:Arial,sans-serif; font-size:12px; color:#aaaaaa; line-height:1.5;">' +
  'Caso deseje voltar a receber o boletim, escreva para<br>' +
  '<a href="mailto:gustavoportugalhamer@gmail.com" style="color:#00acad; text-decoration:none;">gustavoportugalhamer@gmail.com</a></p>' +
  '<p style="margin:8px 0 0 0; font-family:Arial,sans-serif; font-size:11px; color:#cccccc;">' +
  'Grupo Netuno | <a href="https://r2fcapital.com.br" style="color:#00acad; text-decoration:none;">R2F Capital</a></p>' +
  '</td></tr>' +

  '</table></td></tr></table></body></html>';
}
