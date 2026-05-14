/**
 * Google Apps Script — Descadastramento automático do Boletim Grupo Netuno
 *
 * Ao receber um POST com {email: "..."}, este script:
 * 1. Envia email de confirmação ao destinatário
 * 2. Notifica o admin
 * 3. Adiciona o email ao unsubscribed.txt no GitHub (automático)
 */

// ===== CONFIG =====
var GITHUB_TOKEN = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
var GITHUB_REPO = "Nectosapp/boletim-netuno";
var UNSUB_FILE_PATH = "unsubscribed.txt";
var ADMIN_EMAIL = "gustavoportugalhamer@gmail.com";

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var email = (data.email || "").trim().toLowerCase();

    if (!email || !email.match(/^[^\s@]+@[^\s@]+\.[^\s@]+$/)) {
      return _jsonResponse(false, "Email inválido");
    }

    // 1. Adicionar ao unsubscribed.txt no GitHub
    var added = addToUnsubscribedFile(email);

    // 2. Enviar email de confirmação ao destinatário
    GmailApp.sendEmail(email,
      "Confirmação de descadastramento — Boletim Grupo Netuno",
      "", {
        htmlBody: getConfirmationEmailHtml(email),
        name: "Grupo Netuno | R2F Capital",
        replyTo: ADMIN_EMAIL
      }
    );

    // 3. Notificar o admin
    GmailApp.sendEmail(ADMIN_EMAIL,
      "[Boletim] Descadastramento: " + email,
      "O email " + email + " solicitou descadastramento do Boletim Diário de Mercado.\n\n" +
      "Data: " + new Date().toLocaleString("pt-BR") + "\n\n" +
      (added
        ? "O email foi adicionado automaticamente ao unsubscribed.txt no GitHub."
        : "O email já estava no unsubscribed.txt (nenhuma ação necessária).")
    );

    return _jsonResponse(true, "Descadastramento realizado com sucesso");

  } catch (err) {
    // Tentar notificar admin sobre o erro
    try {
      GmailApp.sendEmail(ADMIN_EMAIL,
        "[Boletim] ERRO no descadastramento",
        "Erro ao processar descadastramento:\n" + err.toString() + "\n\nDados: " + (e.postData ? e.postData.contents : "sem dados")
      );
    } catch(e2) {}
    return _jsonResponse(false, err.toString());
  }
}

function doGet(e) {
  return _jsonResponse(true, "Boletim Netuno Unsubscribe API - OK");
}

// ===== GITHUB: adicionar email ao unsubscribed.txt =====
function addToUnsubscribedFile(email) {
  var url = "https://api.github.com/repos/" + GITHUB_REPO + "/contents/" + UNSUB_FILE_PATH;

  // 1. Ler conteúdo atual do arquivo
  var getResponse = UrlFetchApp.fetch(url, {
    method: "GET",
    headers: {
      "Authorization": "token " + GITHUB_TOKEN,
      "Accept": "application/vnd.github.v3+json"
    },
    muteHttpExceptions: true
  });

  var fileData = JSON.parse(getResponse.getContentText());
  var currentContent = Utilities.newBlob(Utilities.base64Decode(fileData.content)).getDataAsString();
  var sha = fileData.sha;

  // 2. Verificar se email já está na lista
  var lines = currentContent.split("\n").map(function(l) { return l.trim().toLowerCase(); });
  if (lines.indexOf(email) >= 0) {
    return false; // já existe
  }

  // 3. Adicionar email ao final
  var newContent = currentContent.trimRight() + "\n" + email + "\n";
  var encoded = Utilities.base64Encode(newContent, Utilities.Charset.UTF_8);

  // 4. Commit no GitHub
  var putResponse = UrlFetchApp.fetch(url, {
    method: "PUT",
    headers: {
      "Authorization": "token " + GITHUB_TOKEN,
      "Accept": "application/vnd.github.v3+json"
    },
    contentType: "application/json",
    payload: JSON.stringify({
      message: "unsubscribe: " + email,
      content: encoded,
      sha: sha
    }),
    muteHttpExceptions: true
  });

  var status = putResponse.getResponseCode();
  if (status !== 200 && status !== 201) {
    throw new Error("GitHub API error (" + status + "): " + putResponse.getContentText());
  }

  return true;
}

// ===== Helper =====
function _jsonResponse(success, message) {
  return ContentService.createTextOutput(JSON.stringify({
    success: success, message: message
  })).setMimeType(ContentService.MimeType.JSON);
}

// ===== Função de teste (executar manualmente para autorizar permissões) =====
function testAuth() {
  Logger.log("Token GitHub configurado: " + (GITHUB_TOKEN ? "SIM" : "NÃO"));
  Logger.log("Teste de permissões OK - GmailApp e UrlFetchApp autorizados.");
}

// ===== Email HTML de confirmação =====
function getConfirmationEmailHtml(email) {
  return '<!DOCTYPE html>' +
  '<html lang="pt-BR"><head><meta charset="UTF-8"></head>' +
  '<body style="margin:0; padding:0; background-color:#e5e7eb;">' +
  '<table width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color:#e5e7eb;">' +
  '<tr><td align="center" style="padding:30px 10px;">' +
  '<table width="560" cellpadding="0" cellspacing="0" border="0" style="background-color:#ffffff; border-radius:12px; overflow:hidden; box-shadow:0 2px 12px rgba(0,0,0,0.08);">' +
  '<tr><td style="background-color:#1a2744; padding:30px; text-align:center;">' +
  '<table cellpadding="0" cellspacing="0" border="0" style="margin:0 auto;"><tr>' +
  '<td style="background-color:#00acad; border-radius:8px; padding:10px 12px;">' +
  '<span style="font-family:Arial,sans-serif; font-size:20px; font-weight:700; color:#ffffff;">N</span></td>' +
  '<td style="padding-left:12px;">' +
  '<span style="font-family:Arial,sans-serif; font-size:18px; font-weight:700; color:#ffffff; letter-spacing:0.5px;">GRUPO NETUNO</span></td>' +
  '</tr></table></td></tr>' +
  '<tr><td style="padding:35px 30px;">' +
  '<table width="100%" cellpadding="0" cellspacing="0" border="0">' +
  '<tr><td align="center" style="padding-bottom:20px;">' +
  '<table cellpadding="0" cellspacing="0" border="0"><tr>' +
  '<td style="width:56px; height:56px; border-radius:50%; background-color:rgba(0,172,173,0.1); border:2px solid #00acad; text-align:center; vertical-align:middle; font-size:24px; color:#00acad;">&#10003;</td>' +
  '</tr></table></td></tr>' +
  '<tr><td style="text-align:center; padding-bottom:16px;">' +
  '<p style="margin:0; font-family:Arial,sans-serif; font-size:20px; font-weight:700; color:#1a2744;">Inscrição cancelada com sucesso</p>' +
  '</td></tr>' +
  '<tr><td style="text-align:center; padding-bottom:20px;">' +
  '<p style="margin:0; font-family:Arial,sans-serif; font-size:14px; color:#555555; line-height:1.7;">' +
  'Confirmamos que o e-mail abaixo foi removido da nossa lista de distribuição do <strong>Boletim Diário de Mercado</strong> do Grupo Netuno.</p>' +
  '</td></tr>' +
  '<tr><td align="center" style="padding-bottom:24px;">' +
  '<table cellpadding="0" cellspacing="0" border="0"><tr>' +
  '<td style="background-color:rgba(0,172,173,0.08); border:1px solid rgba(0,172,173,0.25); border-radius:6px; padding:10px 20px;">' +
  '<span style="font-family:Arial,sans-serif; font-size:13px; color:#00acad; font-weight:600;">' + email + '</span>' +
  '</td></tr></table></td></tr>' +
  '<tr><td style="text-align:center; padding-bottom:10px;">' +
  '<p style="margin:0; font-family:Arial,sans-serif; font-size:13px; color:#888888; line-height:1.6;">' +
  'A partir de agora, você não receberá mais nossos comunicados por e-mail.<br>' +
  'Caso mude de ideia, entre em contato conosco a qualquer momento.</p>' +
  '</td></tr>' +
  '</table></td></tr>' +
  '<tr><td style="padding:0 30px;"><table width="100%" cellpadding="0" cellspacing="0" border="0"><tr>' +
  '<td style="border-top:1px solid #eeeeee;"></td></tr></table></td></tr>' +
  '<tr><td style="padding:20px 30px; text-align:center;">' +
  '<p style="margin:0 0 6px 0; font-family:Arial,sans-serif; font-size:12px; color:#aaaaaa; line-height:1.5;">' +
  'Caso deseje voltar a receber o boletim, escreva para<br>' +
  '<a href="mailto:gustavoportugalhamer@gmail.com" style="color:#00acad; text-decoration:none;">gustavoportugalhamer@gmail.com</a></p>' +
  '<p style="margin:8px 0 0 0; font-family:Arial,sans-serif; font-size:11px; color:#cccccc;">' +
  'Grupo Netuno | <a href="https://r2fcapital.com.br" style="color:#00acad; text-decoration:none;">R2F Capital</a></p>' +
  '</td></tr>' +
  '</table></td></tr></table></body></html>';
}
