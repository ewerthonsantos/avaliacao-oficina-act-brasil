/**
 * Código do Google Apps Script para receber as respostas do formulário de
 * avaliação da oficina (ACT-Brasil) e gravá-las em uma planilha do Google
 * Sheets. Veja o passo a passo de publicação nas instruções enviadas junto
 * com este arquivo.
 */

// Nome da aba onde as respostas serão gravadas. Se a aba não existir, ela é
// criada automaticamente.
const SHEET_NAME = 'Respostas';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
    }

    // Na primeira resposta, escreve a linha de cabeçalho com base nas
    // perguntas enviadas pela página (assim você não precisa digitar os
    // cabeçalhos manualmente, nem mantê-los sincronizados com o HTML).
    if (sheet.getLastRow() === 0) {
      const headers = ['Data e hora', ...data.questions];
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    const timestamp = Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm:ss');
    sheet.appendRow([timestamp, ...data.answers]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
