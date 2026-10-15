// Cole este código em Extensões > Apps Script de uma planilha do Google.
// Depois: Implantar > Nova implantação > App da Web
//   Executar como: Eu | Quem pode acessar: Qualquer pessoa
// Copie o URL gerado e cole em SUGGEST_URL no index.html.
const ABA = 'Sugestoes';
const TIPOS = ['Vídeo', 'Carrossel', 'Card', 'Stories'];

function aba_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(ABA);
  if (!sh) { sh = ss.insertSheet(ABA); sh.appendRow(['id', 'data', 'tipo', 'link', 'criadoEm']); }
  return sh;
}
function resp_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const d = JSON.parse(e.postData.contents);
    if (!/^\d{4}-\d{2}-\d{2}$/.test(d.data) || TIPOS.indexOf(d.tipo) < 0 || !/^https?:\/\//i.test(d.link) || d.link.length > 500)
      return resp_({ ok: false, erro: 'dados inválidos' });
    aba_().appendRow([Utilities.getUuid(), d.data, d.tipo, d.link, new Date().toISOString()]);
    return resp_({ ok: true });
  } catch (err) {
    return resp_({ ok: false, erro: String(err) });
  } finally {
    lock.releaseLock();
  }
}
function doGet() {
  const v = aba_().getDataRange().getValues().slice(1);
  return resp_(v.map(r => ({ id: r[0], data: r[1], tipo: r[2], link: r[3], criadoEm: r[4] })));
}
