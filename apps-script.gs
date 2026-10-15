// Planilha do cronograma: guarda sugestões, pautas novas e alterações.
// Como instalar:
//  1) Em uma planilha do Google: Extensões > Apps Script. Cole este código.
//  2) Em Configurações do projeto > Propriedades do script, crie:
//       ADMIN_SENHA     = a senha do administrador
//       SENHA_JOAO, SENHA_RENATA, SENHA_ISABELA, SENHA_MARCOS,
//       SENHA_WANDINHO, SENHA_MARIANA, SENHA_FLAVIO = a senha de cada pessoa
//     (as senhas ficam só aqui, nunca no código)
//  3) Implantar > Nova implantação > App da Web.
//       Executar como: Eu | Quem pode acessar: Qualquer pessoa
//  4) Copie o URL gerado e cole em SUGGEST_URL, no começo do script do index.html.

const TIPOS = ['Vídeo', 'Carrossel', 'Card', 'Stories'];
const ESQUEMA = {
  Sugestoes:   ['id', 'data', 'tipo', 'titulo', 'link', 'criadoEm', 'status'],
  Pautas:      ['id', 'data', 'titulo', 'responsavel', 'obs', 'sugestaoId'],
  Alteracoes:  ['chave', 'data', 'responsavel', 'titulo', 'obs', 'removida', 'feito', 'feitoPor']
};

function aba_(nome) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(nome);
  if (!sh) {
    sh = ss.insertSheet(nome);
    sh.appendRow(ESQUEMA[nome]);
    sh.getRange(1, 1, 1, ESQUEMA[nome].length).setFontWeight('bold');
    sh.getRange(2, 1, sh.getMaxRows() - 1, ESQUEMA[nome].length).setNumberFormat('@'); // texto puro: evita datas e fórmulas
  }
  return sh;
}
function txt_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return v === null || v === undefined ? '' : String(v);
}
function ler_(nome) {
  const v = aba_(nome).getDataRange().getValues();
  const cab = v[0];
  return v.slice(1).filter(r => r[0] !== '').map(r => {
    const o = {};
    cab.forEach((k, i) => { o[k] = (k === 'removida' || k === 'feito') ? (r[i] === true || r[i] === 'TRUE') : txt_(r[i]); });
    return o;
  });
}
function resp_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function sanear_(v) { // impede que um texto vire fórmula na planilha
  v = txt_(v);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}
function semAcento_(n) { return String(n).normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase(); }
function ehAdmin_(c) {
  const s = PropertiesService.getScriptProperties().getProperty('ADMIN_SENHA');
  return c.pessoa === 'Admin' && !!s && c.senha === s;
}
function ehPessoa_(c) {
  if (!c.pessoa || c.pessoa === 'Admin') return false;
  const s = PropertiesService.getScriptProperties().getProperty('SENHA_' + semAcento_(c.pessoa));
  return !!s && c.senha === s;
}
function dataOk_(d) { return /^\d{4}-\d{2}-\d{2}$/.test(String(d)); }
function achar_(sh, chave) { // linha (1-based) da primeira coluna igual a chave, ou 0
  const col = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  for (let i = 1; i < col.length; i++) if (txt_(col[i][0]) === chave) return i + 1;
  return 0;
}
function upsertAlteracao_(chave, campos) {
  const sh = aba_('Alteracoes');
  const cab = ESQUEMA.Alteracoes;
  let linha = achar_(sh, chave);
  if (!linha) { sh.appendRow([chave]); linha = sh.getLastRow(); sh.getRange(linha, 1, 1, cab.length).setNumberFormat('@'); sh.getRange(linha, 1).setValue(chave); }
  Object.keys(campos).forEach(k => {
    const i = cab.indexOf(k);
    if (i > 0) sh.getRange(linha, i + 1).setValue(typeof campos[k] === 'boolean' ? campos[k] : sanear_(campos[k]));
  });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const c = JSON.parse(e.postData.contents);
    const a = c.acao;

    if (a === 'sugerir') {
      if (!dataOk_(c.data) || TIPOS.indexOf(c.tipo) < 0 || !c.titulo || String(c.titulo).length > 120 ||
          !/^https?:\/\//i.test(c.link) || String(c.link).length > 500) return resp_({ ok: false, erro: 'dados inválidos' });
      aba_('Sugestoes').appendRow([Utilities.getUuid(), c.data, c.tipo, sanear_(c.titulo), c.link, new Date().toISOString(), 'pendente']);
      return resp_({ ok: true });
    }

    if (a === 'marcar') {
      // a própria pessoa (só nas pautas dela) ou o administrador
      if (!(ehAdmin_(c) || (ehPessoa_(c) && c.pessoa === c.responsavel))) return resp_({ ok: false, erro: 'senha' });
      upsertAlteracao_(String(c.chave), { feito: !!c.feito, feitoPor: c.pessoa });
      return resp_({ ok: true });
    }

    if (!ehAdmin_(c)) return resp_({ ok: false, erro: 'senha' });

    if (a === 'rejeitar') {
      const sh = aba_('Sugestoes'), l = achar_(sh, String(c.id));
      if (l) sh.getRange(l, ESQUEMA.Sugestoes.indexOf('status') + 1).setValue('rejeitada');
      return resp_({ ok: true });
    }
    if (a === 'criar' || a === 'aprovar') {
      if (!dataOk_(c.data) || !c.titulo) return resp_({ ok: false, erro: 'dados inválidos' });
      const id = 'x:' + Utilities.getUuid();
      aba_('Pautas').appendRow([id, c.data, sanear_(c.titulo), sanear_(c.responsavel || 'A definir'), sanear_(c.obs || ''), a === 'aprovar' ? String(c.id) : '']);
      if (a === 'aprovar') {
        const sh = aba_('Sugestoes'), l = achar_(sh, String(c.id));
        if (l) sh.getRange(l, ESQUEMA.Sugestoes.indexOf('status') + 1).setValue('aprovada');
      }
      return resp_({ ok: true, id: id });
    }
    if (a === 'alterar') {
      const campos = {};
      ['data', 'responsavel', 'titulo', 'obs', 'removida'].forEach(k => { if (c[k] !== undefined) campos[k] = c[k]; });
      if (campos.data !== undefined && !dataOk_(campos.data)) return resp_({ ok: false, erro: 'dados inválidos' });
      upsertAlteracao_(String(c.chave), campos);
      return resp_({ ok: true });
    }
    return resp_({ ok: false, erro: 'ação desconhecida' });
  } catch (err) {
    return resp_({ ok: false, erro: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return resp_({ sugestoes: ler_('Sugestoes'), extras: ler_('Pautas'), alteracoes: ler_('Alteracoes') });
}
