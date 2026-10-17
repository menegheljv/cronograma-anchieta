// Planilha do cronograma: guarda sugestões, pautas novas e alterações.
// Como instalar:
//  1) Em uma planilha do Google: Extensões > Apps Script. Cole este código.
//  2) Em Configurações do projeto > Propriedades do script, crie:
//       ADMIN_SENHA = a senha do administrador (fica só aqui, nunca no código)
//     A equipe entra só com o nome, sem senha.
//  3) Implantar > Nova implantação > App da Web.
//       Executar como: Eu | Quem pode acessar: Qualquer pessoa
//  4) Copie o URL gerado e cole em SUGGEST_URL, no começo do script do index.html.

const TIPOS = ['Vídeo', 'Foto', 'Stories', 'Carrossel', 'Cobertura', 'Card'];
// Planilha de respostas do formulário de demandas. A conta que publica este script precisa ter acesso a ela.
const DEMANDAS_PLANILHA = '1KUYV7CNIi0U2AjwB4ibu13Mv78am1W3hesGJuip53ww';
const DEMANDAS_ABA = 'Respostas ao formulário 1';
const ESQUEMA = {
  Sugestoes:   ['id', 'data', 'tipo', 'titulo', 'link', 'criadoEm', 'status', 'info'],
  Pautas:      ['id', 'data', 'titulo', 'responsavel', 'obs', 'sugestaoId', 'formato', 'link'],
  Alteracoes:  ['chave', 'data', 'responsavel', 'titulo', 'obs', 'removida', 'feito', 'feitoPor', 'formato', 'link', 'programada', 'legenda', 'horario', 'colab']
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
function txt_(v, hora) {
  if (v instanceof Date && hora) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'HH:mm');
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const t = v === null || v === undefined ? '' : String(v);
  return /^'[=+\-]/.test(t) ? t.slice(1) : t;
}
function ler_(nome) {
  const v = aba_(nome).getDataRange().getValues();
  const cab = v[0];
  return v.slice(1).filter(r => r[0] !== '').map(r => {
    const o = {};
    cab.forEach((k, i) => { o[k] = (k === 'removida' || k === 'feito' || k === 'programada') ? (r[i] === true || /^true$/i.test(String(r[i]).trim())) : txt_(r[i], k === 'horario'); });
    return o;
  });
}
function resp_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function sanear_(v) { // impede que um texto vire fórmula na planilha
  v = txt_(v);
  return /^[=+\-]/.test(v) ? "'" + v : v;   // o @ de arroba é permitido
}
function ehAdmin_(c) {
  const s = PropertiesService.getScriptProperties().getProperty('ADMIN_SENHA');
  // o site envia o que foi digitado em minúsculas e sem espaços; aqui a senha guardada é tratada igual (maiúsculas e espaços não importam)
  const norm = t => String(t === null || t === undefined ? '' : t).trim().toLowerCase();
  return c.pessoa === 'Admin' && !!s && norm(s) !== '' && norm(c.senha) === norm(s);
}
const EQUIPE = ['João', 'Renata', 'Isabela', 'Marcos', 'Wandinho', 'Mariana', 'Flávio'];
function ehPessoa_(c) { // a equipe entra só com o nome
  return EQUIPE.indexOf(c.pessoa) >= 0;
}
function dataOk_(d) { return /^\d{4}-\d{2}-\d{2}$/.test(String(d)); }
function achar_(sh, chave) { // linha (1-based) da primeira coluna igual a chave, ou 0
  const col = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  for (let i = 1; i < col.length; i++) if (txt_(col[i][0]) === chave) return i + 1;
  return 0;
}
function upsertAlteracao_(chave, campos) {
  // uma leitura e uma gravação por pauta (antes eram várias chamadas, o que deixava o check lento)
  const sh = aba_('Alteracoes');
  const cab = ESQUEMA.Alteracoes;
  let linha = achar_(sh, chave);
  let fila;
  if (linha) fila = sh.getRange(linha, 1, 1, cab.length).getValues()[0];
  else { fila = cab.map(() => ''); fila[0] = chave; linha = sh.getLastRow() + 1; sh.getRange(linha, 1, 1, cab.length).setNumberFormat('@'); }
  Object.keys(campos).forEach(k => {
    const i = cab.indexOf(k);
    if (i > 0) fila[i] = typeof campos[k] === 'boolean' ? campos[k] : sanear_(campos[k]);
  });
  sh.getRange(linha, 1, 1, cab.length).setValues([fila]);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const c = JSON.parse(e.postData.contents);
    const a = c.acao;

    if (a === 'sugerir') {
      if (!dataOk_(c.data) || TIPOS.indexOf(c.tipo) < 0 || !c.titulo || String(c.titulo).length > 120 ||
          (c.link && !/^https?:\/\//i.test(c.link)) || String(c.link || '').length > 500) return resp_({ ok: false, erro: 'dados inválidos' });
      aba_('Sugestoes').appendRow([Utilities.getUuid(), c.data, c.tipo, sanear_(c.titulo), c.link || '', new Date().toISOString(), 'pendente', sanear_(String(c.info || '').slice(0, 300))]);
      return resp_({ ok: true });
    }

    if (a === 'contatos') {
      if (!(ehAdmin_(c) || ehPessoa_(c))) return resp_({ ok: false, erro: 'senha' });
      try { return resp_({ ok: true, contatos: contatos_() }); } catch (err) { return resp_({ ok: true, contatos: {} }); }
    }

    if (a === 'detalhar') {
      // legenda, horário e colab: o administrador ou a própria pessoa (só nas pautas dela, sem senha)
      if (!(ehAdmin_(c) || (ehPessoa_(c) && c.pessoa === c.responsavel))) return resp_({ ok: false, erro: 'senha' });
      const leg = String(c.legenda || '').trim().slice(0, 2200), hor = String(c.horario || '').trim(), col = String(c.colab || '').trim().slice(0, 100);
      if (hor && !/^\d{2}:\d{2}$/.test(hor)) return resp_({ ok: false, erro: 'dados inválidos' });
      upsertAlteracao_(String(c.chave), { legenda: leg, horario: hor, colab: col });
      return resp_({ ok: true });
    }

    if (a === 'concluir') {
      // legenda, horário, colab e o check de uma vez só (a própria pessoa nas pautas dela, ou o administrador)
      if (!(ehAdmin_(c) || (ehPessoa_(c) && c.pessoa === c.responsavel))) return resp_({ ok: false, erro: 'senha' });
      const leg = String(c.legenda || '').trim().slice(0, 2200), hor = String(c.horario || '').trim(), col = String(c.colab || '').trim().slice(0, 100);
      if (hor && !/^\d{2}:\d{2}$/.test(hor)) return resp_({ ok: false, erro: 'dados inválidos' });
      upsertAlteracao_(String(c.chave), { legenda: leg, horario: hor, colab: col, feito: true, feitoPor: c.pessoa });
      return resp_({ ok: true });
    }

    if (a === 'marcar') {
      // a própria pessoa (só nas pautas dela, sem senha) ou o administrador
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
      const link = /^https?:\/\//i.test(c.link || '') ? String(c.link) : '';
      aba_('Pautas').appendRow([id, c.data, sanear_(c.titulo), sanear_(c.responsavel || 'A definir'), sanear_(c.obs || ''), a === 'aprovar' ? String(c.id) : '', sanear_(c.formato || ''), link]);
      if (a === 'aprovar') {
        const sh = aba_('Sugestoes'), l = achar_(sh, String(c.id));
        if (l) sh.getRange(l, ESQUEMA.Sugestoes.indexOf('status') + 1).setValue('aprovada');
      }
      return resp_({ ok: true, id: id });
    }
    if (a === 'alterar') {
      const campos = {};
      ['data', 'responsavel', 'titulo', 'obs', 'formato', 'link', 'removida', 'programada'].forEach(k => { if (c[k] !== undefined) campos[k] = c[k]; });
      if (campos.link !== undefined && String(campos.link).trim() && !/^https?:\/\//i.test(String(campos.link).trim())) return resp_({ ok: false, erro: 'dados inválidos' });
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

// Só entrega campos seguros das demandas: e-mail e telefone NÃO saem da planilha. Os arquivos anexados saem como links do Drive.
function demandas_() {
  const sh = SpreadsheetApp.openById(DEMANDAS_PLANILHA).getSheetByName(DEMANDAS_ABA);
  if (!sh) return [];
  const v = sh.getDataRange().getValues();
  const norm = t => String(t).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
  const cab = v[0].map(norm);
  const col = pref => cab.findIndex(c => c.indexOf(norm(pref)) === 0);
  const idx = {
    ts: col('carimbo'), secretaria: col('secretaria'), solicitante: col('nome do solicitante'),
    tipo: col('tipo de material'), objetivo: col('objetivo da'), formato: col('formato desejado'),
    prazo: col('prazo desejado'), info: col('informacoes necessarias'), urgente: col('a demanda e urgente'),
    justif: col('se "sim"'), anexos: col('anexar'), contato: col('contato do solicitante')
  };
  const t = (r, k) => idx[k] < 0 ? '' : (r[idx[k]] instanceof Date ? Utilities.formatDate(r[idx[k]], Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss') : String(r[idx[k]] === null ? '' : r[idx[k]]).trim());
  return v.slice(1).filter(r => t(r, 'ts') || t(r, 'tipo')).map((r, i) => ({
    id: String(i + 2), ts: t(r, 'ts'), secretaria: t(r, 'secretaria'), solicitante: t(r, 'solicitante'),
    tipo: t(r, 'tipo'), objetivo: t(r, 'objetivo'), formato: t(r, 'formato'), prazo: t(r, 'prazo'),
    info: t(r, 'info'), urgente: t(r, 'urgente'), justif: t(r, 'justif'),
    arquivos: (t(r, 'anexos').match(/https:\/\/(?:drive|docs)\.google\.com\/[^\s,]+/g) || []).slice(0, 10),
    anexos: t(r, 'anexos') ? t(r, 'anexos').split(',').length : 0
  }));
}

// Contato (telefone ou e-mail) de quem pediu: só sai para a equipe (nome) ou para o administrador, nunca pela leitura pública.
function contatos_() {
  const sh = SpreadsheetApp.openById(DEMANDAS_PLANILHA).getSheetByName(DEMANDAS_ABA);
  if (!sh) return {};
  const v = sh.getDataRange().getValues();
  const norm = t => String(t).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
  const i = v[0].map(norm).findIndex(c => c.indexOf('contato do solicitante') === 0);
  const o = {};
  if (i < 0) return o;
  v.slice(1).forEach((r, k) => { const t = String(r[i] === null ? '' : r[i]).trim(); if (t) o[String(k + 2)] = t.slice(0, 120); });
  return o;
}

function doGet(e) {
  if (e && e.parameter && e.parameter.acao === 'diag') {   // diagnóstico: mostra a versão do script e se a senha existe (nunca mostra a senha)
    const sp = PropertiesService.getScriptProperties().getProperty('ADMIN_SENHA');
    return resp_({ versao: 'senha-normalizada-v2', temSenha: !!sp, tamanhoSenha: sp ? String(sp).trim().length : 0, temEspacoOuTexto: sp ? /\s/.test(String(sp).trim()) : false });
  }
  if (e && e.parameter && e.parameter.acao === 'demandas') {
    try { return resp_({ demandas: demandas_() }); }
    catch (err) { return resp_({ demandas: [], erro: String(err) }); }
  }
  return resp_({ sugestoes: ler_('Sugestoes'), extras: ler_('Pautas'), alteracoes: ler_('Alteracoes') });
}
