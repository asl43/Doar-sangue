// ═══════════════════════════════════════════════════════════════════
//  REDE DE DOADORES — Google Apps Script
//  Planilha: 1C0dmmsMHo-WNqHfKUxTckBpeYP7_Za_kC2e_r29qFpI
//
//  COMO ATUALIZAR:
//  script.google.com → seu projeto → apague tudo → cole este código
//  Implantar → Gerenciar implantações → ✏️ → Nova versão → Salvar
// ═══════════════════════════════════════════════════════════════════

const SHEET_ID      = '1C0dmmsMHo-WNqHfKUxTckBpeYP7_Za_kC2e_r29qFpI';
const ABA_DOADORES  = 'Doadores';
const ABA_PACIENTES = 'Pacientes';
const ABA_CHAT      = 'Chat';
const ABA_PRESENCA  = 'Presenca';

const CAB_DOADORES  = ['ID','Codigo','Nome','TipoSanguineo','UltimaDoacao','Cidade','Estado','Apto','DataCadastro'];
const CAB_PACIENTES = ['ID','Codigo','Nome','Hospital','Registro','TipoNecessario','QtdDoadores','Observacoes','Status','DataCadastro'];
const CAB_CHAT      = ['Sala','MensagemID','De','NomeExibido','Texto','Hora','DataHoraRegistro'];
const CAB_PRESENCA  = ['Codigo','NomeExibido','UltimoAcesso','Tipo'];

// ── Entry points ────────────────────────────────────────────────
function doGet(e)  { return responder(rotear(e, {})); }
function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch(err) {}
  return responder(rotear(e, body));
}

function rotear(e, body) {
  const p = e.parameter || {};
  const action = p.action || body.action || '';
  try {
    switch (action) {
      case 'ping':            return { ok: true, msg: 'Conectado!' };
      case 'setup':           return setupPlanilha();
      case 'listarDoadores':  return listarDoadores();
      case 'listarPacientes': return listarPacientes();
      case 'addDoador':       return addDoador(body);
      case 'addPaciente':     return addPaciente(body);
      case 'deletar':         return deletar(body);
      case 'atualizarStatus': return atualizarStatus(body);
      case 'enviarMsg':       return enviarMsg(body);
      case 'listarMsgs':      return listarMsgs(p.sala || body.sala);
      case 'heartbeat':       return heartbeat(body);
      case 'presencaSala':    return presencaSala(p.sala || body.sala);
      default:                return { ok: false, erro: 'Ação desconhecida: ' + action };
    }
  } catch(err) {
    return { ok: false, erro: err.message };
  }
}

function responder(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

// ── Utils ────────────────────────────────────────────────────────
function gerarId()     { return String(Date.now()) + '_' + Math.random().toString(36).slice(2,5); }
function gerarCodigo() {
  const c = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  return Array.from({length:6}, () => c[Math.floor(Math.random()*c.length)]).join('');
}
function agora() { return Utilities.formatDate(new Date(), 'America/Manaus', 'dd/MM/yyyy HH:mm'); }
function agoraTs() { return new Date().getTime(); }

function getOuCriarAba(nome, cabs) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh   = ss.getSheetByName(nome);
  if (!sh) {
    sh = ss.insertSheet(nome);
    sh.getRange(1,1,1,cabs.length).setValues([cabs])
      .setBackground('#c40000').setFontColor('#ffffff')
      .setFontWeight('bold').setFontSize(10);
    sh.setFrozenRows(1);
    sh.setColumnWidths(1, cabs.length, 160);
  }
  return sh;
}

function zebraLinha(sh, linha, n) {
  sh.getRange(linha, 1, 1, n).setBackground(linha % 2 === 0 ? '#fff5f5' : '#ffffff');
}

// ── Setup ────────────────────────────────────────────────────────
function setupPlanilha() {
  getOuCriarAba(ABA_DOADORES,  CAB_DOADORES);
  getOuCriarAba(ABA_PACIENTES, CAB_PACIENTES);
  getOuCriarAba(ABA_CHAT,      CAB_CHAT);
  getOuCriarAba(ABA_PRESENCA,  CAB_PRESENCA);
  return { ok: true, msg: 'Planilha configurada!' };
}

// ── Doadores ─────────────────────────────────────────────────────
function addDoador(d) {
  if (!d.nome) return { ok: false, erro: 'Nome obrigatório' };
  if (!d.tipo) return { ok: false, erro: 'Tipo sanguíneo obrigatório' };
  const sh     = getOuCriarAba(ABA_DOADORES, CAB_DOADORES);
  const codigo = gerarCodigo();
  const id     = gerarId();
  sh.appendRow([id, codigo, d.nome, d.tipo, d.data||'', d.cidade||'', d.estado||'', d.apto||'sim', agora()]);
  zebraLinha(sh, sh.getLastRow(), CAB_DOADORES.length);
  return { ok: true, id, codigo };
}

function listarDoadores() {
  const sh    = getOuCriarAba(ABA_DOADORES, CAB_DOADORES);
  const dados = sh.getDataRange().getValues();
  const hoje  = new Date();
  const lista = [];
  for (let i = 1; i < dados.length; i++) {
    const r = dados[i];
    if (String(r[7]).toLowerCase() !== 'sim') continue;
    if (r[4]) {
      const diff = (hoje - new Date(r[4])) / 86400000;
      if (!isNaN(diff) && diff < 90) continue;
    }
    lista.push({ id:r[0], codigo:r[1], nome:r[2], tipo:r[3], data:r[4], cidade:r[5], estado:r[6], apto:r[7], cadastro:r[8] });
  }
  return { ok: true, lista };
}

// ── Pacientes ────────────────────────────────────────────────────
function addPaciente(p) {
  if (!p.nome)     return { ok: false, erro: 'Nome obrigatório' };
  if (!p.hospital) return { ok: false, erro: 'Hospital obrigatório' };
  const sh     = getOuCriarAba(ABA_PACIENTES, CAB_PACIENTES);
  const codigo = gerarCodigo();
  const id     = gerarId();
  sh.appendRow([id, codigo, p.nome, p.hospital, p.registro||'', p.tipo||'', p.qtd||'', p.obs||'', 'precisa', agora()]);
  zebraLinha(sh, sh.getLastRow(), CAB_PACIENTES.length);
  return { ok: true, id, codigo };
}

function listarPacientes() {
  const sh    = getOuCriarAba(ABA_PACIENTES, CAB_PACIENTES);
  const dados = sh.getDataRange().getValues();
  const lista = [];
  for (let i = 1; i < dados.length; i++) {
    const r = dados[i];
    if (String(r[8]).toLowerCase() !== 'precisa') continue;
    lista.push({ id:r[0], codigo:r[1], nome:r[2], hospital:r[3], registro:r[4], tipo:r[5], qtd:r[6], obs:r[7], status:r[8], cadastro:r[9] });
  }
  return { ok: true, lista };
}

function atualizarStatus(body) {
  const { colecao, id, novoStatus } = body;
  const aba  = colecao === 'doadores' ? ABA_DOADORES : ABA_PACIENTES;
  const cabs = colecao === 'doadores' ? CAB_DOADORES : CAB_PACIENTES;
  const colS = colecao === 'doadores' ? 8 : 9;
  const sh   = getOuCriarAba(aba, cabs);
  const dados = sh.getDataRange().getValues();
  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][0]) === String(id)) {
      sh.getRange(i+1, colS).setValue(novoStatus);
      return { ok: true };
    }
  }
  return { ok: false, erro: 'Não encontrado' };
}

function deletar(body) {
  const { colecao, id } = body;
  const aba  = colecao === 'doadores' ? ABA_DOADORES : ABA_PACIENTES;
  const cabs = colecao === 'doadores' ? CAB_DOADORES : CAB_PACIENTES;
  const sh   = getOuCriarAba(aba, cabs);
  const dados = sh.getDataRange().getValues();
  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][0]) === String(id)) { sh.deleteRow(i+1); return { ok: true }; }
  }
  return { ok: false, erro: 'Não encontrado' };
}

// ── Chat ─────────────────────────────────────────────────────────
function enviarMsg(body) {
  const { sala, de, nomeExibido, texto, hora } = body;
  if (!sala || !texto) return { ok: false, erro: 'Sala e texto obrigatórios' };
  const sh    = getOuCriarAba(ABA_CHAT, CAB_CHAT);
  const msgId = gerarId();
  sh.appendRow([sala, msgId, de||'', nomeExibido||'Anônimo', texto, hora||agora(), agora()]);
  zebraLinha(sh, sh.getLastRow(), CAB_CHAT.length);
  return { ok: true, msgId };
}

function listarMsgs(sala) {
  if (!sala) return { ok: false, erro: 'Sala obrigatória' };
  const sh    = getOuCriarAba(ABA_CHAT, CAB_CHAT);
  const dados = sh.getDataRange().getValues();
  const msgs  = [];
  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][0]) === String(sala)) {
      msgs.push({ msgId:dados[i][1], de:dados[i][2], nomeExibido:dados[i][3], texto:dados[i][4], hora:dados[i][5] });
    }
  }
  return { ok: true, msgs };
}

// ── Presença / Heartbeat ─────────────────────────────────────────
// Grava último acesso de cada código. App manda heartbeat a cada 30s.
// "Online" = último acesso há menos de 2 minutos.

function heartbeat(body) {
  const { codigo, nomeExibido, tipo } = body;
  if (!codigo) return { ok: false, erro: 'Código obrigatório' };
  const sh    = getOuCriarAba(ABA_PRESENCA, CAB_PRESENCA);
  const dados = sh.getDataRange().getValues();
  const ts    = agora();

  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][0]) === String(codigo)) {
      sh.getRange(i+1, 3).setValue(ts); // atualiza UltimoAcesso
      return { ok: true };
    }
  }
  // Novo registro de presença
  sh.appendRow([codigo, nomeExibido||codigo, ts, tipo||'']);
  zebraLinha(sh, sh.getLastRow(), CAB_PRESENCA.length);
  return { ok: true };
}

function presencaSala(sala) {
  // Retorna presença dos dois participantes de uma sala (código do doador/paciente)
  if (!sala) return { ok: false, erro: 'Sala obrigatória' };
  const sh    = getOuCriarAba(ABA_PRESENCA, CAB_PRESENCA);
  const dados = sh.getDataRange().getValues();
  const agr   = new Date();

  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][0]) === String(sala)) {
      const ultimo = dados[i][2]; // UltimoAcesso string dd/MM/yyyy HH:mm
      // Converte para Date (formato pt-br)
      let online = false;
      try {
        const partes = String(ultimo).match(/(\d{2})\/(\d{2})\/(\d{4}) (\d{2}):(\d{2})/);
        if (partes) {
          const dt = new Date(partes[3], partes[2]-1, partes[1], partes[4], partes[5]);
          const diffMin = (agr - dt) / 60000;
          online = diffMin <= 2;
        }
      } catch(e) {}
      return { ok: true, codigo: sala, nomeExibido: dados[i][1], ultimoAcesso: dados[i][2], online, tipo: dados[i][3] };
    }
  }
  return { ok: true, codigo: sala, online: false, ultimoAcesso: null };
}

// ── Teste manual ─────────────────────────────────────────────────
function testar() {
  Logger.log(JSON.stringify(setupPlanilha()));
  Logger.log(JSON.stringify(addDoador({ nome:'João', tipo:'O+', cidade:'Manaus', estado:'AM', apto:'sim' })));
  Logger.log(JSON.stringify(addPaciente({ nome:'Maria', hospital:'HPS', tipo:'O+', qtd:'2' })));
  Logger.log(JSON.stringify(listarDoadores()));
  Logger.log(JSON.stringify(listarPacientes()));
}
