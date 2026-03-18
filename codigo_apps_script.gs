// ═══════════════════════════════════════════════════════════════════
//  REDE DE DOADORES — Google Apps Script  (versão 2.0)
//  Planilha: 1C0dmmsMHo-WNqHfKUxTckBpeYP7_Za_kC2e_r29qFpI
//
//  NOVIDADES v2.0:
//  • Doadores: foto (base64), histórico de doações manuais
//  • Pacientes: foto, endereço do hospital, data-prazo de expiração
//  • Exclusão automática: pacientes expiram no dia seguinte ao prazo
//    (ou após 181 dias do cadastro se prazo não informado)
//  • Trigger diário: instalar executando instalarTrigger()
//
//  COMO INSTALAR / ATUALIZAR:
//  1. script.google.com → seu projeto → apague tudo → cole este código
//  2. Salve (Ctrl+S)
//  3. Execute a função instalarTrigger() UMA VEZ (para exclusão automática)
//  4. Implantar → Gerenciar implantações → ✏️ → Nova versão → Salvar
// ═══════════════════════════════════════════════════════════════════

const SHEET_ID      = '1C0dmmsMHo-WNqHfKUxTckBpeYP7_Za_kC2e_r29qFpI';
const ABA_DOADORES  = 'Doadores';
const ABA_PACIENTES = 'Pacientes';
const ABA_CHAT      = 'Chat';
const ABA_PRESENCA  = 'Presenca';

// Colunas atualizadas — NÃO mude a ordem sem migrar os dados
const CAB_DOADORES  = ['ID','Codigo','Nome','TipoSanguineo','Cidade','Estado','Apto','Foto','Doacoes','DataCadastro'];
const CAB_PACIENTES = ['ID','Codigo','Nome','Hospital','Endereco','Registro','TipoNecessario','QtdDoadores','Observacoes','Prazo','Status','Foto','DataCadastro'];
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
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Utils ────────────────────────────────────────────────────────
function gerarId()     { return String(Date.now()) + '_' + Math.random().toString(36).slice(2,5); }
function gerarCodigo() {
  const c = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  return Array.from({length:6}, () => c[Math.floor(Math.random()*c.length)]).join('');
}
function agora() { return Utilities.formatDate(new Date(), 'America/Manaus', 'dd/MM/yyyy HH:mm'); }
function agoraISO() { return Utilities.formatDate(new Date(), 'America/Manaus', 'yyyy-MM-dd'); }

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
// Colunas: ID(0) Codigo(1) Nome(2) TipoSanguineo(3) Cidade(4) Estado(5) Apto(6) Foto(7) Doacoes(8) DataCadastro(9)
function addDoador(d) {
  if (!d.nome) return { ok: false, erro: 'Nome obrigatório' };
  if (!d.tipo) return { ok: false, erro: 'Tipo sanguíneo obrigatório' };
  const sh     = getOuCriarAba(ABA_DOADORES, CAB_DOADORES);
  const codigo = gerarCodigo();
  const id     = gerarId();
  const foto   = d.foto   || '';   // base64 comprimido
  const doacoes = d.doacoes || '[]'; // JSON string

  sh.appendRow([id, codigo, d.nome, d.tipo, d.cidade||'', d.estado||'', d.apto||'sim', foto, doacoes, agora()]);
  zebraLinha(sh, sh.getLastRow(), CAB_DOADORES.length);
  return { ok: true, id, codigo };
}

function listarDoadores() {
  const sh    = getOuCriarAba(ABA_DOADORES, CAB_DOADORES);
  const dados = sh.getDataRange().getValues();
  const lista = [];
  for (let i = 1; i < dados.length; i++) {
    const r = dados[i];
    if (String(r[6]).toLowerCase() !== 'sim') continue; // apto

    // Analisa histórico de doações para verificar intervalo mínimo (90 dias)
    let bloqueadoPorDoacao = false;
    try {
      const doacoes = JSON.parse(String(r[8]) || '[]');
      if (doacoes.length > 0) {
        // Pega a data mais recente
        const datas = doacoes.map(d => new Date(d.data)).filter(d => !isNaN(d));
        if (datas.length > 0) {
          const maisRecente = new Date(Math.max(...datas));
          const diff = (new Date() - maisRecente) / 86400000;
          if (diff < 90) bloqueadoPorDoacao = true;
        }
      }
    } catch(e) {}
    if (bloqueadoPorDoacao) continue;

    let doacoesParsed = [];
    try { doacoesParsed = JSON.parse(String(r[8]) || '[]'); } catch(e){}

    lista.push({
      id: r[0], codigo: r[1], nome: r[2], tipo: r[3],
      cidade: r[4], estado: r[5], apto: r[6],
      foto: r[7] || '',
      doacoes: doacoesParsed,
      cadastro: r[9]
    });
  }
  return { ok: true, lista };
}

// ── Pacientes ────────────────────────────────────────────────────
// Colunas: ID(0) Codigo(1) Nome(2) Hospital(3) Endereco(4) Registro(5) TipoNecessario(6) QtdDoadores(7) Observacoes(8) Prazo(9) Status(10) Foto(11) DataCadastro(12)
function addPaciente(p) {
  if (!p.nome)     return { ok: false, erro: 'Nome obrigatório' };
  if (!p.hospital) return { ok: false, erro: 'Hospital obrigatório' };
  const sh     = getOuCriarAba(ABA_PACIENTES, CAB_PACIENTES);
  const codigo = gerarCodigo();
  const id     = gerarId();
  const foto   = p.foto || '';
  const prazo  = p.prazo || ''; // formato ISO: yyyy-MM-dd

  sh.appendRow([id, codigo, p.nome, p.hospital, p.endereco||'', p.registro||'', p.tipo||'', p.qtd||'', p.obs||'', prazo, 'precisa', foto, agora()]);
  zebraLinha(sh, sh.getLastRow(), CAB_PACIENTES.length);
  return { ok: true, id, codigo };
}

function listarPacientes() {
  const sh    = getOuCriarAba(ABA_PACIENTES, CAB_PACIENTES);
  const dados = sh.getDataRange().getValues();
  const hoje  = new Date();
  hoje.setHours(0,0,0,0);
  const lista = [];

  for (let i = 1; i < dados.length; i++) {
    const r = dados[i];
    if (String(r[10]).toLowerCase() !== 'precisa') continue; // status

    // Verifica expiração
    if (estaExpirado(r, hoje)) continue;

    lista.push({
      id: r[0], codigo: r[1], nome: r[2], hospital: r[3],
      endereco: r[4], registro: r[5], tipo: r[6], qtd: r[7],
      obs: r[8], prazo: r[9], status: r[10],
      foto: r[11] || '',
      cadastro: r[12]
    });
  }
  return { ok: true, lista };
}

// Verifica se registro de paciente está expirado
function estaExpirado(r, hoje) {
  const prazo    = String(r[9]).trim();    // coluna Prazo (ISO yyyy-MM-dd)
  const cadastro = String(r[12]).trim();   // coluna DataCadastro (dd/MM/yyyy HH:mm)

  if (prazo && prazo.match(/^\d{4}-\d{2}-\d{2}$/)) {
    // Tem data prazo: expira no dia SEGUINTE ao prazo
    const venc = new Date(prazo + 'T00:00:00');
    venc.setDate(venc.getDate() + 1); // dia seguinte
    return hoje >= venc;
  } else {
    // Sem prazo: expira após 181 dias do cadastro
    try {
      const partes = cadastro.match(/(\d{2})\/(\d{2})\/(\d{4})/);
      if (partes) {
        const dataCad = new Date(partes[3], partes[2]-1, partes[1]);
        const limite  = new Date(dataCad);
        limite.setDate(limite.getDate() + 181);
        return hoje >= limite;
      }
    } catch(e) {}
  }
  return false;
}

function atualizarStatus(body) {
  const { colecao, id, novoStatus } = body;
  const aba  = colecao === 'doadores' ? ABA_DOADORES : ABA_PACIENTES;
  const cabs = colecao === 'doadores' ? CAB_DOADORES : CAB_PACIENTES;
  const colS = colecao === 'doadores' ? 7 : 11; // coluna Status (1-indexed)
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

// ── Exclusão automática diária ────────────────────────────────────
// Execute instalarTrigger() UMA VEZ para agendar a limpeza diária.
function instalarTrigger() {
  // Remove triggers antigos com o mesmo nome para não duplicar
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'limpezaAutomatica') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('limpezaAutomatica')
    .timeBased().everyDays(1).atHour(3).create(); // roda às 3h AM (fuso Manaus)
  Logger.log('Trigger instalado: limpezaAutomatica roda diariamente às 3h.');
}

function limpezaAutomatica() {
  const sh    = getOuCriarAba(ABA_PACIENTES, CAB_PACIENTES);
  const dados = sh.getDataRange().getValues();
  const hoje  = new Date(); hoje.setHours(0,0,0,0);
  let removidos = 0;

  // Percorre de baixo para cima para não deslocar índices ao deletar
  for (let i = dados.length - 1; i >= 1; i--) {
    const r = dados[i];
    if (String(r[10]).toLowerCase() === 'precisa' && estaExpirado(r, hoje)) {
      sh.deleteRow(i+1);
      removidos++;
    }
  }
  Logger.log('Limpeza automática: ' + removidos + ' paciente(s) expirado(s) removido(s) em ' + agora());
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
function heartbeat(body) {
  const { codigo, nomeExibido, tipo } = body;
  if (!codigo) return { ok: false, erro: 'Código obrigatório' };
  const sh    = getOuCriarAba(ABA_PRESENCA, CAB_PRESENCA);
  const dados = sh.getDataRange().getValues();
  const ts    = agora();
  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][0]) === String(codigo)) {
      sh.getRange(i+1, 3).setValue(ts);
      return { ok: true };
    }
  }
  sh.appendRow([codigo, nomeExibido||codigo, ts, tipo||'']);
  zebraLinha(sh, sh.getLastRow(), CAB_PRESENCA.length);
  return { ok: true };
}

function presencaSala(sala) {
  if (!sala) return { ok: false, erro: 'Sala obrigatória' };
  const sh    = getOuCriarAba(ABA_PRESENCA, CAB_PRESENCA);
  const dados = sh.getDataRange().getValues();
  const agr   = new Date();
  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][0]) === String(sala)) {
      const ultimo = dados[i][2];
      let online = false;
      try {
        const partes = String(ultimo).match(/(\d{2})\/(\d{2})\/(\d{4}) (\d{2}):(\d{2})/);
        if (partes) {
          const dt = new Date(partes[3], partes[2]-1, partes[1], partes[4], partes[5]);
          online = (agr - dt) / 60000 <= 2;
        }
      } catch(e) {}
      return { ok:true, codigo:sala, nomeExibido:dados[i][1], ultimoAcesso:dados[i][2], online, tipo:dados[i][3] };
    }
  }
  return { ok: true, codigo: sala, online: false, ultimoAcesso: null };
}

// ── Testes ───────────────────────────────────────────────────────
function testar() {
  Logger.log(JSON.stringify(setupPlanilha()));
  Logger.log(JSON.stringify(addDoador({ nome:'João Silva', tipo:'O+', cidade:'Manaus', estado:'AM', apto:'sim', foto:'', doacoes:'[]' })));
  Logger.log(JSON.stringify(addPaciente({ nome:'Maria Santos', hospital:'HPS', endereco:'Av. Constantino Nery, 4000', tipo:'O+', qtd:'2', prazo:'2025-12-31', foto:'' })));
  Logger.log(JSON.stringify(listarDoadores()));
  Logger.log(JSON.stringify(listarPacientes()));
}

function testarLimpeza() {
  Logger.log('Antes:');
  Logger.log(JSON.stringify(listarPacientes()));
  limpezaAutomatica();
  Logger.log('Depois:');
  Logger.log(JSON.stringify(listarPacientes()));
}
