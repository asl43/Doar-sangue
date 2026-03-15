// ═══════════════════════════════════════════════════════════════
//  REDE DE DOADORES — Google Apps Script Backend
//  Cole este código em: script.google.com → Novo projeto
//  Depois: Implantar → Novo deployment → App da Web
//  Executar como: Eu | Quem tem acesso: Qualquer pessoa
// ═══════════════════════════════════════════════════════════════

const SHEET_ID = ''; // ← Cole aqui o ID da sua planilha (da URL do Sheets)
// Exemplo: se a URL for https://docs.google.com/spreadsheets/d/1aBcDeFgH.../edit
// então SHEET_ID = '1aBcDeFgH...'

// ── Nomes das abas ───────────────────────────────────────────────
const ABA_DOADORES  = 'Doadores';
const ABA_PACIENTES = 'Pacientes';
const ABA_CHAT      = 'Chat';

// ── Cabeçalhos de cada aba ──────────────────────────────────────
const CAB_DOADORES  = ['ID','Codigo','Nome','Tipo','UltimaDoacao','Cidade','Estado','Apto','DataCadastro'];
const CAB_PACIENTES = ['ID','Codigo','Nome','Hospital','Registro','TipoNecessario','QtdDoadores','Observacoes','Status','DataCadastro'];
const CAB_CHAT      = ['Sala','MensagemID','De','Texto','Hora','DataHora'];

// ════════════════════════════════════════════════════════════════
//  ENTRY POINT — recebe todas as requisições do app
// ════════════════════════════════════════════════════════════════
function doGet(e) {
  return handle(e);
}
function doPost(e) {
  return handle(e);
}

function handle(e) {
  try {
    const params = e.parameter || {};
    const body   = e.postData ? JSON.parse(e.postData.contents || '{}') : {};
    const action = params.action || body.action;

    let result;
    switch (action) {
      case 'ping':           result = { ok: true, msg: 'Conectado!' }; break;
      case 'listarDoadores': result = listarDoadores(); break;
      case 'listarPacientes':result = listarPacientes(); break;
      case 'addDoador':      result = addDoador(body); break;
      case 'addPaciente':    result = addPaciente(body); break;
      case 'atualizarStatus':result = atualizarStatus(body); break;
      case 'deletar':        result = deletar(body); break;
      case 'enviarMsg':      result = enviarMsg(body); break;
      case 'listarMsgs':     result = listarMsgs(params.sala || body.sala); break;
      default:               result = { ok: false, erro: 'Ação desconhecida: ' + action };
    }

    return jsonResp(result);
  } catch (err) {
    return jsonResp({ ok: false, erro: err.message });
  }
}

function jsonResp(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════════════════════════════
//  INICIALIZAÇÃO DA PLANILHA
// ════════════════════════════════════════════════════════════════
function getSheet(nome, cabecalhos) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(nome);
  if (!sheet) {
    sheet = ss.insertSheet(nome);
    sheet.appendRow(cabecalhos);
    // Formata cabeçalho
    const hRange = sheet.getRange(1, 1, 1, cabecalhos.length);
    hRange.setBackground('#c40000').setFontColor('#ffffff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function gerarId() {
  return new Date().getTime() + '_' + Math.random().toString(36).slice(2, 6);
}

function gerarCodigo() {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let c = '';
  for (let i = 0; i < 6; i++) c += chars[Math.floor(Math.random() * chars.length)];
  return c;
}

function agora() {
  return Utilities.formatDate(new Date(), 'America/Manaus', 'dd/MM/yyyy HH:mm');
}

// ════════════════════════════════════════════════════════════════
//  DOADORES
// ════════════════════════════════════════════════════════════════
function addDoador(d) {
  const sheet = getSheet(ABA_DOADORES, CAB_DOADORES);
  const codigo = gerarCodigo();
  const id     = gerarId();

  sheet.appendRow([
    id,
    codigo,
    d.nome || '',
    d.tipo || '',
    d.data || '',
    d.cidade || '',
    d.estado || '',
    d.apto || 'sim',
    agora()
  ]);

  // Alterna cor das linhas
  formatarUltimaLinha(sheet, CAB_DOADORES.length);

  return { ok: true, codigo, id };
}

function listarDoadores() {
  const sheet = getSheet(ABA_DOADORES, CAB_DOADORES);
  const dados = sheet.getDataRange().getValues();
  if (dados.length <= 1) return { ok: true, lista: [] };

  const hoje = new Date();
  const lista = [];

  for (let i = 1; i < dados.length; i++) {
    const r = dados[i];
    const apto     = r[7];
    const ultimaDoa = r[4];

    if (apto !== 'sim') continue;

    // Verifica 90 dias
    if (ultimaDoa) {
      const dataDoc = new Date(ultimaDoa);
      if (!isNaN(dataDoc)) {
        const diff = (hoje - dataDoc) / 86400000;
        if (diff < 90) continue;
      }
    }

    lista.push({
      id: r[0], codigo: r[1], nome: r[2], tipo: r[3],
      data: r[4], cidade: r[5], estado: r[6], apto: r[7],
      linha: i + 1
    });
  }

  return { ok: true, lista };
}

// ════════════════════════════════════════════════════════════════
//  PACIENTES
// ════════════════════════════════════════════════════════════════
function addPaciente(p) {
  const sheet = getSheet(ABA_PACIENTES, CAB_PACIENTES);
  const codigo = gerarCodigo();
  const id     = gerarId();

  sheet.appendRow([
    id,
    codigo,
    p.nome     || '',
    p.hospital || '',
    p.registro || '',
    p.tipo     || '',
    p.qtd      || '',
    p.obs      || '',
    'precisa',
    agora()
  ]);

  formatarUltimaLinha(sheet, CAB_PACIENTES.length);

  return { ok: true, codigo, id };
}

function listarPacientes() {
  const sheet = getSheet(ABA_PACIENTES, CAB_PACIENTES);
  const dados = sheet.getDataRange().getValues();
  if (dados.length <= 1) return { ok: true, lista: [] };

  const lista = [];
  for (let i = 1; i < dados.length; i++) {
    const r = dados[i];
    if (r[8] !== 'precisa') continue;
    lista.push({
      id: r[0], codigo: r[1], nome: r[2], hospital: r[3],
      registro: r[4], tipo: r[5], qtd: r[6], obs: r[7],
      status: r[8], linha: i + 1
    });
  }

  return { ok: true, lista };
}

function atualizarStatus(body) {
  const { colecao, id, novoStatus } = body;
  const aba = colecao === 'doadores' ? ABA_DOADORES : ABA_PACIENTES;
  const cab = colecao === 'doadores' ? CAB_DOADORES : CAB_PACIENTES;
  const sheet = getSheet(aba, cab);
  const dados = sheet.getDataRange().getValues();
  const colStatus = colecao === 'doadores' ? 8 : 9; // coluna Status (1-based)

  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][0]) === String(id)) {
      sheet.getRange(i + 1, colStatus).setValue(novoStatus);
      return { ok: true };
    }
  }
  return { ok: false, erro: 'Registro não encontrado' };
}

function deletar(body) {
  const { colecao, id } = body;
  const aba = colecao === 'doadores' ? ABA_DOADORES : ABA_PACIENTES;
  const cab = colecao === 'doadores' ? CAB_DOADORES : CAB_PACIENTES;
  const sheet = getSheet(aba, cab);
  const dados = sheet.getDataRange().getValues();

  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, erro: 'Registro não encontrado' };
}

// ════════════════════════════════════════════════════════════════
//  CHAT
// ════════════════════════════════════════════════════════════════
function enviarMsg(body) {
  const sheet = getSheet(ABA_CHAT, CAB_CHAT);
  const msgId = gerarId();

  sheet.appendRow([
    body.sala  || '',
    msgId,
    body.de    || 'Anônimo',
    body.texto || '',
    body.hora  || '',
    agora()
  ]);

  formatarUltimaLinha(sheet, CAB_CHAT.length);
  return { ok: true, msgId };
}

function listarMsgs(sala) {
  if (!sala) return { ok: false, erro: 'Sala não informada' };
  const sheet = getSheet(ABA_CHAT, CAB_CHAT);
  const dados = sheet.getDataRange().getValues();

  const msgs = [];
  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][0]) === String(sala)) {
      msgs.push({
        msgId: dados[i][1],
        de:    dados[i][2],
        texto: dados[i][3],
        hora:  dados[i][4]
      });
    }
  }

  return { ok: true, msgs };
}

// ════════════════════════════════════════════════════════════════
//  UTILITÁRIOS
// ════════════════════════════════════════════════════════════════
function formatarUltimaLinha(sheet, numCols) {
  const ultima = sheet.getLastRow();
  if (ultima <= 1) return;
  const cor = ultima % 2 === 0 ? '#fff5f5' : '#ffffff';
  sheet.getRange(ultima, 1, 1, numCols).setBackground(cor);
}

// ── Função de teste — rode manualmente para verificar ───────────
function testar() {
  Logger.log('=== TESTE ===');
  Logger.log(JSON.stringify(addDoador({ nome: 'João Teste', tipo: 'O+', cidade: 'Manaus', estado: 'AM', apto: 'sim' })));
  Logger.log(JSON.stringify(addPaciente({ nome: 'Maria Teste', hospital: 'HPS', tipo: 'O+', qtd: 2 })));
  Logger.log(JSON.stringify(listarDoadores()));
  Logger.log(JSON.stringify(listarPacientes()));
}
