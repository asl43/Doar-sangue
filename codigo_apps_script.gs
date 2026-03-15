// ═══════════════════════════════════════════════════════════════════
//  REDE DE DOADORES — Google Apps Script
//  Planilha: https://docs.google.com/spreadsheets/d/1C0dmmsMHo-WNqHfKUxTckBpeYP7_Za_kC2e_r29qFpI
//
//  INSTRUÇÕES:
//  1. Abra script.google.com/home
//  2. Abra o projeto vinculado à planilha (ou crie novo e cole este código)
//  3. Clique em "Implantar" → "Gerenciar implantações"
//  4. Se já implantou antes: clique no lápis ✏️ → "Nova versão" → Salvar
//  5. Se é novo: Implantar → Novo deployment → App da Web
//     Executar como: EU | Acesso: QUALQUER PESSOA → Implantar
//  6. Copie a URL /exec gerada
// ═══════════════════════════════════════════════════════════════════

const SHEET_ID      = '1C0dmmsMHo-WNqHfKUxTckBpeYP7_Za_kC2e_r29qFpI';
const ABA_DOADORES  = 'Doadores';
const ABA_PACIENTES = 'Pacientes';
const ABA_CHAT      = 'Chat';

const CAB_DOADORES  = ['ID','Codigo','Nome','TipoSanguineo','UltimaDoacao','Cidade','Estado','Apto','DataCadastro'];
const CAB_PACIENTES = ['ID','Codigo','Nome','Hospital','Registro','TipoNecessario','QtdDoadores','Observacoes','Status','DataCadastro'];
const CAB_CHAT      = ['Sala','MensagemID','De','Texto','Hora','DataHoraRegistro'];

// ═══════════════════════════════════════════════════════════════════
//  ENTRY POINTS
// ═══════════════════════════════════════════════════════════════════
function doGet(e)  { return responder(rotear(e, {})); }
function doPost(e) {
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch(err) {}
  return responder(rotear(e, body));
}

function rotear(e, body) {
  const action = (e.parameter && e.parameter.action) || body.action || '';
  try {
    switch (action) {
      case 'ping':             return { ok: true, msg: 'Conectado! Planilha: ' + SHEET_ID };
      case 'setup':            return setupPlanilha();
      case 'listarDoadores':   return listarDoadores();
      case 'listarPacientes':  return listarPacientes();
      case 'addDoador':        return addDoador(body);
      case 'addPaciente':      return addPaciente(body);
      case 'editarDoador':     return editarRegistro(ABA_DOADORES, CAB_DOADORES, body);
      case 'editarPaciente':   return editarRegistro(ABA_PACIENTES, CAB_PACIENTES, body);
      case 'deletar':          return deletar(body);
      case 'atualizarStatus':  return atualizarStatus(body);
      case 'enviarMsg':        return enviarMsg(body);
      case 'listarMsgs':       return listarMsgs(e.parameter.sala || body.sala);
      case 'sincronizar':      return sincronizar();  // retorna TUDO de uma vez
      default:                 return { ok: false, erro: 'Ação desconhecida: ' + action };
    }
  } catch(err) {
    return { ok: false, erro: err.message, stack: err.stack };
  }
}

function responder(data) {
  const saida = ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
  return saida;
}

// ═══════════════════════════════════════════════════════════════════
//  SETUP — cria abas e cabeçalhos automaticamente
// ═══════════════════════════════════════════════════════════════════
function setupPlanilha() {
  getOuCriarAba(ABA_DOADORES,  CAB_DOADORES);
  getOuCriarAba(ABA_PACIENTES, CAB_PACIENTES);
  getOuCriarAba(ABA_CHAT,      CAB_CHAT);
  return { ok: true, msg: 'Planilha configurada com sucesso!' };
}

function getOuCriarAba(nome, cabecalhos) {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  let   sheet = ss.getSheetByName(nome);

  if (!sheet) {
    sheet = ss.insertSheet(nome);
    const hRange = sheet.getRange(1, 1, 1, cabecalhos.length);
    hRange.setValues([cabecalhos])
          .setBackground('#c40000')
          .setFontColor('#ffffff')
          .setFontWeight('bold')
          .setFontSize(10);
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, cabecalhos.length, 160);
  }
  return sheet;
}

function gerarId()     { return String(Date.now()) + '_' + Math.random().toString(36).slice(2,5); }
function gerarCodigo() {
  const c = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  return Array.from({length:6}, () => c[Math.floor(Math.random()*c.length)]).join('');
}
function agora() {
  return Utilities.formatDate(new Date(), 'America/Manaus', 'dd/MM/yyyy HH:mm');
}

function zebraLinha(sheet, linha, numCols) {
  const cor = linha % 2 === 0 ? '#fff5f5' : '#ffffff';
  sheet.getRange(linha, 1, 1, numCols).setBackground(cor);
}

// ═══════════════════════════════════════════════════════════════════
//  DOADORES
// ═══════════════════════════════════════════════════════════════════
function addDoador(d) {
  if (!d.nome)  return { ok: false, erro: 'Nome é obrigatório' };
  if (!d.tipo)  return { ok: false, erro: 'Tipo sanguíneo é obrigatório' };

  const sheet  = getOuCriarAba(ABA_DOADORES, CAB_DOADORES);
  const codigo = gerarCodigo();
  const id     = gerarId();

  sheet.appendRow([id, codigo, d.nome, d.tipo, d.data||'', d.cidade||'', d.estado||'', d.apto||'sim', agora()]);
  zebraLinha(sheet, sheet.getLastRow(), CAB_DOADORES.length);

  return { ok: true, id, codigo };
}

function listarDoadores() {
  const sheet = getOuCriarAba(ABA_DOADORES, CAB_DOADORES);
  const dados = sheet.getDataRange().getValues();
  if (dados.length <= 1) return { ok: true, lista: [] };

  const hoje  = new Date();
  const lista = [];

  for (let i = 1; i < dados.length; i++) {
    const r    = dados[i];
    const apto = String(r[7]).toLowerCase();
    if (apto !== 'sim') continue;

    // Verifica intervalo de 90 dias
    const ultimaDoa = r[4];
    if (ultimaDoa) {
      const dt   = new Date(ultimaDoa);
      const diff = (hoje - dt) / 86400000;
      if (!isNaN(dt) && diff < 90) continue;
    }

    lista.push({ id:r[0], codigo:r[1], nome:r[2], tipo:r[3], data:r[4], cidade:r[5], estado:r[6], apto:r[7], cadastro:r[8] });
  }
  return { ok: true, lista };
}

// ═══════════════════════════════════════════════════════════════════
//  PACIENTES
// ═══════════════════════════════════════════════════════════════════
function addPaciente(p) {
  if (!p.nome)     return { ok: false, erro: 'Nome é obrigatório' };
  if (!p.hospital) return { ok: false, erro: 'Hospital é obrigatório' };

  const sheet  = getOuCriarAba(ABA_PACIENTES, CAB_PACIENTES);
  const codigo = gerarCodigo();
  const id     = gerarId();

  sheet.appendRow([id, codigo, p.nome, p.hospital, p.registro||'', p.tipo||'', p.qtd||'', p.obs||'', 'precisa', agora()]);
  zebraLinha(sheet, sheet.getLastRow(), CAB_PACIENTES.length);

  return { ok: true, id, codigo };
}

function listarPacientes() {
  const sheet = getOuCriarAba(ABA_PACIENTES, CAB_PACIENTES);
  const dados = sheet.getDataRange().getValues();
  if (dados.length <= 1) return { ok: true, lista: [] };

  const lista = [];
  for (let i = 1; i < dados.length; i++) {
    const r = dados[i];
    if (String(r[8]).toLowerCase() !== 'precisa') continue;
    lista.push({ id:r[0], codigo:r[1], nome:r[2], hospital:r[3], registro:r[4], tipo:r[5], qtd:r[6], obs:r[7], status:r[8], cadastro:r[9] });
  }
  return { ok: true, lista };
}

// ═══════════════════════════════════════════════════════════════════
//  EDITAR (bidirecional — permite alterar qualquer campo pelo app)
// ═══════════════════════════════════════════════════════════════════
function editarRegistro(aba, cabs, body) {
  const { id, campos } = body;
  if (!id || !campos) return { ok: false, erro: 'id e campos são obrigatórios' };

  const sheet = getOuCriarAba(aba, cabs);
  const dados = sheet.getDataRange().getValues();

  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][0]) === String(id)) {
      // Atualiza cada campo enviado
      Object.keys(campos).forEach(campo => {
        const col = cabs.indexOf(campo);
        if (col >= 0) sheet.getRange(i + 1, col + 1).setValue(campos[campo]);
      });
      return { ok: true };
    }
  }
  return { ok: false, erro: 'Registro não encontrado' };
}

// ═══════════════════════════════════════════════════════════════════
//  ATUALIZAR STATUS
// ═══════════════════════════════════════════════════════════════════
function atualizarStatus(body) {
  const { colecao, id, novoStatus } = body;
  const aba  = colecao === 'doadores' ? ABA_DOADORES  : ABA_PACIENTES;
  const cabs = colecao === 'doadores' ? CAB_DOADORES  : CAB_PACIENTES;
  const colS = colecao === 'doadores' ? 8 : 9; // coluna Status (1-based)

  const sheet = getOuCriarAba(aba, cabs);
  const dados = sheet.getDataRange().getValues();

  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][0]) === String(id)) {
      sheet.getRange(i + 1, colS).setValue(novoStatus);
      return { ok: true };
    }
  }
  return { ok: false, erro: 'Não encontrado' };
}

// ═══════════════════════════════════════════════════════════════════
//  DELETAR
// ═══════════════════════════════════════════════════════════════════
function deletar(body) {
  const { colecao, id } = body;
  const aba  = colecao === 'doadores'  ? ABA_DOADORES
             : colecao === 'pacientes' ? ABA_PACIENTES
             : null;
  if (!aba) return { ok: false, erro: 'Coleção inválida' };

  const cabs  = colecao === 'doadores' ? CAB_DOADORES : CAB_PACIENTES;
  const sheet = getOuCriarAba(aba, cabs);
  const dados = sheet.getDataRange().getValues();

  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, erro: 'Registro não encontrado' };
}

// ═══════════════════════════════════════════════════════════════════
//  CHAT
// ═══════════════════════════════════════════════════════════════════
function enviarMsg(body) {
  const { sala, de, texto, hora } = body;
  if (!sala || !texto) return { ok: false, erro: 'Sala e texto são obrigatórios' };

  const sheet = getOuCriarAba(ABA_CHAT, CAB_CHAT);
  const msgId = gerarId();

  sheet.appendRow([sala, msgId, de||'Anônimo', texto, hora||agora(), agora()]);
  zebraLinha(sheet, sheet.getLastRow(), CAB_CHAT.length);

  return { ok: true, msgId };
}

function listarMsgs(sala) {
  if (!sala) return { ok: false, erro: 'Sala obrigatória' };
  const sheet = getOuCriarAba(ABA_CHAT, CAB_CHAT);
  const dados = sheet.getDataRange().getValues();

  const msgs = [];
  for (let i = 1; i < dados.length; i++) {
    if (String(dados[i][0]) === String(sala)) {
      msgs.push({ msgId:dados[i][1], de:dados[i][2], texto:dados[i][3], hora:dados[i][4] });
    }
  }
  return { ok: true, msgs };
}

// ═══════════════════════════════════════════════════════════════════
//  SINCRONIZAR — retorna doadores + pacientes + últimas msgs em 1 chamada
// ═══════════════════════════════════════════════════════════════════
function sincronizar() {
  const doadores  = listarDoadores();
  const pacientes = listarPacientes();
  return {
    ok: true,
    doadores:  doadores.lista  || [],
    pacientes: pacientes.lista || [],
    timestamp: new Date().getTime()
  };
}

// ═══════════════════════════════════════════════════════════════════
//  TESTE — rode manualmente no editor para verificar
// ═══════════════════════════════════════════════════════════════════
function testar() {
  Logger.log('=== SETUP ===');
  Logger.log(JSON.stringify(setupPlanilha()));

  Logger.log('=== ADD DOADOR ===');
  Logger.log(JSON.stringify(addDoador({ nome:'João Silva', tipo:'O+', cidade:'Manaus', estado:'AM', apto:'sim', data:'2024-09-01' })));

  Logger.log('=== ADD PACIENTE ===');
  Logger.log(JSON.stringify(addPaciente({ nome:'Maria Santos', hospital:'HPS Getúlio Vargas', tipo:'O+', qtd:'2', obs:'Urgente' })));

  Logger.log('=== LISTAR DOADORES ===');
  Logger.log(JSON.stringify(listarDoadores()));

  Logger.log('=== LISTAR PACIENTES ===');
  Logger.log(JSON.stringify(listarPacientes()));

  Logger.log('=== SYNC ===');
  Logger.log(JSON.stringify(sincronizar()));
}
