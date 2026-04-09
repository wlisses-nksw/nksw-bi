// =====================================================================
// NKSW BI — Sistema de Autenticação
// =====================================================================
// INSTRUÇÕES DE CONFIGURAÇÃO:
//   1. Crie um novo Google Apps Script (script.google.com)
//   2. Cole este código
//   3. Execute initSheet() uma vez para criar as planilhas
//   4. Execute setupFirstAdmin() para criar o primeiro usuário admin
//   5. Implante como Web App:
//        - Executar como: Eu (Me)
//        - Quem tem acesso: Qualquer pessoa (Anyone)
//   6. Copie a URL gerada e cole no dashboard (botão "Configurar URL")
// =====================================================================

var SHEET_NAME_USERS  = 'Usuarios';
var SHEET_NAME_TOKENS = 'Tokens';
var TOKEN_EXPIRY_HOURS = 24;

// ── Entrada principal ────────────────────────────────────────────────
function doPost(e) {
  var result;
  try {
    var params = JSON.parse(e.postData.contents);
    var action = params.action;

    if      (action === 'login')         result = login(params.email, params.senhaHash);
    else if (action === 'listUsers')     result = listUsers(params.token);
    else if (action === 'createUser')    result = createUser(params.token, params.usuario);
    else if (action === 'updateUser')    result = updateUser(params.token, params.usuario);
    else if (action === 'resetPassword') result = resetPassword(params.token, params.id, params.senhaHash);
    else                                 result = { ok: false, msg: 'Ação desconhecida: ' + action };
  } catch (err) {
    result = { ok: false, msg: err.message };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Login ────────────────────────────────────────────────────────────
function login(email, senhaHash) {
  if (!email || !senhaHash) return { ok: false, msg: 'Email e senha obrigatórios' };

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_USERS);
  if (!sheet) return { ok: false, msg: 'Planilha de usuários não encontrada. Execute initSheet().' };

  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  var iEmail  = headers.indexOf('email');
  var iHash   = headers.indexOf('senha_hash');
  var iAtivo  = headers.indexOf('ativo');
  var iNome   = headers.indexOf('nome');
  var iRole   = headers.indexOf('role');
  var iTabs   = headers.indexOf('tabs');
  var iAccess = headers.indexOf('ultimo_acesso');

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (String(row[iEmail]).toLowerCase() !== email.toLowerCase()) continue;

    var ativo = row[iAtivo];
    if (ativo !== true && ativo !== 'TRUE' && ativo !== 'true') {
      return { ok: false, msg: 'Usuário inativo. Contate o administrador.' };
    }
    if (row[iHash] !== senhaHash) {
      return { ok: false, msg: 'Email ou senha incorretos.' };
    }

    // Gera token
    var token  = Utilities.getUuid();
    var expiry = new Date(Date.now() + TOKEN_EXPIRY_HOURS * 3600 * 1000).toISOString();
    getOrCreateSheet(SHEET_NAME_TOKENS).appendRow([token, email, expiry]);

    // Atualiza último acesso
    if (iAccess >= 0) sheet.getRange(i + 1, iAccess + 1).setValue(new Date().toISOString());

    var tabs = String(row[iTabs] || '').split(',').map(function(t) { return t.trim(); }).filter(Boolean);

    return {
      ok:    true,
      token: token,
      nome:  row[iNome],
      email: email.toLowerCase(),
      role:  row[iRole],
      tabs:  tabs
    };
  }

  return { ok: false, msg: 'Usuário não encontrado.' };
}

// ── Validar token ────────────────────────────────────────────────────
function validateToken(token) {
  if (!token) return { ok: false, msg: 'Token ausente' };

  var sheet = getOrCreateSheet(SHEET_NAME_TOKENS);
  var data  = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] !== token) continue;
    if (new Date(data[i][2]) <= new Date()) return { ok: false, msg: 'Sessão expirada. Faça login novamente.' };

    var email    = data[i][1];
    var userData = getUserByEmail(email);
    if (!userData) return { ok: false, msg: 'Usuário não encontrado' };

    var ativo = userData.ativo;
    if (ativo !== true && ativo !== 'TRUE' && ativo !== 'true') {
      return { ok: false, msg: 'Usuário inativo' };
    }

    return { ok: true, email: email, nome: userData.nome, role: userData.role, tabs: userData.tabs };
  }

  return { ok: false, msg: 'Token inválido' };
}

function requireAdmin(token) {
  var v = validateToken(token);
  if (!v.ok) throw new Error(v.msg || 'Não autenticado');
  if (v.role !== 'admin') throw new Error('Acesso restrito a administradores');
  return v;
}

// ── Listar usuários ──────────────────────────────────────────────────
function listUsers(token) {
  requireAdmin(token);

  var sheet   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_USERS);
  var data    = sheet.getDataRange().getValues();
  var headers = data[0];

  var users = [];
  for (var i = 1; i < data.length; i++) {
    var row   = data[i];
    var ativo = row[headers.indexOf('ativo')];
    users.push({
      id:        row[headers.indexOf('id')],
      nome:      row[headers.indexOf('nome')],
      email:     row[headers.indexOf('email')],
      tabs:      String(row[headers.indexOf('tabs')] || '').split(',').map(function(t) { return t.trim(); }).filter(Boolean),
      role:      row[headers.indexOf('role')],
      ativo:     ativo === true || ativo === 'TRUE' || ativo === 'true',
      criado_em: row[headers.indexOf('criado_em')]
    });
  }

  return { ok: true, users: users };
}

// ── Criar usuário ────────────────────────────────────────────────────
function createUser(token, usuario) {
  requireAdmin(token);

  var ss      = SpreadsheetApp.getActiveSpreadsheet();
  var sheet   = ss.getSheetByName(SHEET_NAME_USERS);
  var data    = sheet.getDataRange().getValues();
  var headers = data[0];

  // Verifica duplicata
  var iEmail = headers.indexOf('email');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][iEmail]).toLowerCase() === usuario.email.toLowerCase()) {
      return { ok: false, msg: 'Email já cadastrado.' };
    }
  }

  var id  = Utilities.getUuid();
  var row = headers.map(function(h) {
    if (h === 'id')         return id;
    if (h === 'nome')       return usuario.nome;
    if (h === 'email')      return usuario.email.toLowerCase();
    if (h === 'senha_hash') return usuario.senhaHash;
    if (h === 'tabs')       return Array.isArray(usuario.tabs) ? usuario.tabs.join(',') : (usuario.tabs || '');
    if (h === 'role')       return usuario.role || 'viewer';
    if (h === 'ativo')      return true;
    if (h === 'criado_em')  return new Date().toISOString();
    return '';
  });

  sheet.appendRow(row);
  return { ok: true, id: id };
}

// ── Atualizar usuário ────────────────────────────────────────────────
function updateUser(token, usuario) {
  requireAdmin(token);

  var sheet   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_USERS);
  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  var iId     = headers.indexOf('id');

  for (var i = 1; i < data.length; i++) {
    if (data[i][iId] !== usuario.id) continue;

    var set = function(field, value) {
      var col = headers.indexOf(field);
      if (col >= 0 && value !== undefined) sheet.getRange(i + 1, col + 1).setValue(value);
    };

    set('nome',  usuario.nome);
    set('role',  usuario.role);
    set('ativo', usuario.ativo);
    set('tabs',  Array.isArray(usuario.tabs) ? usuario.tabs.join(',') : usuario.tabs);

    return { ok: true };
  }

  return { ok: false, msg: 'Usuário não encontrado.' };
}

// ── Redefinir senha ──────────────────────────────────────────────────
function resetPassword(token, id, senhaHash) {
  requireAdmin(token);

  var sheet   = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_USERS);
  var data    = sheet.getDataRange().getValues();
  var headers = data[0];
  var iId     = headers.indexOf('id');
  var iHash   = headers.indexOf('senha_hash');

  for (var i = 1; i < data.length; i++) {
    if (data[i][iId] !== id) continue;
    sheet.getRange(i + 1, iHash + 1).setValue(senhaHash);
    return { ok: true };
  }

  return { ok: false, msg: 'Usuário não encontrado.' };
}

// ── Helpers ──────────────────────────────────────────────────────────
function getUserByEmail(email) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_USERS);
  if (!sheet) return null;
  var data    = sheet.getDataRange().getValues();
  var headers = data[0];

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][headers.indexOf('email')]).toLowerCase() !== email.toLowerCase()) continue;
    var tabs = String(data[i][headers.indexOf('tabs')] || '').split(',').map(function(t) { return t.trim(); }).filter(Boolean);
    return {
      nome:  data[i][headers.indexOf('nome')],
      role:  data[i][headers.indexOf('role')],
      tabs:  tabs,
      ativo: data[i][headers.indexOf('ativo')]
    };
  }
  return null;
}

function getOrCreateSheet(name) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (name === SHEET_NAME_USERS) {
      sheet.getRange(1, 1, 1, 9).setValues([[
        'id', 'nome', 'email', 'senha_hash', 'tabs', 'role', 'ativo', 'criado_em', 'ultimo_acesso'
      ]]);
      sheet.setFrozenRows(1);
    } else if (name === SHEET_NAME_TOKENS) {
      sheet.getRange(1, 1, 1, 3).setValues([['token', 'email', 'expiry']]);
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

// ── Aquecimento — evita cold start (configure gatilho a cada 10 min) ─
// No editor: Gatilhos (⏰) → Novo gatilho → warmUp → Por tempo → A cada 10 minutos
function warmUp() {
  // Mantém o script "quente" — a primeira requisição real chega instantânea
  CacheService.getScriptCache().put('warmup', new Date().toISOString(), 60);
}

// ── SHA-256 (usado apenas no setup pelo editor) ──────────────────────
function sha256(str) {
  var bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256, str, Utilities.Charset.UTF_8
  );
  return bytes.map(function(b) {
    return ('0' + (b < 0 ? b + 256 : b).toString(16)).slice(-2);
  }).join('');
}

// ── Setup inicial (execute uma vez pelo editor) ──────────────────────
function initSheet() {
  getOrCreateSheet(SHEET_NAME_USERS);
  getOrCreateSheet(SHEET_NAME_TOKENS);
  SpreadsheetApp.getUi().alert('Planilhas criadas! Execute setupFirstAdmin() para criar o admin.');
}

function setupFirstAdmin() {
  var ui     = SpreadsheetApp.getUi();
  var sheet  = getOrCreateSheet(SHEET_NAME_USERS);
  var data   = sheet.getDataRange().getValues();

  if (data.length > 1) {
    ui.alert('Já existem usuários cadastrados. Use o painel admin no dashboard para gerenciá-los.');
    return;
  }

  var nome  = ui.prompt('Nome do administrador:').getResponseText().trim();
  var email = ui.prompt('Email do administrador:').getResponseText().trim().toLowerCase();
  var senha = ui.prompt('Senha (será convertida em hash SHA-256):').getResponseText();

  if (!nome || !email || !senha) { ui.alert('Operação cancelada.'); return; }

  var allTabs = 'indicadores,marketing,webanalytics,vendas,produtos,pedidos,clientes,logistica,metas';

  sheet.appendRow([
    Utilities.getUuid(),
    nome,
    email,
    sha256(senha),
    allTabs,
    'admin',
    true,
    new Date().toISOString(),
    ''
  ]);

  var url = ScriptApp.getService().getUrl();
  ui.alert('Admin criado!\n\nURL do Web App:\n' + url + '\n\nCole essa URL no dashboard em "Configurar URL".');
}

// ── Limpeza de tokens expirados (opcional — configure como gatilho diário) ─
function cleanExpiredTokens() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_TOKENS);
  if (!sheet) return;
  var data = sheet.getDataRange().getValues();
  var now  = new Date();
  for (var i = data.length - 1; i >= 1; i--) {
    if (new Date(data[i][2]) <= now) sheet.deleteRow(i + 1);
  }
}
