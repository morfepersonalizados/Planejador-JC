// ============================================================
// NOSSO DINHEIRO - Google Apps Script (Backend API)
// Cole este codigo no Apps Script da planilha "Nosso dinheiro"
// (Extensoes > Apps Script) e faca o deploy como Web App.
// Depois de colar, configure em Configuracoes do projeto > Propriedades
// do script (Script Properties):
//   GOOGLE_CLIENT_ID   = 866951819808-tqeoto2rqnr34p7250mmq6l7mn93afs2.apps.googleusercontent.com
//   AUTHORIZED_EMAILS  = seuemail@gmail.com,emaildela@gmail.com
// (o mesmo Client ID do ERP da loja pode ser reaproveitado com seguranca:
// ele so identifica o APP, quem autoriza de verdade e este backend,
// comparando o email do token com AUTHORIZED_EMAILS)
// ============================================================

// Tabelas do sistema
var TABELAS = [
  'config', 'entradas', 'despesas_fixas', 'despesas', 'potes', 'aportes',
  'divisoes', 'parcelamentos', 'parcelas'
];

// Rode esta funcao manualmente (selecione "autorizarPermissoes" no menu
// suspenso ao lado de Executar) sempre que o app mostrar "Offline" e o
// erro mencionar "UrlFetchApp" ou "external_request" — isso forca o
// Google a pedir a autorizacao que a validacao do login com Google
// precisa (doGet sozinho nao serve pra isso: ele falha antes de chegar
// nessa parte do codigo, entao nunca aciona o pedido de permissao).
function autorizarPermissoes() {
  UrlFetchApp.fetch('https://www.google.com', { muteHttpExceptions: true });
}

// ============================================================
// AUTENTICACAO - so login com Google, validado aqui no servidor. Sem
// senha fixa no codigo (o index.html e publico, entao uma senha ali
// ficaria visivel pra qualquer um que abrisse "Ver codigo-fonte").
// ============================================================
function validarAcesso(params) {
  if (params.token) {
    var resultado = validarTokenGoogle(params.token);
    if (resultado.ok) return { ok: true, metodo: 'google', email: resultado.email };
    return { ok: false, erro: resultado.erro };
  }
  return { ok: false, erro: 'Nao autenticado' };
}

// Valida um token de login do Google (ID token) chamando o endpoint oficial
// do Google, que confirma a assinatura do token e devolve os dados nele.
// So aceita se: o token for valido e nao estiver vencido, o "audience"
// (client ID) bater com o configurado, e o email estiver na lista de
// autorizados (AUTHORIZED_EMAILS, separados por virgula).
function validarTokenGoogle(idToken) {
  var props = PropertiesService.getScriptProperties();
  var clientIdEsperado = props.getProperty('GOOGLE_CLIENT_ID');
  var emailsAutorizados = (props.getProperty('AUTHORIZED_EMAILS') || '')
    .split(',').map(function(e){ return e.trim().toLowerCase(); }).filter(String);

  if (!clientIdEsperado || emailsAutorizados.length === 0) {
    return { ok: false, erro: 'Login com Google ainda nao foi configurado no backend (faltam as Script Properties GOOGLE_CLIENT_ID e AUTHORIZED_EMAILS)' };
  }

  try {
    var resposta = UrlFetchApp.fetch(
      'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken),
      { muteHttpExceptions: true }
    );
    var dados = JSON.parse(resposta.getContentText());

    if (dados.error) return { ok: false, erro: 'Token invalido ou expirado, faca login novamente' };
    if (dados.aud !== clientIdEsperado) return { ok: false, erro: 'Token nao pertence a este aplicativo' };
    if (dados.email_verified !== 'true' && dados.email_verified !== true) return { ok: false, erro: 'Email do Google nao verificado' };
    if (emailsAutorizados.indexOf(String(dados.email).toLowerCase()) === -1) return { ok: false, erro: 'Este email do Google nao tem acesso a este sistema' };

    return { ok: true, email: dados.email };
  } catch (err) {
    return { ok: false, erro: 'Erro ao validar login do Google: ' + err.toString() };
  }
}

// GET: Ler dados
//
// Usa o mesmo LockService do doPost. Sem essa trava, uma leitura podia
// cair no meio de uma gravacao (salvarTabela faz clearContents() e SO
// DEPOIS setValues() - nao e atomico) e ver a aba momentaneamente vazia.
function doGet(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (err) {
    return jsonResponse({ ok: false, erro: 'Sistema ocupado gravando outra coisa, tente de novo em alguns segundos' });
  }
  try {
    var params = e.parameter;

    var acesso = validarAcesso(params);
    if (!acesso.ok) {
      return jsonResponse({ ok: false, erro: acesso.erro || 'Nao autorizado' });
    }

    var acao = params.acao || 'ler_tudo';
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    if (acao === 'ler_tudo') {
      var dados = {};
      for (var i = 0; i < TABELAS.length; i++) {
        dados[TABELAS[i]] = lerTabela(ss, TABELAS[i]);
      }
      return jsonResponse({ ok: true, dados: dados });
    }

    if (acao === 'ler') {
      var tabela = params.tabela;
      if (!tabela) return jsonResponse({ ok: false, erro: 'Tabela nao informada' });
      return jsonResponse({ ok: true, dados: lerTabela(ss, tabela) });
    }

    return jsonResponse({ ok: false, erro: 'Acao nao reconhecida' });

  } catch (err) {
    return jsonResponse({ ok: false, erro: err.toString() });
  } finally {
    lock.releaseLock();
  }
}

// POST: Salvar dados
//
// A trava evita que dois salvamentos quase simultaneos (ex: os dois
// aparelhos do casal salvando ao mesmo tempo) leiam a aba no mesmo
// instante, cada um calcule em cima do mesmo valor antigo, e um
// sobrescreva o outro por inteiro ao gravar.
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (err) {
    return jsonResponse({ ok: false, erro: 'Sistema ocupado gravando outra coisa, tente de novo em alguns segundos' });
  }

  try {
    var body = JSON.parse(e.postData.contents);

    var acesso = validarAcesso(body);
    if (!acesso.ok) {
      return jsonResponse({ ok: false, erro: acesso.erro || 'Nao autorizado' });
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var acao = body.acao;
    var tabela = body.tabela;
    var dados = body.dados;
    var id = body.id;

    if (acao === 'salvar_tudo') {
      for (var i = 0; i < TABELAS.length; i++) {
        var t = TABELAS[i];
        if (dados[t] !== undefined) {
          salvarTabela(ss, t, dados[t]);
        }
      }
      return jsonResponse({ ok: true, msg: 'Sync completo realizado' });
    }

    if (acao === 'upsert') {
      if (!tabela || !dados) return jsonResponse({ ok: false, erro: 'Dados incompletos' });
      upsertRegistro(ss, tabela, dados);
      return jsonResponse({ ok: true, msg: 'Salvo com sucesso' });
    }

    if (acao === 'excluir') {
      if (!tabela || !id) return jsonResponse({ ok: false, erro: 'Dados incompletos' });
      excluirRegistro(ss, tabela, id);
      return jsonResponse({ ok: true, msg: 'Excluido com sucesso' });
    }

    if (acao === 'salvar_tabela') {
      if (!tabela || !dados) return jsonResponse({ ok: false, erro: 'Dados incompletos' });
      salvarTabela(ss, tabela, dados);
      return jsonResponse({ ok: true, msg: 'Tabela salva' });
    }

    return jsonResponse({ ok: false, erro: 'Acao nao reconhecida' });

  } catch (err) {
    return jsonResponse({ ok: false, erro: err.toString() });
  } finally {
    lock.releaseLock();
  }
}

// Helpers (identicos ao motor generico ja usado e testado no ERP da loja)
function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet(ss, nome) {
  var sheet = ss.getSheetByName(nome);
  if (!sheet) {
    sheet = ss.insertSheet(nome);
  }
  return sheet;
}

function lerTabela(ss, tabela) {
  var sheet = ss.getSheetByName(tabela);
  if (!sheet) return [];

  var dados = sheet.getDataRange().getValues();
  if (dados.length < 2) return [];

  var headers = dados[0];
  var resultado = [];
  var fuso = ss.getSpreadsheetTimeZone();

  for (var r = 1; r < dados.length; r++) {
    var row = dados[r];
    if (row[0] === '') continue;
    var obj = {};
    for (var c = 0; c < headers.length; c++) {
      var val = row[c];
      if (Object.prototype.toString.call(val) === '[object Date]') {
        val = Utilities.formatDate(val, fuso, 'yyyy-MM-dd');
      } else if (typeof val === 'string' && (val.charAt(0) === '[' || val.charAt(0) === '{')) {
        try { val = JSON.parse(val); } catch(e) {}
      }
      obj[headers[c]] = val;
    }
    resultado.push(obj);
  }
  return resultado;
}

function salvarTabela(ss, tabela, registros) {
  var sheet = getOrCreateSheet(ss, tabela);
  sheet.clearContents();

  if (!registros || registros.length === 0) return;

  var camposSet = {};
  var headers = [];
  for (var i = 0; i < registros.length; i++) {
    var keys = Object.keys(registros[i]);
    for (var k = 0; k < keys.length; k++) {
      if (!camposSet[keys[k]]) {
        camposSet[keys[k]] = true;
        headers.push(keys[k]);
      }
    }
  }

  var rows = [headers];

  for (var i = 0; i < registros.length; i++) {
    var row = [];
    for (var h = 0; h < headers.length; h++) {
      var val = registros[i][headers[h]];
      if (Array.isArray(val) || (typeof val === 'object' && val !== null)) {
        row.push(JSON.stringify(val));
      } else {
        row.push(val !== undefined ? val : '');
      }
    }
    rows.push(row);
  }

  var range = sheet.getRange(1, 1, rows.length, headers.length);
  range.setNumberFormat('@');
  range.setValues(rows);
}

function upsertRegistro(ss, tabela, registro) {
  var registros = lerTabela(ss, tabela);
  var idx = -1;
  for (var i = 0; i < registros.length; i++) {
    if (registros[i].id === registro.id) { idx = i; break; }
  }
  if (idx >= 0) {
    registros[idx] = registro;
  } else {
    registros.push(registro);
  }
  salvarTabela(ss, tabela, registros);
}

function excluirRegistro(ss, tabela, id) {
  var registros = lerTabela(ss, tabela);
  var novos = [];
  for (var i = 0; i < registros.length; i++) {
    if (registros[i].id !== id) novos.push(registros[i]);
  }
  salvarTabela(ss, tabela, novos);
}
