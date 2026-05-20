// ============================================================
//  CETIPE — Backend Google Apps Script
//  Cole este código em: script.google.com → Novo projeto
//  Depois: Implantar → Nova implantação → Aplicativo da Web
//    Executar como: Eu mesmo
//    Quem pode acessar: Qualquer pessoa
//  Copie a URL gerada e cole em index.html → APPS_SCRIPT_URL
// ============================================================

// ── Senha de acesso (mude para algo só você sabe) ─────────────
var SENHA = "1567";

// ── IDs das planilhas (mesmos do config.py) ──────────────────
var ID_ENTRADA = "1n1i9Cr1C9J1CCsKMF0EntFT0OIVCc3PvyfKcACGQbwA";
var ID_SAIDA   = "1jezbNv1F_De2E1wInTdDxV02EvN1ey9mXTnulnbVsVM";

// ── Nome das abas ─────────────────────────────────────────────
var ABA_ENTRADA       = "App Entrada";
var ABA_SAIDA         = "App Saida";
var ABA_FUNCIONARIOS  = "Funcionarios";

var CABECALHO_ENTRADA = ["Carimbo de data/hora", "Responsavel pelo preenchimento:", "SETOR", "FUNCIONÁRIOS", "DESCREVA O HORARIO DE ENTRADA"];
var CABECALHO_SAIDA   = ["Carimbo de data/hora", "Responsavel pelo preenchimento:", "SETOR", "FUNCIONÁRIOS", "DESCREVA O HORARIO DE SAÍDA"];

// ── GET: retorna lista de funcionários por setor (JSONP) ──────
function doGet(e) {
  e = e || { parameter: {} };  // evita erro ao rodar no editor
  var callback = e.parameter.callback;
  var modo     = e.parameter.modo || "";

  if (e.parameter.senha !== SENHA) {
    if (modo === "iframe") return _htmlMsg({ erro: "Acesso negado." });
    return responder(callback, { erro: "Acesso negado." });
  }

  try {
    var ss  = SpreadsheetApp.openById(ID_ENTRADA);
    var aba = ss.getSheetByName(ABA_FUNCIONARIOS);

    if (!aba) {
      aba = ss.insertSheet(ABA_FUNCIONARIOS);
      aba.appendRow(["Setor", "Nome", "Status"]);
      if (modo === "iframe") return _htmlMsg({ funcionarios: {} });
      return responder(callback, { funcionarios: {} });
    }

    var dados     = aba.getDataRange().getValues();
    var resultado = {};

    for (var i = 1; i < dados.length; i++) {
      var setor  = String(dados[i][0] || "").trim();
      var nome   = String(dados[i][1] || "").trim();
      var status = String(dados[i][2] || "Ativo").trim();
      if (!setor || !nome) continue;
      if (status === "Inativo") continue;
      if (!resultado[setor]) resultado[setor] = [];
      resultado[setor].push(nome);
    }

    if (modo === "iframe") return _htmlMsg({ funcionarios: resultado });
    return responder(callback, { funcionarios: resultado });

  } catch (err) {
    if (modo === "iframe") return _htmlMsg({ erro: err.toString() });
    return responder(callback, { erro: err.toString() });
  }
}

// Retorna HTML que faz postMessage para a janela pai (mobile-friendly)
function _htmlMsg(obj) {
  var json = JSON.stringify(obj);
  return HtmlService.createHtmlOutput(
    '<script>window.parent.postMessage(' + json + ',"*");<\/script>'
  );
}

// ── POST: salva registro de entrada ou saída ──────────────────
function doPost(e) {
  if (e.parameter.senha !== SENHA) {
    return jsonSaida({ erro: "Acesso negado." });
  }

  try {
    var tipo        = e.parameter.tipo;        // "entrada" ou "saida"
    var responsavel = e.parameter.responsavel;
    var setor       = e.parameter.setor;
    var funcionario = e.parameter.funcionario;
    var horario     = e.parameter.horario;

    // Timestamp no formato que o CETIPE espera
    var agora = new Date();
    var timestamp = Utilities.formatDate(agora, "America/Sao_Paulo", "dd/MM/yyyy HH:mm:ss");

    if (tipo === "entrada") {
      var ss  = SpreadsheetApp.openById(ID_ENTRADA);
      var aba = ss.getSheetByName(ABA_ENTRADA);
      if (!aba) {
        aba = ss.insertSheet(ABA_ENTRADA);
        aba.appendRow(CABECALHO_ENTRADA);
      }
      aba.appendRow([timestamp, responsavel, setor, funcionario, horario]);

    } else if (tipo === "saida") {
      var ss2  = SpreadsheetApp.openById(ID_SAIDA);
      var aba2 = ss2.getSheetByName(ABA_SAIDA);
      if (!aba2) {
        aba2 = ss2.insertSheet(ABA_SAIDA);
        aba2.appendRow(CABECALHO_SAIDA);
      }
      aba2.appendRow([timestamp, responsavel, setor, funcionario, horario]);

    } else {
      return jsonSaida({ erro: "Tipo invalido: use 'entrada' ou 'saida'." });
    }

    return jsonSaida({ ok: true, timestamp: timestamp });

  } catch (err) {
    return jsonSaida({ erro: err.toString() });
  }
}

// ── Helpers ───────────────────────────────────────────────────
function responder(callback, obj) {
  if (callback) {
    // JSONP para requisições GET sem CORS
    return ContentService
      .createTextOutput(callback + "(" + JSON.stringify(obj) + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return jsonSaida(obj);
}

function jsonSaida(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
