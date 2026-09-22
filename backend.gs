/**
 * BACKEND — Procedimentos de Trabalho
 * -------------------------------------------------------------------------
 * Cole este código no Apps Script de uma Planilha Google (Extensões >
 * Apps Script), salve, e implante como Web App (Implantar > Nova
 * implantação > Tipo: App da Web):
 *   - Executar como: Eu (sua conta)
 *   - Quem tem acesso: Qualquer pessoa
 * Copie a URL do Web App gerada e cole no arquivo HTML, na constante
 * API_URL (procure por "COLE_AQUI_A_URL_DO_WEB_APP").
 *
 * Este script cria automaticamente, na planilha ativa, duas abas:
 *   - "Procedimentos": um registro por procedimento (estado atual)
 *   - "Historico": um registro por gravação (todas as versões salvas)
 * E, no Google Drive, uma pasta raiz "Procedimentos de Trabalho - Fotos"
 * com uma subpasta por procedimento, para onde as fotos são enviadas.
 * -------------------------------------------------------------------------
 */

const SHEET_PROCEDIMENTOS = 'Procedimentos';
const SHEET_HISTORICO = 'Historico';
const DRIVE_ROOT_FOLDER_NAME = 'Procedimentos de Trabalho - Fotos';

const COL_PROC = ['ID', 'Titulo', 'Setor', 'Elaborado', 'Data', 'RevisaoAtual', 'AtualizadoEm', 'DriveFolderId', 'DadosJSON'];
const COL_HIST = ['ProcedimentoID', 'Revisao', 'SalvoEm', 'DadosJSON'];

// ============================= UTIL =========================================

function getOrCreateSheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.setFrozenRows(1);
  }
  return sh;
}

function sheetProc_() { return getOrCreateSheet_(SHEET_PROCEDIMENTOS, COL_PROC); }
function sheetHist_() { return getOrCreateSheet_(SHEET_HISTORICO, COL_HIST); }

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function findRowById_(sh, id) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return -1; // só tem cabeçalho, nenhuma linha de dado ainda
  const values = sh.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) return i + 2; // linha real na planilha
  }
  return -1;
}

// ============================= ROTEAMENTO ===================================

function doGet(e) {
  try {
    const action = e.parameter.action;
    let result;
    if (action === 'listProcedures') result = listProcedures();
    else if (action === 'getProcedure') result = getProcedure(e.parameter.id);
    else if (action === 'getHistory') result = getHistory(e.parameter.id);
    else if (action === 'getHistoryVersion') result = getHistoryVersion(e.parameter.id, e.parameter.row);
    else if (action === 'ping') result = { ok: true, hora: new Date().toISOString() };
    else result = { error: 'Ação GET desconhecida: ' + action };
    return jsonOut_(result);
  } catch (err) {
    return jsonOut_({ error: String(err) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    let result;
    if (action === 'saveProcedure') result = saveProcedure(body.data);
    else if (action === 'uploadPhoto') result = uploadPhoto(body.data);
    else if (action === 'deleteProcedure') result = deleteProcedure(body.data.id);
    else result = { error: 'Ação POST desconhecida: ' + action };
    return jsonOut_(result);
  } catch (err) {
    return jsonOut_({ error: String(err) });
  }
}

// ============================= PROCEDIMENTOS ================================

function listProcedures() {
  const sh = sheetProc_();
  const last = sh.getLastRow();
  if (last < 2) return { procedimentos: [] };
  const rows = sh.getRange(2, 1, last - 1, 9).getValues();
  const out = [];
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0]) continue;
    out.push({
      id: r[0], titulo: r[1], setor: r[2], elaborado: r[3],
      data: r[4], revisaoAtual: r[5], atualizadoEm: r[6]
    });
  }
  out.sort(function (a, b) { return String(b.atualizadoEm).localeCompare(String(a.atualizadoEm)); });
  return { procedimentos: out };
}

function getProcedure(id) {
  const sh = sheetProc_();
  const row = findRowById_(sh, id);
  if (row === -1) return { error: 'Procedimento não encontrado' };
  const vals = sh.getRange(row, 1, 1, 9).getValues()[0];
  return {
    id: vals[0], titulo: vals[1], setor: vals[2], elaborado: vals[3],
    data: vals[4], revisaoAtual: vals[5], atualizadoEm: vals[6],
    driveFolderId: vals[7], conteudo: JSON.parse(vals[8] || '{}')
  };
}

function saveProcedure(data) {
  // data: { id (vazio = novo), titulo, setor, elaborado, data, revisaoAtual, conteudo:{...} }
  const sh = sheetProc_();
  const now = new Date().toISOString();
  const json = JSON.stringify(data.conteudo || {});
  let id = data.id;

  if (!id) {
    id = 'PROC-' + new Date().getTime();
    sh.appendRow([id, data.titulo || '', data.setor || '', data.elaborado || '', data.data || '', data.revisaoAtual || '00', now, '', json]);
  } else {
    const row = findRowById_(sh, id);
    if (row === -1) {
      sh.appendRow([id, data.titulo || '', data.setor || '', data.elaborado || '', data.data || '', data.revisaoAtual || '00', now, '', json]);
    } else {
      const folderId = sh.getRange(row, 8).getValue();
      sh.getRange(row, 2, 1, 8).setValues([[
        data.titulo || '', data.setor || '', data.elaborado || '', data.data || '',
        data.revisaoAtual || '00', now, folderId, json
      ]]);
    }
  }

  sheetHist_().appendRow([id, data.revisaoAtual || '00', now, json]);
  return { id: id, atualizadoEm: now };
}

function deleteProcedure(id) {
  const sh = sheetProc_();
  const row = findRowById_(sh, id);
  if (row !== -1) sh.deleteRow(row);
  return { ok: true };
}

// ============================= HISTÓRICO =====================================

function getHistory(id) {
  const sh = sheetHist_();
  const last = sh.getLastRow();
  if (last < 2) return { historico: [] };
  const rows = sh.getRange(2, 1, last - 1, 4).getValues();
  const out = [];
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      out.push({ row: i + 2, revisao: rows[i][1], salvoEm: rows[i][2] });
    }
  }
  out.reverse(); // mais recente primeiro
  return { historico: out };
}

function getHistoryVersion(id, rowNumber) {
  const sh = sheetHist_();
  const vals = sh.getRange(Number(rowNumber), 1, 1, 4).getValues()[0];
  if (String(vals[0]) !== String(id)) return { error: 'Versão não encontrada' };
  return { revisao: vals[1], salvoEm: vals[2], conteudo: JSON.parse(vals[3] || '{}') };
}

// ============================= FOTOS (DRIVE) =================================

function ensureRootFolder_() {
  const it = DriveApp.getFoldersByName(DRIVE_ROOT_FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(DRIVE_ROOT_FOLDER_NAME);
}

function ensureProcedureFolder_(procedureId) {
  const root = ensureRootFolder_();
  const it = root.getFoldersByName(procedureId);
  return it.hasNext() ? it.next() : root.createFolder(procedureId);
}

function uploadPhoto(data) {
  // data: { procedureId, filename, mimeType, base64 }
  if (!data.procedureId) return { error: 'procedureId ausente' };
  const folder = ensureProcedureFolder_(data.procedureId);
  const bytes = Utilities.base64Decode(data.base64);
  const blob = Utilities.newBlob(bytes, data.mimeType || 'image/jpeg', data.filename || 'foto.jpg');
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return {
    fileId: file.getId(),
    url: 'https://drive.google.com/uc?export=view&id=' + file.getId()
  };
}
