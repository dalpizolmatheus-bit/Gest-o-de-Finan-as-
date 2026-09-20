/**
 * Gestor Financeiro Pessoal — backend em Google Apps Script.
 *
 * Como publicar:
 *  1. Abra a planilha Google que vai servir de banco de dados.
 *  2. Extensões > Apps Script.
 *  3. Apague o conteúdo de Code.gs e cole este arquivo inteiro.
 *  4. Implantar > Nova implantação > tipo "App da Web".
 *     - Executar como: Eu
 *     - Quem pode acessar: Qualquer pessoa
 *  5. Copie a URL do app da Web gerada e cole no app (tela de Configurações).
 *  6. Abra a URL uma vez no navegador (ou chame ?action=bootstrap) para criar
 *     as abas e cabeçalhos automaticamente, com categorias padrão.
 */

const SHEETS = {
  LANCAMENTOS: "Lancamentos",
  CATEGORIAS: "Categorias",
  PARCELAMENTOS: "Parcelamentos",
  INVESTIMENTOS: "Investimentos",
};

const HEADERS = {
  [SHEETS.LANCAMENTOS]: [
    "id", "data", "tipo", "valor", "macrocategoria", "microcategoria",
    "descricao", "conta", "forma_entrada", "parcelamento_id",
    "comprovante_url", "criado_em",
  ],
  [SHEETS.CATEGORIAS]: ["macrocategoria", "microcategoria", "ativa"],
  [SHEETS.PARCELAMENTOS]: [
    "id", "descricao", "valor_total", "num_parcelas",
    "data_primeira_parcela", "macrocategoria", "criado_em",
  ],
  [SHEETS.INVESTIMENTOS]: [
    "id", "banco", "saldo", "observacao", "criado_em", "atualizado_em",
  ],
};

const DEFAULT_CATEGORIES = [
  ["Alimentação", "Mercado"], ["Alimentação", "Restaurante"], ["Alimentação", "Delivery"],
  ["Saúde", "Consulta"], ["Saúde", "Farmácia"], ["Saúde", "Plano de saúde"],
  ["Transporte/Automóvel", "Combustível"], ["Transporte/Automóvel", "Manutenção"], ["Transporte/Automóvel", "Estacionamento"],
  ["Moradia", "Aluguel/Financiamento"], ["Moradia", "Contas (água/luz/internet)"], ["Moradia", "Manutenção"],
  ["Sítio", "Manutenção"], ["Sítio", "Insumos"],
  ["Lazer/Festas", "Festas"], ["Lazer/Festas", "Passeios"], ["Lazer/Festas", "Assinaturas"],
  ["Educação", "Cursos"], ["Educação", "Material"],
  ["Outros", "Outros"],
  ["Receita", "Salário"], ["Receita", "Extra"], ["Receita", "Juros"],
];

function doGet(e) {
  const action = (e.parameter.action || "list").toLowerCase();
  try {
    if (action === "bootstrap") {
      bootstrap_();
      return jsonOut_({ ok: true, message: "Planilha inicializada." });
    }
    if (action === "list") {
      const sheetName = requireSheetParam_(e);
      return jsonOut_({ ok: true, data: listRows_(sheetName) });
    }
    if (action === "listall") {
      bootstrap_(); // garante que as abas existem
      return jsonOut_({
        ok: true,
        data: {
          [SHEETS.LANCAMENTOS]: listRows_(SHEETS.LANCAMENTOS),
          [SHEETS.CATEGORIAS]: listRows_(SHEETS.CATEGORIAS),
          [SHEETS.PARCELAMENTOS]: listRows_(SHEETS.PARCELAMENTOS),
          [SHEETS.INVESTIMENTOS]: listRows_(SHEETS.INVESTIMENTOS),
        },
      });
    }
    return jsonOut_({ ok: false, error: "Ação GET desconhecida: " + action });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = (body.action || "").toLowerCase();
    const sheetName = body.sheet;

    if (action === "create") {
      const row = createRow_(sheetName, body.data);
      return jsonOut_({ ok: true, data: row });
    }
    if (action === "bulkcreate") {
      const rows = (body.items || []).map((item) => createRow_(sheetName, item));
      return jsonOut_({ ok: true, data: rows });
    }
    if (action === "update") {
      const row = updateRow_(sheetName, body.id, body.data);
      return jsonOut_({ ok: true, data: row });
    }
    if (action === "delete") {
      deleteRow_(sheetName, body.id);
      // apagar um parcelamento também apaga as parcelas geradas dele
      if (sheetName === SHEETS.PARCELAMENTOS) {
        deleteWhere_(SHEETS.LANCAMENTOS, "parcelamento_id", body.id);
      }
      return jsonOut_({ ok: true });
    }
    return jsonOut_({ ok: false, error: "Ação POST desconhecida: " + action });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

// ---------- helpers ----------

function requireSheetParam_(e) {
  const sheetName = e.parameter.sheet;
  if (!sheetName || !HEADERS[sheetName]) {
    throw new Error("Parâmetro 'sheet' inválido ou ausente.");
  }
  return sheetName;
}

function bootstrap_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(HEADERS).forEach((name) => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
    }
    const headers = HEADERS[name];
    const firstRow = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
    const hasHeaders = headers.every((h, i) => firstRow[i] === h);
    if (!hasHeaders) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
    }
  });

  const catSheet = ss.getSheetByName(SHEETS.CATEGORIAS);
  if (catSheet.getLastRow() < 2) {
    const rows = DEFAULT_CATEGORIES.map(([macro, micro]) => [macro, micro, true]);
    catSheet.getRange(2, 1, rows.length, 3).setValues(rows);
  }

  // remover a aba padrão "Página1"/"Sheet1" se ainda estiver vazia e sozinha
  const defaultSheet = ss.getSheetByName("Página1") || ss.getSheetByName("Sheet1");
  if (defaultSheet && ss.getSheets().length > 1 && defaultSheet.getLastRow() === 0) {
    ss.deleteSheet(defaultSheet);
  }
}

function getSheet_(name) {
  if (!HEADERS[name]) throw new Error("Aba desconhecida: " + name);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    bootstrap_();
    sheet = ss.getSheetByName(name);
  }
  return sheet;
}

function listRows_(name) {
  const sheet = getSheet_(name);
  const headers = HEADERS[name];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  return values
    .filter((row) => row.some((cell) => cell !== "" && cell !== null))
    .map((row) => rowToObject_(headers, row));
}

function rowToObject_(headers, row) {
  const obj = {};
  headers.forEach((h, i) => {
    let value = row[i];
    if (value instanceof Date) {
      value = Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
    }
    obj[h] = value;
  });
  return obj;
}

function createRow_(name, data) {
  const sheet = getSheet_(name);
  const headers = HEADERS[name];
  const id = data.id || Utilities.getUuid();
  const record = Object.assign({}, data, { id });
  if (headers.includes("criado_em") && !record.criado_em) {
    record.criado_em = new Date().toISOString();
  }
  const row = headers.map((h) => (record[h] !== undefined ? record[h] : ""));
  sheet.appendRow(row);
  return record;
}

function findRowIndexById_(sheet, headers, id) {
  const idCol = headers.indexOf("id");
  if (idCol === -1) throw new Error("Aba sem coluna 'id'.");
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const ids = sheet.getRange(2, idCol + 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2; // linha real na planilha
  }
  return -1;
}

function updateRow_(name, id, data) {
  const sheet = getSheet_(name);
  const headers = HEADERS[name];
  const rowIndex = findRowIndexById_(sheet, headers, id);
  if (rowIndex === -1) throw new Error("Registro não encontrado: " + id);
  const current = rowToObject_(headers, sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0]);
  const updated = Object.assign({}, current, data, { id });
  const row = headers.map((h) => (updated[h] !== undefined ? updated[h] : ""));
  sheet.getRange(rowIndex, 1, 1, headers.length).setValues([row]);
  return updated;
}

function deleteRow_(name, id) {
  const sheet = getSheet_(name);
  const headers = HEADERS[name];
  const rowIndex = findRowIndexById_(sheet, headers, id);
  if (rowIndex !== -1) sheet.deleteRow(rowIndex);
}

function deleteWhere_(name, field, value) {
  const sheet = getSheet_(name);
  const headers = HEADERS[name];
  const col = headers.indexOf(field);
  if (col === -1) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const values = sheet.getRange(2, col + 1, lastRow - 1, 1).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (String(values[i][0]) === String(value)) {
      sheet.deleteRow(i + 2);
    }
  }
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
