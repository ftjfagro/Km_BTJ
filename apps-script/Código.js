/**
 * Km_BTJ — Apps Script v2
 */

const SHEET_LOG = "Log";
const SHEET_CAD = "Cadastros";
const SHEET_LANC = "Lançamentos";
const SHEET_REL = "Relatório";
const SHEET_DESP = "Despesas";
const DRIVE_ROOT = "Comprovantes Km_BTJ";
const DESP_DATA_START = 2;

const LOG_DATA_START = 8;
const LOG_DATA_END = 207;
const LANC_DATA_START = 10;
const LANC_DATA_END = 69;
const REL_KM_START = 12;
const REL_KM_END = 71;

const COLAB_PADRAO = "Felipe Torquato Junqueira Franco";
const CARRO_PADRAO = "Corolla FSZ8B48";
const COL_CARRO = 16;

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.action === "ocr") return jsonOut_(lerOdometro_(data.image));
    if (data.action === "config") return jsonOut_(getConfig_());
    if (data.action === "list") return jsonOut_(listarLancamentos_(data.colaborador));
    if (data.action === "listDespesas") return jsonOut_(listarDespesas_());
    if (data.action === "saveDespesa") return jsonOut_(salvarDespesa_(data));
    if (data.action === "updateDespesa") return jsonOut_(atualizarDespesa_(data));
    if (data.action === "deleteDespesa") return jsonOut_(excluirDespesa_(data));
    if (data.action === "ocrExtrato") return jsonOut_(lerExtratoPedagio_(data.image));
    if (data.action === "checarDuplicatas") return jsonOut_({ ok: true, passagens: checarDuplicatas_(data.passagens) });
    if (data.action === "ocrCupom") return jsonOut_(lerCupom_(data.image));
    return jsonOut_(salvarLancamento_(data));
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

function doGet(e) {
  try {
    if (e && e.parameter && e.parameter.action === "config") {
      return jsonOut_(getConfig_());
    }
    return jsonOut_({ ok: true, msg: "Km_BTJ endpoint ativo (v2)." });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  }
}

function getConfig_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cad = ss.getSheetByName(SHEET_CAD);

  const destinos = [];
  const cidadesRange = cad.getRange(3, 7, 40, 1).getValues();
  for (const [c] of cidadesRange) {
    if (c === "" || c === null) break;
    destinos.push(String(c));
  }

  const colaboradores = [];
  const colabRange = cad.getRange(4, 1, 60, 3).getValues();
  for (const [nome, setor, cpf] of colabRange) {
    if (nome === "" || nome === null) break;
    colaboradores.push({ nome: String(nome), setor: String(setor || ""), cpf: String(cpf || "") });
  }

  const taxas = lerTabelaTaxas_(cad);

  const carros = [];
  const carrosRange = cad.getRange(3, 16, 40, 1).getValues();
  for (const [c] of carrosRange) {
    if (c === "" || c === null) continue;
    if (String(c).toLowerCase() === "carros") continue;
    carros.push(String(c));
  }

  return { ok: true, destinos: destinos, colaboradores: colaboradores, taxas: taxas, carros: carros };
}

function lerTabelaTaxas_(cad) {
  const taxas = [];
  const range = cad.getRange(3, 12, 60, 3).getValues();
  for (const [colab, taxa, vigencia] of range) {
    if (colab === "" || colab === null) continue;
    if (String(colab).toLowerCase() === "colaborador") continue;
    taxas.push({
      colaborador: String(colab),
      taxa: Number(taxa),
      vigenteDesde: vigencia instanceof Date ? vigencia : new Date(vigencia),
    });
  }
  return taxas;
}

function taxaVigente_(colaborador, dataViagem) {
  const cad = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_CAD);
  const taxas = lerTabelaTaxas_(cad);
  const d = dataViagem instanceof Date ? dataViagem : new Date(dataViagem);

  function melhor(nome) {
    let melhorTaxa = null, melhorData = null;
    for (const t of taxas) {
      if (t.colaborador.toLowerCase() !== nome.toLowerCase()) continue;
      if (t.vigenteDesde > d) continue;
      if (melhorData === null || t.vigenteDesde > melhorData) {
        melhorData = t.vigenteDesde;
        melhorTaxa = t.taxa;
      }
    }
    return melhorTaxa;
  }

  return melhor(colaborador) ?? melhor("Geral") ?? 0.88;
}

function salvarLancamento_(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_LOG);
  if (!sheet) throw new Error('Aba "' + SHEET_LOG + '" não encontrada.');

  const colaborador = data.colaborador || COLAB_PADRAO;
  const carro = data.carro || CARRO_PADRAO;
  const dataViagem = parseYmd_(data.data);
  const chaveData = ymd_(dataViagem);

  let row = acharLinha_(sheet, chaveData, carro);
  if (!row) {
    row = findNextEmptyRow_(sheet);
    if (!row) throw new Error("Sem linhas livres no Log — peça para estender a planilha.");
  }

  const taxa = taxaVigente_(colaborador, dataViagem);

  sheet.getRange(row, 1, 1, 6).setValues([[
    dataViagem,
    data.tipo || "Viagem",
    data.origem || "",
    data.destino || "",
    data.kmInicial === undefined || data.kmInicial === null ? "" : data.kmInicial,
    data.kmFinal === undefined || data.kmFinal === null ? "" : data.kmFinal,
  ]]);
  sheet.getRange(row, 9, 1, 4).setValues([[
    data.categoria || "",
    data.descricao || "",
    data.valor === undefined ? "" : data.valor,
    data.observacao || "",
  ]]);
  sheet.getRange(row, 15).setValue(taxa);
  sheet.getRange(row, COL_CARRO).setValue(carro);

  return { ok: true, row: row, taxa: taxa };
}

function acharLinha_(sheet, chaveData, carro) {
  const n = LOG_DATA_END - LOG_DATA_START + 1;
  const vals = sheet.getRange(LOG_DATA_START, 1, n, COL_CARRO).getValues();
  for (let i = 0; i < n; i++) {
    const d = vals[i][0];
    const c = vals[i][COL_CARRO - 1] || CARRO_PADRAO;
    if (d instanceof Date && ymd_(d) === chaveData && String(c) === String(carro)) {
      return LOG_DATA_START + i;
    }
  }
  return null;
}

function listarLancamentos_(colaborador) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_LOG);
  if (!sheet) throw new Error('Aba "' + SHEET_LOG + '" não encontrada.');

  const n = LOG_DATA_END - LOG_DATA_START + 1;
  const vals = sheet.getRange(LOG_DATA_START, 1, n, COL_CARRO).getValues();
  const out = [];
  for (let i = 0; i < n; i++) {
    const r = vals[i];
    const data = r[0];
    if (!(data instanceof Date)) continue;
    out.push({
      data: ymd_(data),
      tipo: r[1] || "Viagem",
      origem: r[2] || "",
      destino: r[3] || "",
      kmInicial: r[4] === "" || r[4] === null ? null : Number(r[4]),
      kmFinal: r[5] === "" || r[5] === null ? null : Number(r[5]),
      observacao: r[11] || "",
      taxa: r[14] === "" || r[14] === null ? null : Number(r[14]),
      carro: r[COL_CARRO - 1] || CARRO_PADRAO,
    });
  }
  return { ok: true, lancamentos: out };
}

function lerOdometro_(base64Image) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
  if (!apiKey) throw new Error("OPENAI_API_KEY não configurada nas Propriedades do script.");
  if (!base64Image) throw new Error("Nenhuma imagem recebida.");

  const payload = {
    model: "gpt-4o-mini",
    max_tokens: 50,
    messages: [{
      role: "user",
      content: [
        {
          type: "text",
          text: [
            "Esta é a foto do painel de um carro (hodômetro digital).",
            "Preciso APENAS do ODÔMETRO TOTAL — a quilometragem acumulada do veículo.",
            "Ele aparece ao lado da sigla \"ODO\" (geralmente no canto inferior do painel) e costuma ter 5 ou 6 dígitos.",
            "IGNORE completamente estes outros números, que NÃO são o odômetro:",
            "- autonomia / km restante (número com 'km' perto do ícone de bomba de combustível)",
            "- média de consumo (ex: '11.2 km/L')",
            "- temperatura externa (ex: '13°C')",
            "- horário (ex: '20:59')",
            "- a escala do velocímetro (0, 20, 40 ... 240) e a velocidade atual",
            "Responda SOMENTE com o número inteiro do ODO, sem pontos, sem espaços, sem unidade, sem texto.",
            "Exemplo de resposta correta: 238849",
            "Se não conseguir identificar com clareza o número do ODO, responda apenas: 0"
          ].join("\n")
        },
        { type: "image_url", image_url: { url: "data:image/jpeg;base64," + base64Image } }
      ]
    }]
  };

  const resp = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", {
    method: "post",
    contentType: "application/json",
    headers: { Authorization: "Bearer " + apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  const body = JSON.parse(resp.getContentText());
  if (body.error) throw new Error(body.error.message || "Erro na API da OpenAI");

  const text = (body.choices && body.choices[0] && body.choices[0].message.content || "").trim();
  const match = text.match(/\d[\d.]*\d|\d+/);
  if (!match) return { ok: true, km: null, raw: text };

  const km = parseInt(match[0].replace(/\./g, ""), 10);
  if (!km) return { ok: true, km: null, raw: text };
  return { ok: true, km: km };
}

function migrar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cad = ss.getSheetByName(SHEET_CAD);
  const log = ss.getSheetByName(SHEET_LOG);
  const lanc = ss.getSheetByName(SHEET_LANC);
  const rel = ss.getSheetByName(SHEET_REL);

  cad.getRange(2, 12).setValue("Taxas de reembolso").setFontWeight("bold");
  cad.getRange(3, 12, 1, 3).setValues([["Colaborador", "Taxa (R$/km)", "Vigente desde"]]).setFontWeight("bold");
  cad.getRange(4, 12, 2, 3).setValues([
    [COLAB_PADRAO, 1.12, new Date(2026, 0, 1)],
    ["Geral", 0.88, new Date(2026, 0, 1)],
  ]);
  cad.getRange(4, 14, 2, 1).setNumberFormat("dd/mm/yyyy");
  cad.getRange(4, 13, 2, 1).setNumberFormat("R$ #,##0.00");
  Logger.log("1/4 Tabela de Taxas criada no Cadastros (L2:N5).");

  const nRows = LOG_DATA_END - LOG_DATA_START + 1;
  const datas = log.getRange(LOG_DATA_START, 1, nRows, 1).getValues();

  const formulasG = [], formulasH = [], taxasO = [];
  let preenchidas = 0;
  for (let i = 0; i < nRows; i++) {
    const r = LOG_DATA_START + i;
    formulasG.push(['=IF(AND(E' + r + '<>"",F' + r + '<>""),F' + r + '-E' + r + ',"")']);
    formulasH.push(['=IF(AND(G' + r + '<>"",O' + r + '<>""),G' + r + '*O' + r + ',"")']);
    const d = datas[i][0];
    if (d !== "" && d !== null) {
      taxasO.push([taxaVigente_(COLAB_PADRAO, d)]);
      preenchidas++;
    } else {
      taxasO.push([""]);
    }
  }
  log.getRange(LOG_DATA_START, 7, nRows, 1).setFormulas(formulasG);
  log.getRange(LOG_DATA_START, 8, nRows, 1).setFormulas(formulasH);
  log.getRange(LOG_DATA_START, 15, nRows, 1).setValues(taxasO);
  log.getRange(LOG_DATA_START, 15, nRows, 1).setNumberFormat("0.00");
  Logger.log("2/4 Log: fórmulas G/H reescritas; taxa aplicada preenchida em " + preenchidas + " lançamentos históricos.");

  lanc.getRange(9, 12).setValue("Taxa (R$/km)");
  const nLanc = LANC_DATA_END - LANC_DATA_START + 1;
  const formulasL = [], formulasK = [];
  for (let i = 0; i < nLanc; i++) {
    const r = LANC_DATA_START + i;
    const k = i + 1;
    formulasL.push(['=IFERROR(INDEX(Log!$O:$O,MATCH(' + k + ',Log!$N:$N,0)),"")']);
    formulasK.push(['=IF(AND(J' + r + '<>"",L' + r + '<>""),J' + r + '*L' + r + ',"")']);
  }
  lanc.getRange(LANC_DATA_START, 12, nLanc, 1).setFormulas(formulasL);
  lanc.getRange(LANC_DATA_START, 11, nLanc, 1).setFormulas(formulasK);
  Logger.log("3/4 Lançamentos: taxa por linha (L) e reembolso (K) reescritos.");

  const nRel = REL_KM_END - REL_KM_START + 1;
  const formulasRelG = [], formulasRelH = [];
  for (let i = 0; i < nRel; i++) {
    const r = REL_KM_START + i;
    const k = i + 1;
    formulasRelG.push(['=IFERROR(INDEX(Lançamentos!$L:$L,MATCH(' + k + ',Lançamentos!$N:$N,0)),"")']);
    formulasRelH.push(['=IF(AND(F' + r + '<>"",G' + r + '<>""),F' + r + '*G' + r + ',"")']);
  }
  rel.getRange(REL_KM_START, 7, nRel, 1).setFormulas(formulasRelG);
  rel.getRange(REL_KM_START, 8, nRel, 1).setFormulas(formulasRelH);
  Logger.log("4/4 Relatório: coluna R$/km (G) e total (H) reescritos.");

  Logger.log("MIGRACAO CONCLUIDA. Confira: Cadastros L2:N5, Log coluna O, e os totais do Relatório (abril deve somar R$ 2.884,48).");
}

function migrarCarros() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cad = ss.getSheetByName(SHEET_CAD);
  const log = ss.getSheetByName(SHEET_LOG);

  cad.getRange(2, 16).setValue("Carros").setFontWeight("bold");
  cad.getRange(3, 16, 2, 1).setValues([["Corolla FSZ8B48"], ["Outlander FXJ5336"]]);
  Logger.log("1/2 Tabela de Carros criada no Cadastros (P2:P4).");

  log.getRange(7, COL_CARRO).setValue("Carro").setFontWeight("bold");
  const n = LOG_DATA_END - LOG_DATA_START + 1;
  const datas = log.getRange(LOG_DATA_START, 1, n, 1).getValues();
  const carroCol = [];
  let marcadas = 0;
  for (let i = 0; i < n; i++) {
    if (datas[i][0] instanceof Date) { carroCol.push([CARRO_PADRAO]); marcadas++; }
    else carroCol.push([""]);
  }
  log.getRange(LOG_DATA_START, COL_CARRO, n, 1).setValues(carroCol);
  Logger.log("2/2 Coluna Carro no Log: " + marcadas + " lançamentos históricos marcados como " + CARRO_PADRAO + ".");
  Logger.log("MIGRACAO DE CARROS CONCLUIDA.");
}

const DESP_HEADERS = ["ID", "Data", "Carro", "Tipo", "Valor (R$)", "Descrição", "Comprovante", "Origem", "Registrado em"];

function getDespSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_DESP);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_DESP);
    sheet.getRange(1, 1, 1, DESP_HEADERS.length).setValues([DESP_HEADERS]).setFontWeight("bold");
    sheet.setFrozenRows(1);
    sheet.getRange("B:B").setNumberFormat("dd/mm/yyyy");
    sheet.getRange("E:E").setNumberFormat("R$ #,##0.00");
  }
  return sheet;
}

function novoId_() {
  return "d" + new Date().getTime() + Math.floor(Math.random() * 1000);
}

function salvarDespesa_(data) {
  const sheet = getDespSheet_();
  const dataDesp = parseYmd_(data.data);
  const chaveData = ymd_(dataDesp);
  const carro = data.carro || CARRO_PADRAO;
  const tipo = data.tipo || "Outros";
  const valor = Number(data.valor) || 0;

  let link = data.comprovanteLink || "";
  if (!link && data.comprovanteImage) {
    link = uploadComprovante_(data.comprovanteImage, chaveData, tipo);
  }

  const id = novoId_();
  sheet.appendRow([id, dataDesp, carro, tipo, valor, data.descricao || "", link, data.origem || "manual", new Date()]);
  return { ok: true, id: id, somado: false };
}

function existeDespesaIgual_(chaveData, descricao, valor) {
  const sheet = getDespSheet_();
  const last = sheet.getLastRow();
  if (last < DESP_DATA_START) return false;
  const vals = sheet.getRange(DESP_DATA_START, 1, last - DESP_DATA_START + 1, 6).getValues();
  const descNorm = String(descricao || "").trim().toLowerCase();
  for (const r of vals) {
    const d = r[1];
    if (!(d instanceof Date)) continue;
    if (ymd_(d) === chaveData && String(r[5] || "").trim().toLowerCase() === descNorm && Math.abs((Number(r[4]) || 0) - valor) < 0.005) {
      return true;
    }
  }
  return false;
}

function checarDuplicatas_(passagens) {
  return (passagens || []).map(p => ({
    ...p,
    jaLancado: existeDespesaIgual_(p.data, p.local, Number(p.valor) || 0),
  }));
}

function acharDespesaPorId_(sheet, id) {
  const last = sheet.getLastRow();
  if (last < DESP_DATA_START) return null;
  const ids = sheet.getRange(DESP_DATA_START, 1, last - DESP_DATA_START + 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return DESP_DATA_START + i;
  }
  return null;
}

function atualizarDespesa_(data) {
  const sheet = getDespSheet_();
  const row = acharDespesaPorId_(sheet, data.id);
  if (!row) throw new Error("Despesa não encontrada (id " + data.id + ").");
  if (data.data !== undefined) sheet.getRange(row, 2).setValue(parseYmd_(data.data));
  if (data.carro !== undefined) sheet.getRange(row, 3).setValue(data.carro);
  if (data.tipo !== undefined) sheet.getRange(row, 4).setValue(data.tipo);
  if (data.valor !== undefined) sheet.getRange(row, 5).setValue(Number(data.valor) || 0);
  if (data.descricao !== undefined) sheet.getRange(row, 6).setValue(data.descricao);
  return { ok: true, id: data.id };
}

function excluirDespesa_(data) {
  const sheet = getDespSheet_();
  const row = acharDespesaPorId_(sheet, data.id);
  if (!row) throw new Error("Despesa não encontrada (id " + data.id + ").");
  sheet.deleteRow(row);
  return { ok: true, id: data.id };
}

function listarDespesas_() {
  const sheet = getDespSheet_();
  const last = sheet.getLastRow();
  const out = [];
  if (last >= DESP_DATA_START) {
    const vals = sheet.getRange(DESP_DATA_START, 1, last - DESP_DATA_START + 1, DESP_HEADERS.length).getValues();
    for (const r of vals) {
      if (!(r[1] instanceof Date)) continue;
      out.push({
        id: r[0] || "", data: ymd_(r[1]), carro: r[2] || "", tipo: r[3] || "", valor: Number(r[4]) || 0,
        descricao: r[5] || "", comprovante: r[6] || "", origem: r[7] || "",
      });
    }
  }
  return { ok: true, despesas: out };
}

function migrarDespesasId() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_DESP);
  if (!sheet) { Logger.log("Aba Despesas ainda não existe — nada a migrar."); return; }
  const primeiroCabecalho = sheet.getRange(1, 1).getValue();
  if (String(primeiroCabecalho).toUpperCase() === "ID") { Logger.log("Coluna ID já existe."); return; }
  sheet.insertColumnBefore(1);
  sheet.getRange(1, 1).setValue("ID").setFontWeight("bold");
  const last = sheet.getLastRow();
  for (let r = DESP_DATA_START; r <= last; r++) {
    if (sheet.getRange(r, 2).getValue() instanceof Date) {
      sheet.getRange(r, 1).setValue(novoId_());
    }
  }
  Logger.log("Coluna ID adicionada à aba Despesas e IDs gerados.");
}

function periodoLabel_(chaveData) {
  const [y, m, d] = chaveData.split("-").map(Number);
  let py = y, pm = m;
  if (d >= 26) { pm = m + 1; if (pm > 12) { pm = 1; py = y + 1; } }
  return py + "-" + String(pm).padStart(2, "0");
}

function getPastaPeriodo_(chaveData) {
  const root = getOrCreateFolder_(DriveApp.getRootFolder(), DRIVE_ROOT);
  return getOrCreateFolder_(root, periodoLabel_(chaveData));
}

function getOrCreateFolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function uploadComprovante_(base64, chaveData, tipo) {
  const bytes = Utilities.base64Decode(base64);
  const nome = tipo + "_" + chaveData + "_" + new Date().getTime() + ".jpg";
  const blob = Utilities.newBlob(bytes, "image/jpeg", nome);
  const pasta = getPastaPeriodo_(chaveData);
  const file = pasta.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

function lerExtratoPedagio_(base64Image) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
  if (!apiKey) throw new Error("OPENAI_API_KEY não configurada nas Propriedades do script.");
  if (!base64Image) throw new Error("Nenhuma imagem recebida.");

  const payload = {
    model: "gpt-4o-mini",
    max_tokens: 1500,
    messages: [{
      role: "user",
      content: [
        {
          type: "text",
          text: [
            "Esta é a foto/print de um extrato de pedágio (Sem Parar, ConectCar, Veloe, C6, etc.).",
            "Extraia TODAS as passagens de pedágio visíveis. Para cada uma capture: data, local/praça/rodovia, e valor em reais.",
            "Responda APENAS com um JSON válido, sem texto antes ou depois, no formato:",
            '{"passagens":[{"data":"AAAA-MM-DD","local":"nome da praça ou rodovia","valor":9.20}]}',
            "Regras:",
            "- data no formato AAAA-MM-DD. Se o ano não aparecer, use o ano atual.",
            "- valor como número decimal com ponto (ex: 9.20), sem 'R$'.",
            "- ignore linhas de saldo, recarga, mensalidade — só passagens de pedágio.",
            '- se não houver nenhuma passagem, responda {"passagens":[]}.'
          ].join("\n")
        },
        { type: "image_url", image_url: { url: "data:image/jpeg;base64," + base64Image } }
      ]
    }]
  };

  const resp = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", {
    method: "post",
    contentType: "application/json",
    headers: { Authorization: "Bearer " + apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  const body = JSON.parse(resp.getContentText());
  if (body.error) throw new Error(body.error.message || "Erro na API da OpenAI");

  let text = (body.choices && body.choices[0] && body.choices[0].message.content || "").trim();
  text = text.replace(/^```json/i, "").replace(/^```/, "").replace(/```$/, "").trim();
  let parsed;
  try { parsed = JSON.parse(text); } catch (e) { return { ok: true, passagens: [], raw: text }; }
  return { ok: true, passagens: parsed.passagens || [] };
}

function lerCupom_(base64Image) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
  if (!apiKey) throw new Error("OPENAI_API_KEY não configurada nas Propriedades do script.");
  if (!base64Image) throw new Error("Nenhuma imagem recebida.");

  const payload = {
    model: "gpt-4o-mini",
    max_tokens: 300,
    messages: [{
      role: "user",
      content: [
        {
          type: "text",
          text: [
            "Esta é a foto de um cupom fiscal / nota / recibo (ex: restaurante, estacionamento, posto).",
            "Extraia: o VALOR TOTAL pago, a DATA, e uma DESCRIÇÃO curta (nome do estabelecimento).",
            "Responda APENAS com JSON válido, sem texto antes ou depois, no formato:",
            '{"valor":48.90,"data":"AAAA-MM-DD","descricao":"Restaurante Sabor Caseiro"}',
            "Regras: valor como número decimal com ponto (o TOTAL, não subtotais); data AAAA-MM-DD (se faltar ano, use o atual);",
            "descrição curta. Se algum campo não for legível, use null nele."
          ].join("\n")
        },
        { type: "image_url", image_url: { url: "data:image/jpeg;base64," + base64Image } }
      ]
    }]
  };

  const resp = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", {
    method: "post",
    contentType: "application/json",
    headers: { Authorization: "Bearer " + apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  const body = JSON.parse(resp.getContentText());
  if (body.error) throw new Error(body.error.message || "Erro na API da OpenAI");
  let text = (body.choices && body.choices[0] && body.choices[0].message.content || "").trim();
  text = text.replace(/^```json/i, "").replace(/^```/, "").replace(/```$/, "").trim();
  try { return { ok: true, cupom: JSON.parse(text) }; }
  catch (e) { return { ok: true, cupom: null, raw: text }; }
}

function autorizarDrive() {
  const f = getOrCreateFolder_(DriveApp.getRootFolder(), DRIVE_ROOT);
  Logger.log("Pasta pronta: " + f.getName() + " — autorização de Drive concedida.");
}

function consertarDatasDeslocadas() {
  const DIAS_ERRADOS = ["2026-07-12", "2026-07-13", "2026-07-14", "2026-07-15"];
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOG);
  const n = LOG_DATA_END - LOG_DATA_START + 1;
  const datas = sheet.getRange(LOG_DATA_START, 1, n, 1).getValues();
  let corrigidas = 0;
  for (let i = 0; i < n; i++) {
    const d = datas[i][0];
    if (!(d instanceof Date)) continue;
    if (DIAS_ERRADOS.indexOf(ymd_(d)) !== -1) {
      const nova = new Date(d.getFullYear(), d.getMonth(), d.getDate() + 1, 12, 0, 0);
      sheet.getRange(LOG_DATA_START + i, 1).setValue(nova);
      corrigidas++;
    }
  }
  Logger.log("Datas corrigidas (+1 dia): " + corrigidas + " linhas. Confira a aba Log.");
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function ymd_(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function parseYmd_(s) {
  if (s instanceof Date) return s;
  const partes = String(s).slice(0, 10).split("-");
  return new Date(Number(partes[0]), Number(partes[1]) - 1, Number(partes[2]), 12, 0, 0);
}

function findNextEmptyRow_(sheet) {
  const values = sheet.getRange(LOG_DATA_START, 1, LOG_DATA_END - LOG_DATA_START + 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === "" || values[i][0] === null) return LOG_DATA_START + i;
  }
  return null;
}

function autorizar() {
  UrlFetchApp.fetch("https://api.openai.com/v1/models", {
    headers: { Authorization: "Bearer " + PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY") },
    muteHttpExceptions: true,
  });
  Logger.log("Autorizacao concedida.");
}