import { useState, useEffect, useRef, useMemo } from "react";
import { registerSW } from "virtual:pwa-register";
import * as XLSX from "xlsx";

// ─── Config fixa (infraestrutura) ─────────────────────────────────────────────
const APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbwDv9SdQANzfKb5oWQAU_7evjHN5BKLbW3kxb2eZbhVM1Weku0xUmDyHup7KRVTF8bPdw/exec";
const SOLICITANTE = "Felipe Torquato Junqueira Franco";
const SETOR = "Diretor";
const CPF = "372.742.538-59";

// Identidade BTJ
const BTJ_NAVY = "#14213f";   // navy do novo ícone (velocímetro BTJ)
const BTJ_BLUE = "#1EABE3";   // azul de ação (botões) — mantido pra contraste com texto branco
const BTJ_LIGHT = "#5fd0ff";  // ciano acento do novo ícone (textos claros sobre navy)

// Fallback caso a config da planilha não carregue (offline no primeiro uso)
const DEFAULT_CONFIG = {
  destinos: ["Sud", "Ilha", "RP", "SP", "Campinas", "Jundiaí", "Ribeirão", "VCP", "Foods", "Sud Foods", "Foods RP", "Foods Prudente", "Prudente"],
  carros: ["Corolla FSZ8B48", "Outlander FXJ5336"],
  taxas: [
    { categoriaKm: "Viajante", taxa: 1.12, vigenteDesde: "2025-12-31" },
    { categoriaKm: "Geral", taxa: 0.88, vigenteDesde: "2025-12-31" },
  ],
  colaboradores: [{ nome: SOLICITANTE, categoriaKm: "Viajante", faixa: "Aberto", setor: "Diretor", cpf: "" }],
  limites: [],
};
const CARRO_PADRAO = "Corolla FSZ8B48";

// GPS: município oficial → apelido usado na planilha
const CITY_GPS_MAP = {
  "Sud Mennucci": "Sud",
  "Ilha Solteira": "Ilha",
  "Ribeirão Preto": "RP",
  "São Paulo": "SP",
  "Campinas": "Campinas",
  "Jundiaí": "Jundiaí",
  "Presidente Prudente": "Prudente",
};

const WEEKDAYS_PT = ["domingo", "segunda-feira", "terça-feira", "quarta-feira", "quinta-feira", "sexta-feira", "sábado"];
const WEEKDAYS_ABREV_PT = ["dom", "seg", "ter", "qua", "qui", "sex", "sáb"];
const MONTHS_PT = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"];

// ─── Helpers de data ──────────────────────────────────────────────────────────
function todayISO() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}
function parseISO(s) {
  const [y, m, d] = s.split("-").map(Number);
  return new Date(y, m - 1, d);
}
// "agora mesmo", "há 3 min", "há 2 h", ou a hora cheia se passou muito.
function tempoRelativo(date) {
  if (!date) return "";
  const seg = Math.floor((Date.now() - date.getTime()) / 1000);
  if (seg < 45) return "agora mesmo";
  if (seg < 90) return "há 1 min";
  const min = Math.floor(seg / 60);
  if (min < 60) return `há ${min} min`;
  const h = Math.floor(min / 60);
  if (h < 6) return `há ${h} h`;
  return `às ${String(date.getHours()).padStart(2, "0")}:${String(date.getMinutes()).padStart(2, "0")}`;
}
function formatDateBR(s) {  const [y, m, d] = s.split("-");
  return `${d}/${m}/${y}`;
}
function formatDateShort(s) {
  const [, m, d] = s.split("-");
  return `${d}/${m}`;
}
function weekdayPT(iso) {
  return WEEKDAYS_PT[parseISO(iso).getDay()];
}
function weekdayAbrev(iso) {
  return WEEKDAYS_ABREV_PT[parseISO(iso).getDay()];
}
// Semana de domingo a sábado: a chave é o ISO do domingo daquela semana.
function weekKey(iso) {
  const d = parseISO(iso);
  d.setDate(d.getDate() - d.getDay()); // dom=0 ... sáb=6
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}
function weekLabel(key) {
  const ini = parseISO(key);
  const fim = new Date(ini);
  fim.setDate(fim.getDate() + 6);
  const f = dt => `${String(dt.getDate()).padStart(2, "0")}/${String(dt.getMonth() + 1).padStart(2, "0")}`;
  return `${f(ini)} – ${f(fim)}`;
}
function monthKey(iso) {
  return iso.slice(0, 7); // "2026-07"
}

// Ciclo de reembolso 26→25. Uma data pertence ao período que "fecha" no dia 25.
// Ex: 26/06 a 25/07 → período "2026-07" (Julho). 25/06 → "2026-06" (Junho).
function periodKey(iso) {
  const [y, m, d] = iso.split("-").map(Number);
  let py = y, pm = m;
  if (d >= 26) { pm = m + 1; if (pm > 12) { pm = 1; py = y + 1; } }
  return `${py}-${String(pm).padStart(2, "0")}`;
}
function periodLabel(key) {
  const [y, m] = key.split("-").map(Number);
  return `${MONTHS_PT[m - 1]} ${y}`;
}
function periodRange(key) {
  const [y, m] = key.split("-").map(Number);
  // fim = dia 25 do mês do período; início = dia 26 do mês anterior
  const fim = `${y}-${String(m).padStart(2, "0")}-25`;
  let iy = y, im = m - 1; if (im < 1) { im = 12; iy = y - 1; }
  const ini = `${iy}-${String(im).padStart(2, "0")}-26`;
  return { ini, fim };
}
function monthLabelFromKey(key) {
  const [y, m] = key.split("-").map(Number);
  return `${MONTHS_PT[m - 1]} ${y}`;
}
function prevPeriodKey(key) {
  let [y, m] = key.split("-").map(Number);
  m -= 1; if (m < 1) { m = 12; y -= 1; }
  return `${y}-${String(m).padStart(2, "0")}`;
}

// ─── Taxa vigente (mesma regra do Apps Script, no cliente) ────────────────────
function taxaVigente(taxas, colaboradores, colaborador, isoDate) {
  const d = parseISO(isoDate);
  const c = (colaboradores || []).find(c => String(c.nome).toLowerCase() === String(colaborador).toLowerCase());
  const categoriaKm = c ? c.categoriaKm : "";
  function melhor(cat) {
    if (!cat) return null;
    let mTaxa = null, mData = null;
    for (const t of taxas) {
      if (String(t.categoriaKm).toLowerCase() !== cat.toLowerCase()) continue;
      const v = typeof t.vigenteDesde === "string" ? new Date(t.vigenteDesde) : new Date(t.vigenteDesde);
      if (v > d) continue;
      if (mData === null || v > mData) { mData = v; mTaxa = Number(t.taxa); }
    }
    return mTaxa;
  }
  return melhor(categoriaKm) ?? melhor("Geral") ?? 0.88;
}

// Categorias de despesa que podem ser rateadas entre várias pessoas (e que
// têm limite de reembolso por pessoa, definido por faixa do colaborador).
const CATEGORIAS_RATEAVEIS = ["Alimentação", "Hotel", "Outros"];

// Acha a faixa (Aberto/A/B) do colaborador na config já carregada.
function faixaDoColaborador(colaboradores, nome) {
  const c = (colaboradores || []).find(c => String(c.nome).toLowerCase() === String(nome).toLowerCase());
  return c ? String(c.faixa || "") : "";
}

// Limite por pessoa vigente pra uma categoria + faixa numa data.
// Retorna null se não houver faixa ou não houver limite cadastrado (sem trava).
function limiteVigente(limites, faixa, categoria, isoDate) {
  if (!faixa) return null;
  const d = parseISO(isoDate);
  let mLimite = null, mData = null;
  for (const l of (limites || [])) {
    if (String(l.categoria).toLowerCase() !== String(categoria).toLowerCase()) continue;
    if (String(l.faixa).toLowerCase() !== faixa.toLowerCase()) continue;
    const v = typeof l.vigenteDesde === "string" ? new Date(l.vigenteDesde) : new Date(l.vigenteDesde);
    if (v > d) continue;
    if (mData === null || v > mData) { mData = v; mLimite = Number(l.limite); }
  }
  return mLimite;
}

// ─── Storage ──────────────────────────────────────────────────────────────────
const KEY_RECORDS = "km_registros_v3";
const KEY_RECORDS_OLD = "km_registros_v2";
const KEY_CONFIG = "km_config_cache";
const KEY_LAST_CAR = "km_ultimo_carro";

function loadLastCar() {
  return localStorage.getItem(KEY_LAST_CAR) || CARRO_PADRAO;
}
function saveLastCar(c) {
  localStorage.setItem(KEY_LAST_CAR, c);
}

// ─── Sessão de login (fica salva pra sempre, até "Sair" explícito) ───────────
const KEY_SESSAO = "km_sessao_v1";
function loadSessao() {
  try {
    const raw = localStorage.getItem(KEY_SESSAO);
    return raw ? JSON.parse(raw) : null;
  } catch { return null; }
}
function saveSessao(colaborador) {
  localStorage.setItem(KEY_SESSAO, JSON.stringify(colaborador));
}
function limparSessao() {
  localStorage.removeItem(KEY_SESSAO);
}

// ─── Envio de relatórios (status por período: enviado/pendente + retorno) ────
const KEY_ENVIOS = "km_envios_v1";
function loadEnvios() {
  try { return JSON.parse(localStorage.getItem(KEY_ENVIOS)) || {}; } catch { return {}; }
}
function saveEnvios(envios) {
  localStorage.setItem(KEY_ENVIOS, JSON.stringify(envios));
}
// Período imediatamente anterior a uma chave "AAAA-MM"
function periodoAnterior(key) {
  const [y, m] = key.split("-").map(Number);
  const d = new Date(y, m - 2, 1);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
}
// Status de um período: aberto | fechado | pendente | enviado
function statusPeriodo(periodo, envios, hojeKey) {
  const e = envios[periodo];
  if (e && e.status === "enviado") return "enviado";
  if (e && e.status === "pendente") return "pendente";
  if (e && e.status === "reaberto") return "reaberto"; // reabertura aprovada: editável até reenviar
  return periodo === hojeKey ? "aberto" : "fechado";
}

const KEY_PHOTO_QUEUE = "km_fotos_pendentes";

// Registro válido = data ISO plausível. Lançamentos com data quebrada (ex: o
// famoso 31/12/1969, a "data zero" do computador, criado por um bug antigo)
// são descartados na carga — sem isso eles ficam ressuscitando: moram no
// localStorage e o app os re-sincroniza pra planilha mesmo depois de apagados lá.
function registroValido_(r) {
  return r && typeof r.data === "string" && /^\d{4}-\d{2}-\d{2}$/.test(r.data) && r.data >= "2020-01-01";
}

function loadRecords() {
  try {
    const raw = localStorage.getItem(KEY_RECORDS);
    if (raw) {
      const recs = JSON.parse(raw);
      const limpos = recs.filter(registroValido_);
      if (limpos.length !== recs.length) localStorage.setItem(KEY_RECORDS, JSON.stringify(limpos));
      return limpos;
    }
    const old = localStorage.getItem(KEY_RECORDS_OLD);
    if (old) {
      const migrated = JSON.parse(old).map(r => ({ ...r, observacao: "", synced: true })).filter(registroValido_);
      localStorage.setItem(KEY_RECORDS, JSON.stringify(migrated));
      return migrated;
    }
    return [];
  } catch { return []; }
}
function persistRecords(recs) {
  localStorage.setItem(KEY_RECORDS, JSON.stringify(recs));
}
function loadCachedConfig() {
  try {
    const raw = localStorage.getItem(KEY_CONFIG);
    return raw ? JSON.parse(raw) : DEFAULT_CONFIG;
  } catch { return DEFAULT_CONFIG; }
}
function loadPhotoQueue() {
  try {
    const raw = localStorage.getItem(KEY_PHOTO_QUEUE);
    return raw ? JSON.parse(raw) : [];
  } catch { return []; }
}
function persistPhotoQueue(q) {
  try { localStorage.setItem(KEY_PHOTO_QUEUE, JSON.stringify(q)); }
  catch { avisar("Memória do navegador cheia — a foto não pôde ser guardada offline. Digite o KM manualmente."); }
}

// ─── API (Apps Script) ────────────────────────────────────────────────────────
async function apiLogin(email, senha) {
  const res = await fetch(APPS_SCRIPT_URL, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({ action: "login", email, senha }),
  });
  const data = await res.json();
  if (!data.ok) throw new Error(data.error || "Não foi possível entrar.");
  return data.colaborador;
}

async function apiFetchConfig(colaborador) {
  const res = await fetch(APPS_SCRIPT_URL, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({ action: "config", colaborador }),
  });
  const data = await res.json();
  if (!data.ok) throw new Error(data.error || "Erro ao carregar config");
  return data;
}

async function apiOcr(base64) {
  const res = await fetch(APPS_SCRIPT_URL, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({ action: "ocr", image: base64 }),
  });
  const data = await res.json();
  if (!data.ok) throw new Error(data.error || "Erro ao ler o odômetro");
  return data.km ?? null;
}

async function apiSaveDespesa(d) {
  const res = await fetch(APPS_SCRIPT_URL, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({ action: "saveDespesa", ...d }),
  });
  const data = await res.json();
  if (!data.ok) throw new Error(data.error || "Erro ao salvar despesa");
  return data;
}

async function apiAddCarro(carro, colaborador) {
  const res = await fetch(APPS_SCRIPT_URL, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({ action: "addCarro", carro, colaborador }),
  });
  const data = await res.json();
  if (!data.ok) throw new Error(data.error || "Erro ao cadastrar carro");
  return data;
}

// Emite o relatório do período: o backend gera PDF+Excel, arquiva no Drive e
// envia o e-mail. Demora ~15-20s. Retorna { ok, periodo, reemissao,
// enviadoPara, comAssinatura, pdfUrl, excelUrl, enviadoEm }.
async function apiAprovarEEmitir(periodo, assinaturaBase64) {
  const res = await fetch(APPS_SCRIPT_URL, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({ action: "aprovarEEmitir", periodo, assinaturaBase64: assinaturaBase64 || undefined }),
  });
  const data = await res.json();
  if (!data.ok) throw new Error(data.error || "Erro ao emitir o relatório");
  return data;
}

// Exclui um lançamento de km na planilha (por data+carro+colaborador).
async function apiDeleteLancamento(data, carro, colaborador) {
  const res = await fetch(APPS_SCRIPT_URL, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({ action: "deleteLancamento", data, carro, colaborador }),
  });
  const out = await res.json();
  if (!out.ok) throw new Error(out.error || "Erro ao excluir o lançamento");
  return out;
}

// ─── Reabertura de relatório já enviado ──────────────────────────────────────
async function apiSolicitarReabertura(colaborador, periodo, motivo) {
  const res = await fetch(APPS_SCRIPT_URL, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({ action: "solicitarReabertura", colaborador, periodo, motivo }),
  });
  const data = await res.json();
  if (!data.ok) throw new Error(data.error || "Erro ao solicitar a reabertura");
  return data;
}

async function apiStatusReaberturas(colaborador) {
  const res = await fetch(APPS_SCRIPT_URL, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({ action: "statusReaberturas", colaborador }),
  });
  const data = await res.json();
  if (!data.ok) throw new Error(data.error || "Erro ao consultar reaberturas");
  return data.reaberturas || [];
}

async function apiListarReaberturas(aprovador) {
  const res = await fetch(APPS_SCRIPT_URL, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({ action: "listarReaberturas", aprovador }),
  });
  const data = await res.json();
  if (!data.ok) throw new Error(data.error || "Erro ao listar solicitações");
  return data.reaberturas || [];
}

async function apiDecidirReabertura(id, decisao, aprovador) {
  const res = await fetch(APPS_SCRIPT_URL, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({ action: "decidirReabertura", id, decisao, aprovador }),
  });
  const data = await res.json();
  if (!data.ok) throw new Error(data.error || "Erro ao registrar a decisão");
  return data;
}

async function apiUpdateDespesa(d) {
  const res = await fetch(APPS_SCRIPT_URL, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({ action: "updateDespesa", ...d }),
  });
  const data = await res.json();
  if (!data.ok) throw new Error(data.error || "Erro ao atualizar despesa");
  return data;
}

async function apiDeleteDespesa(id) {
  const res = await fetch(APPS_SCRIPT_URL, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({ action: "deleteDespesa", id }),
  });
  const data = await res.json();
  if (!data.ok) throw new Error(data.error || "Erro ao excluir despesa");
  return data;
}

async function apiListDespesas(colaborador) {
  const res = await fetch(APPS_SCRIPT_URL, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({ action: "listDespesas", colaborador }),
  });
  const data = await res.json();
  if (!data.ok) throw new Error(data.error || "Erro ao listar despesas");
  return data.despesas || [];
}

async function apiOcrCupom(base64) {
  const res = await fetch(APPS_SCRIPT_URL, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({ action: "ocrCupom", image: base64 }),
  });
  const data = await res.json();
  if (!data.ok) throw new Error(data.error || "Erro ao ler o cupom");
  return data.cupom || null;
}

async function apiCheckDuplicatas(passagens) {
  const res = await fetch(APPS_SCRIPT_URL, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({ action: "checarDuplicatas", passagens }),
  });
  const data = await res.json();
  if (!data.ok) throw new Error(data.error || "Erro ao checar duplicatas");
  return data.passagens || passagens;
}

async function apiOcrExtrato(base64) {
  const res = await fetch(APPS_SCRIPT_URL, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({ action: "ocrExtrato", image: base64 }),
  });
  const data = await res.json();
  if (!data.ok) throw new Error(data.error || "Erro ao ler o extrato");
  return data.passagens || [];
}

async function apiList(colaborador) {
  const res = await fetch(APPS_SCRIPT_URL, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({ action: "list", colaborador }),
  });
  const data = await res.json();
  if (!data.ok) throw new Error(data.error || "Erro ao listar lançamentos");
  return data.lancamentos || [];
}

async function apiSave(record, colaborador) {
  const res = await fetch(APPS_SCRIPT_URL, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({
      action: "save",
      data: record.data,
      tipo: "Viagem",
      origem: record.origem || "",
      destino: record.destino || "",
      kmInicial: record.kmInicial ?? "",
      kmFinal: record.kmFinal ?? "",
      observacao: record.observacao || "",
      colaborador,
      carro: record.carro || CARRO_PADRAO,
    }),
  });
  const data = await res.json();
  if (!data.ok) throw new Error(data.error || "Erro ao gravar na planilha");
  return data; // { ok, row, taxa }
}

// ─── Imagem: redimensiona antes de enviar (menor, mais rápida, mais barata) ──
function fileToResizedBase64(file, maxDim = 1280, quality = 0.72) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const img = new Image();
      img.onload = () => {
        const scale = Math.min(1, maxDim / Math.max(img.width, img.height));
        const canvas = document.createElement("canvas");
        canvas.width = Math.round(img.width * scale);
        canvas.height = Math.round(img.height * scale);
        canvas.getContext("2d").drawImage(img, 0, 0, canvas.width, canvas.height);
        resolve(canvas.toDataURL("image/jpeg", quality).split(",")[1]);
      };
      img.onerror = reject;
      img.src = reader.result;
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

// ─── GPS → cidade ─────────────────────────────────────────────────────────────
function gpsCidade() {
  return new Promise((resolve) => {
    if (!navigator.geolocation) return resolve(null);
    navigator.geolocation.getCurrentPosition(
      async (pos) => {
        try {
          const { latitude, longitude } = pos.coords;
          const res = await fetch(
            `https://nominatim.openstreetmap.org/reverse?lat=${latitude}&lon=${longitude}&format=json&accept-language=pt-BR`
          );
          const data = await res.json();
          const a = data.address || {};
          const municipio = a.city || a.town || a.village || a.municipality || null;
          if (!municipio) return resolve(null);
          resolve(CITY_GPS_MAP[municipio] || municipio);
        } catch { resolve(null); }
      },
      () => resolve(null),
      { timeout: 8000, maximumAge: 300000 }
    );
  });
}

// ─── Export Excel (formato corporativo preservado) ────────────────────────────
function exportToExcel(records, mesLabel, taxas, colaboradores, solicitante) {
  const sorted = [...records].sort((a, b) => a.data.localeCompare(b.data));
  const wb = XLSX.utils.book_new();
  const rows = [
    ["", "RELATÓRIO DE DESPESAS"],
    [], [],
    ["", "Solicitante", "", solicitante],
    ["", "Inserir CPF", "", CPF.replace(/\D/g, "")],
    ["", "Setor", "", SETOR],
    ["", "Referência", "", mesLabel],
    [],
    ["", "INSERIR DESPESAS DE VIAGEM", "", "", "", "", "", "INSERIR  KM RODADO"],
    ["", "Data", "Dia da semana", "TIPO DE DESPESA", "DESCRIÇÃO DA DESPESA", "VALOR", "", "CIDADE DE ORIGEM", "CIDADE DESTINO", "Km Inicial", "KM Final", "KM RODADO", "R$ KM", "R$ KM TOTAL"],
  ];
  let totalReais = 0;
  sorted.forEach(r => {
    const km = Math.max(0, (r.kmFinal || 0) - (r.kmInicial || 0));
    const taxa = taxaVigente(taxas, colaboradores, solicitante, r.data);
    const total = km * taxa;
    totalReais += total;
    rows.push([
      "", formatDateBR(r.data), weekdayPT(r.data), "", "", "", "",
      r.origem || "", r.destino || "", r.kmInicial || "", r.kmFinal || "",
      km > 0 ? km : 0, km > 0 ? taxa : "", km > 0 ? total : 0,
    ]);
  });
  rows.push(
    [],
    ["", "", "", "", "", "", "", "", "", "", "", "Valor das despesas", "", 0],
    ["", "", "", "", "", "", "", "", "", "", "", "Outros", "", 0],
    ["", "", "", "", "", "", "", "", "", "", "", "Total a Receber", "", totalReais],
  );
  const ws = XLSX.utils.aoa_to_sheet(rows);
  XLSX.utils.book_append_sheet(wb, ws, "Despesas");
  XLSX.writeFile(wb, `Relatorio_KM_${mesLabel.replace(/\s/g, "_")}.xlsx`);
}

// ─── Coerência do odômetro (contra dia anterior e seguinte, MESMO carro) ──────
// Retorna { field, msg } se houver conflito, ou null se estiver ok.
function checkCoherence(records, dateISO, kmIni, kmFin, ignoreId, carro) {
  const carroRef = carro || CARRO_PADRAO;
  const others = records.filter(r =>
    r.data !== dateISO && r.id !== ignoreId && (r.carro || CARRO_PADRAO) === carroRef
  );

  const prev = others
    .filter(r => r.data < dateISO && r.kmFinal != null)
    .sort((a, b) => b.data.localeCompare(a.data))[0];
  const prevIniOnly = prev ? null : others
    .filter(r => r.data < dateISO && r.kmInicial != null)
    .sort((a, b) => b.data.localeCompare(a.data))[0];

  const next = others
    .filter(r => r.data > dateISO && r.kmInicial != null)
    .sort((a, b) => a.data.localeCompare(b.data))[0];

  if (kmIni != null) {
    const ref = prev ? { v: prev.kmFinal, label: "KM final", d: prev.data }
      : prevIniOnly ? { v: prevIniOnly.kmInicial, label: "KM inicial", d: prevIniOnly.data }
      : null;
    if (ref && kmIni < ref.v) {
      return { field: "ini", msg: `KM inicial (${kmIni.toLocaleString("pt-BR")}) é menor que o ${ref.label} de ${formatDateBR(ref.d)} (${ref.v.toLocaleString("pt-BR")}) neste carro. O odômetro não anda pra trás — confira o valor.` };
    }
  }

  if (kmFin != null && next) {
    if (kmFin > next.kmInicial) {
      return { field: "fin", msg: `KM final (${kmFin.toLocaleString("pt-BR")}) é maior que o KM inicial de ${formatDateBR(next.data)} (${next.kmInicial.toLocaleString("pt-BR")}) neste carro. Confira o valor.` };
    }
  }

  if (kmIni != null && kmFin != null && kmFin < kmIni) {
    return { field: "fin", msg: `KM final (${kmFin.toLocaleString("pt-BR")}) menor que o inicial (${kmIni.toLocaleString("pt-BR")}).` };
  }

  return null;
}

// ─── Cálculos de resumo ───────────────────────────────────────────────────────
function kmOf(r) {
  if (r.kmInicial == null || r.kmFinal == null) return 0;
  return Math.max(0, r.kmFinal - r.kmInicial);
}
function isOpen(r) {
  return r.kmInicial == null || r.kmFinal == null;
}
function monthSummary(records, key, taxas, colaboradores, solicitante, despesas, keyFn) {
  const agrupa = keyFn || periodKey;
  const recs = records.filter(r => agrupa(r.data) === key).sort((a, b) => a.data.localeCompare(b.data));
  const trabalho = recs.reduce((s, r) => s + kmOf(r), 0);
  const receberKm = recs.reduce((s, r) => s + kmOf(r) * taxaVigente(taxas, colaboradores, solicitante, r.data), 0);

  // Pedágio do período, e por dia (soma todos os carros daquele dia numa linha só).
  const pedagioPorDia = {};
  const pedagioItensPorDia = {};
  (despesas || []).forEach(d => {
    if ((d.tipo || "").toLowerCase().indexOf("ped") !== 0) return;
    if (agrupa(d.data) !== key) return;
    pedagioPorDia[d.data] = (pedagioPorDia[d.data] || 0) + (Number(d.valor) || 0);
    (pedagioItensPorDia[d.data] = pedagioItensPorDia[d.data] || []).push({ local: d.descricao || "Pedágio", valor: Number(d.valor) || 0 });
  });
  const pedagioMes = Object.values(pedagioPorDia).reduce((s, v) => s + v, 0);

  // Dias com pedágio mas sem apontamento de KM viram linhas "sintéticas".
  const diasComKm = new Set(recs.map(r => r.data));
  const diasSoDeaPedagio = Object.keys(pedagioPorDia)
    .filter(d => !diasComKm.has(d))
    .map(d => ({ id: `pedagio-${d}`, data: d, soPedagio: true }));

  const recsComPedagio = [...recs, ...diasSoDeaPedagio].sort((a, b) => a.data.localeCompare(b.data));

  // KM pessoal é por carro (odômetros diferentes) e depois somado.
  const carros = [...new Set(recs.map(r => r.carro || CARRO_PADRAO))];
  let pessoal = null;
  for (const c of carros) {
    const recsC = recs.filter(r => (r.carro || CARRO_PADRAO) === c);
    const odos = [];
    recsC.forEach(r => {
      if (r.kmInicial != null) odos.push({ d: r.data + "A", v: r.kmInicial });
      if (r.kmFinal != null) odos.push({ d: r.data + "B", v: r.kmFinal });
    });
    odos.sort((a, b) => a.d.localeCompare(b.d));

    // O km rodado ENTRE grupos (ex: da última viagem da semana passada até a
    // primeira desta) é uso pessoal e aparece no grupo em que ele termina:
    // o odômetro da última viagem anterior ao grupo entra como ponto de partida.
    if (odos.length >= 1) {
      const primeiraData = recsC.reduce((m, r) => (m === null || r.data < m ? r.data : m), null);
      let odoAnterior = null;
      records.forEach(r => {
        if ((r.carro || CARRO_PADRAO) !== c) return;
        if (r.data >= primeiraData) return; // só viagens anteriores ao grupo
        const v = r.kmFinal != null ? r.kmFinal : r.kmInicial;
        if (v == null) return;
        if (odoAnterior === null || r.data > odoAnterior.d || (r.data === odoAnterior.d && v > odoAnterior.v)) {
          odoAnterior = { d: r.data, v };
        }
      });
      if (odoAnterior) odos.unshift({ d: odoAnterior.d + "B", v: odoAnterior.v });
    }

    if (odos.length >= 2) {
      const span = odos[odos.length - 1].v - odos[0].v;
      const trabC = recsC.reduce((s, r) => s + kmOf(r), 0);
      pessoal = (pessoal || 0) + Math.max(0, span - trabC);
    }
  }
  return {
    recs: recsComPedagio, trabalho, receber: receberKm + pedagioMes, receberKm, pedagioMes, pedagioPorDia, pedagioItensPorDia,
    pessoal, viagens: recs.filter(r => kmOf(r) > 0).length,
  };
}

// ─── Diálogos do app (substituem os alert/confirm nativos do navegador, que
// mostram "ftjfagro.github.io diz" no título — aqui é "Aviso · Km BTJ") ──────
let _dialogPush = null;
function _abrirDialogo(d) {
  return new Promise(res => {
    if (_dialogPush) _dialogPush({ ...d, res });
    // Fallback improvável (host ainda não montado): usa o nativo pra não engolir o aviso
    else res(d.tipo === "confirmar" ? window.confirm(d.msg) : (window.alert(d.msg), true));
  });
}
function avisar(msg) { return _abrirDialogo({ tipo: "aviso", msg }); }
function confirmar(msg) { return _abrirDialogo({ tipo: "confirmar", msg }); }

function DialogHost() {
  const [fila, setFila] = useState([]);
  useEffect(() => {
    _dialogPush = d => setFila(f => [...f, d]);
    return () => { _dialogPush = null; };
  }, []);
  const atual = fila[0];
  if (!atual) return null;
  const fechar = v => { atual.res(v); setFila(f => f.slice(1)); };
  return (
    <div className="fixed inset-0 flex items-center justify-center p-5" style={{ background: "rgba(0,0,0,0.45)", zIndex: 60 }}>
      <div className="bg-white rounded-xl p-4 w-full max-w-xs shadow-lg">
        <p className="text-[11px] font-bold uppercase tracking-wide mb-1" style={{ color: BTJ_BLUE }}>
          {atual.tipo === "confirmar" ? "Confirmação · Km BTJ" : "Aviso · Km BTJ"}
        </p>
        <p className="text-sm text-gray-700 whitespace-pre-wrap mb-3">{atual.msg}</p>
        <div className="flex gap-1.5 justify-end">
          {atual.tipo === "confirmar" && (
            <button onClick={() => fechar(false)} className="rounded-lg px-4 py-2 text-xs text-gray-600 border border-gray-200">Cancelar</button>
          )}
          <button onClick={() => fechar(true)} className="rounded-lg px-5 py-2 text-xs font-semibold text-white" style={{ background: BTJ_BLUE }}>OK</button>
        </div>
      </div>
    </div>
  );
}

// ─── UI base ──────────────────────────────────────────────────────────────────
function Card({ children, className = "" }) {
  return <div className={`bg-white rounded-2xl border border-gray-100 shadow-sm ${className}`}>{children}</div>;
}

function KmBox({ label, value, state, onCamera, onGallery, onManual, loading, error }) {
  // state: "done" | "pending" | "queued"
  const styles = {
    done: { bg: "bg-emerald-50", labelC: "text-emerald-700", valC: "text-emerald-900" },
    pending: { bg: "bg-amber-50", labelC: "text-amber-700", valC: "text-amber-600" },
    queued: { bg: "bg-sky-50", labelC: "text-sky-700", valC: "text-sky-600" },
  }[state];
  const border = error ? "border-2 border-red-500" : "";
  return (
    <div className={`flex-1 ${styles.bg} ${border} rounded-xl p-2.5 text-center`}>
      <p className={`text-[11px] ${error ? "text-red-600" : styles.labelC}`}>{label}</p>
      <input
        type="number"
        inputMode="numeric"
        placeholder="—"
        value={value ?? ""}
        onChange={e => onManual(e.target.value === "" ? null : Number(e.target.value))}
        onKeyDown={e => { if (e.key === "Enter") e.currentTarget.blur(); }}
        className={`w-full bg-transparent text-center text-lg font-semibold ${error ? "text-red-600" : styles.valC} focus:outline-none`}
      />
      {state === "queued" ? (
        <div className="bg-sky-200 rounded-lg py-1 text-[11px] text-sky-900">☁ aguardando conexão</div>
      ) : (
        <div className="flex gap-1">
          <button
            onClick={onCamera}
            disabled={loading}
            className={`flex-1 rounded-lg py-1 text-sm ${state === "done" ? "bg-white" : "bg-amber-400"}`}
            aria-label="Tirar foto"
            title="Tirar foto"
          >
            {loading ? "⟳" : "📷"}
          </button>
          <button
            onClick={onGallery}
            disabled={loading}
            className={`flex-1 rounded-lg py-1 text-sm ${state === "done" ? "bg-white" : "bg-amber-400"}`}
            aria-label="Escolher da galeria"
            title="Escolher da galeria"
          >
            🖼️
          </button>
        </div>
      )}
    </div>
  );
}

// ─── Tela: Nova Despesa (manual) ──────────────────────────────────────────────
const TIPOS_DESPESA = ["Alimentação", "Mobilidade", "Hotel", "Pedágio", "Outros"];
const TIPO_COR = {
  "Alimentação": { bg: "#E1F5EE", fg: "#085041" },
  "Mobilidade": { bg: "#E6F1FB", fg: "#0C447C" },
  "Hotel": { bg: "#EEEDFE", fg: "#3C3489" },
  "Pedágio": { bg: "#FAEEDA", fg: "#633806" },
  "Outros": { bg: "#F1EFE8", fg: "#444441" },
};

const NOVO_CARRO_VALUE = "__novo_carro__";
const logoUrl = `${import.meta.env.BASE_URL}icons/icon.svg`;

// Input de valor "estilo caixa registradora": digita só números, as duas
// últimas casas viram centavos automaticamente (ex: 1000 -> R$ 10,00).
function digitsParaReais(digits) {
  return (parseInt(digits || "0", 10)) / 100;
}
function ValorInput({ digits, onDigitsChange, className, placeholder }) {
  return (
    <input
      type="text" inputMode="numeric"
      value={digits ? digitsParaReais(digits).toLocaleString("pt-BR", { minimumFractionDigits: 2 }) : ""}
      placeholder={placeholder || "0,00"}
      onChange={e => onDigitsChange(e.target.value.replace(/\D/g, "").replace(/^0+(?=\d)/, ""))}
      onKeyDown={e => { if (e.key === "Enter") e.currentTarget.blur(); }}
      className={className}
    />
  );
}


function CadastrarCarroModal({ colaborador, onSaved, onCancel }) {
  const [modelo, setModelo] = useState("");
  const [placa, setPlaca] = useState("");
  const [saving, setSaving] = useState(false);

  async function salvar() {
    if (!modelo.trim()) { avisar("Digite o modelo/apelido do carro."); return; }
    if (!placa.trim()) { avisar("Digite a placa."); return; }
    const carroCompleto = `${modelo.trim()} ${placa.trim().toUpperCase()}`;
    setSaving(true);
    try {
      await apiAddCarro(carroCompleto, colaborador);
      onSaved(carroCompleto);
    } catch (e) {
      avisar("Erro ao cadastrar carro: " + (e.message || ""));
    } finally {
      setSaving(false);
    }
  }

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4" style={{ background: "rgba(0,0,0,0.45)" }}>
      <div className="bg-white rounded-xl p-4 w-full max-w-sm">
        <p className="font-semibold text-sm mb-3" style={{ color: BTJ_NAVY }}>🚗 Cadastrar novo carro</p>
        <p className="text-[11px] text-gray-500 mb-0.5">Modelo / apelido</p>
        <input type="text" value={modelo} onChange={e => setModelo(e.target.value)}
          onKeyDown={e => { if (e.key === "Enter") e.currentTarget.blur(); }}
          placeholder="Ex: Fiat Argo" autoFocus
          className="w-full border border-gray-200 rounded-lg px-2 py-1.5 text-sm mb-2.5" />
        <p className="text-[11px] text-gray-500 mb-0.5">Placa</p>
        <input type="text" value={placa} onChange={e => setPlaca(e.target.value.toUpperCase())}
          onKeyDown={e => { if (e.key === "Enter") e.currentTarget.blur(); }}
          placeholder="Ex: ABC1D23"
          className="w-full border border-gray-200 rounded-lg px-2 py-1.5 text-sm mb-3" />
        <div className="flex gap-2">
          <button onClick={salvar} disabled={saving}
            className="flex-[2] rounded-lg py-2 text-sm font-semibold text-white disabled:opacity-60" style={{ background: BTJ_BLUE }}>
            {saving ? "Cadastrando..." : "Cadastrar e usar"}
          </button>
          <button onClick={onCancel} disabled={saving} className="flex-1 rounded-lg py-2 text-sm text-gray-600 border border-gray-200">Cancelar</button>
        </div>
      </div>
    </div>
  );
}

function DespesaManual({ carros, carroInicial, limites, faixa, colaborador, travado, onNovoCarro, onSaved, onCancel }) {
  const [data, setData] = useState(todayISO());
  const [carro, setCarro] = useState(carroInicial);
  const [mostrarCadastroCarro, setMostrarCadastroCarro] = useState(false);
  const [tipo, setTipo] = useState("Alimentação");
  const [valorDigits, setValorDigits] = useState("");
  const valor = valorDigits ? digitsParaReais(valorDigits) : 0;
  const [descricao, setDescricao] = useState("");
  const [compB64, setCompB64] = useState(null);
  const [compNome, setCompNome] = useState("");
  const [saving, setSaving] = useState(false);
  const camRef = useRef(null);
  const galRef = useRef(null);
  const camCupomRef = useRef(null);
  const galCupomRef = useRef(null);

  const [lendoCupom, setLendoCupom] = useState(false);
  const [usarIA, setUsarIA] = useState(true);

  // ─── Rateio ────────────────────────────────────────────────────────────
  const [pessoas, setPessoas] = useState(1);
  const [comQuem, setComQuem] = useState("");
  const [escolhaExcedente, setEscolhaExcedente] = useState(null); // null | "elegivel" | "total"
  const [justificativa, setJustificativa] = useState("");

  const precisaRateio = CATEGORIAS_RATEAVEIS.includes(tipo);
  const limiteAplicavel = precisaRateio ? limiteVigente(limites, faixa, tipo, data) : null;
  const nPessoas = Math.max(1, Number(pessoas) || 1);
  const valorNum = Number(valor) || 0;
  const valorPorPessoa = precisaRateio && valorNum > 0 ? valorNum / nPessoas : null;
  const excede = limiteAplicavel != null && valorPorPessoa != null && valorPorPessoa > limiteAplicavel + 0.005;

  useEffect(() => { setEscolhaExcedente(null); setJustificativa(""); }, [tipo, valor, pessoas]);

  async function pickComprovante(file, lerIA) {
    if (!file) return;
    try {
      const b64 = await fileToResizedBase64(file, 1600, 0.7);
      setCompB64(b64);
      setCompNome(file.name || "comprovante.jpg");
      if (lerIA) {
        setLendoCupom(true);
        try {
          const c = await apiOcrCupom(b64);
          if (c) {
            if (c.valor != null) setValorDigits(String(Math.round(Number(c.valor) * 100)));
            if (c.data) setData(String(c.data).slice(0, 10));
            if (c.descricao) setDescricao(c.descricao);
          } else {
            avisar("Não consegui ler o cupom automaticamente. Preencha os campos na mão (a foto foi guardada como comprovante).");
          }
        } catch (e) {
          avisar("Falha ao ler o cupom: " + (e.message || "") + ". A foto foi guardada; preencha na mão.");
        } finally {
          setLendoCupom(false);
        }
      }
    } catch { avisar("Não consegui processar a imagem do comprovante."); }
  }

  async function salvar() {
    if (!valor || Number(valor) <= 0) { avisar("Informe o valor da despesa."); return; }
    if (travado && travado(data)) { avisar("Este período já foi fechado e enviado — não dá mais pra lançar despesas nele."); return; }
    if (excede && !escolhaExcedente) {
      avisar(`O valor por pessoa (R$ ${valorPorPessoa.toLocaleString("pt-BR", { minimumFractionDigits: 2 })}) passa do limite de R$ ${limiteAplicavel.toLocaleString("pt-BR", { minimumFractionDigits: 2 })}. Escolha uma das opções abaixo antes de salvar.`);
      return;
    }
    if (excede && escolhaExcedente === "total" && !justificativa.trim()) {
      avisar("Para registrar o valor total acima do limite, escreva uma justificativa.");
      return;
    }
    const valorTotalPago = Number(valor);
    const valorFinal = (excede && escolhaExcedente === "elegivel")
      ? Number((limiteAplicavel * nPessoas).toFixed(2))
      : valorTotalPago;
    setSaving(true);
    try {
      await apiSaveDespesa({
        data, carro, tipo, valor: valorFinal,
        descricao, comprovanteImage: compB64 || undefined, origem: "manual",
        pessoasRateio: precisaRateio ? nPessoas : 1,
        rateioCom: precisaRateio ? comQuem : "",
        valorTotalPago,
        justificativa: (excede && escolhaExcedente === "total") ? justificativa.trim() : "",
        colaborador,
      });
      onSaved();
    } catch (e) {
      avisar("Erro ao salvar: " + (e.message || "tente novamente."));
    } finally {
      setSaving(false);
    }
  }

  return (
    <Card className="mt-2.5 p-3.5">
      <div className="mb-2.5">
        <p className="text-[11px] text-gray-500 mb-0.5">Data</p>
        <input type="date" value={data} max={todayISO()} onChange={e => setData(e.target.value)}
          className="w-full border border-gray-200 rounded-lg px-2 py-1.5 text-sm" />
      </div>

      <p className="text-[11px] text-gray-500 mb-1">Tipo</p>
      <div className="flex gap-1.5 flex-wrap mb-2.5">
        {TIPOS_DESPESA.map(t => (
          <button key={t} onClick={() => setTipo(t)}
            className="text-xs px-3 py-1.5 rounded-lg font-medium"
            style={t === tipo ? { background: BTJ_BLUE, color: "#fff" } : { background: "#F1EFE8", color: "#5F5E5A" }}>
            {t}
          </button>
        ))}
      </div>

      {tipo === "Pedágio" && (
        <div className="mb-2.5">
          <p className="text-[11px] text-gray-500 mb-0.5">🚗 Carro (pedágio avulso)</p>
          <select value={carro} onChange={e => {
              if (e.target.value === NOVO_CARRO_VALUE) { setMostrarCadastroCarro(true); return; }
              setCarro(e.target.value);
            }}
            className="w-full border border-gray-200 rounded-lg px-2 py-1.5 text-sm bg-white">
            {carros.map(c => <option key={c} value={c}>{c}</option>)}
            <option value={NOVO_CARRO_VALUE}>+ Outro (cadastrar novo)</option>
          </select>
        </div>
      )}

      {mostrarCadastroCarro && (
        <CadastrarCarroModal
          colaborador={colaborador}
          onCancel={() => setMostrarCadastroCarro(false)}
          onSaved={(novoCarro) => { onNovoCarro(novoCarro); setCarro(novoCarro); setMostrarCadastroCarro(false); }}
        />
      )}

      <div className="flex gap-2 mb-2.5">
        <div className="flex-1">
          <p className="text-[11px] text-gray-500 mb-0.5">Valor (R$)</p>
          <ValorInput digits={valorDigits} onDigitsChange={setValorDigits}
            className="w-full border border-gray-200 rounded-lg px-2 py-1.5 text-sm" />
        </div>
        <div className="flex-[1.4]">
          <p className="text-[11px] text-gray-500 mb-0.5">
            {tipo === "Mobilidade" ? "Descrição (Passagens, Táxi, Uber, Estacionamento...)" : "Descrição (opcional)"}
          </p>
          <input type="text" value={descricao} onChange={e => setDescricao(e.target.value)}
            onKeyDown={e => { if (e.key === "Enter") e.currentTarget.blur(); }}
            placeholder={tipo === "Mobilidade" ? "Ex: Uber até o cliente" : ""}
            className="w-full border border-gray-200 rounded-lg px-2 py-1.5 text-sm" />
        </div>
      </div>


      {precisaRateio && (
        <div className="mb-2.5">
          <div className="flex gap-2">
            <div className="w-24">
              <p className="text-[11px] text-gray-500 mb-0.5">Rateado entre</p>
              <input type="number" min="1" inputMode="numeric" value={pessoas}
                onChange={e => setPessoas(e.target.value === "" ? "" : Math.max(1, Number(e.target.value)))}
                onKeyDown={e => { if (e.key === "Enter") e.currentTarget.blur(); }}
                className="w-full border border-gray-200 rounded-lg px-2 py-1.5 text-sm text-center" />
            </div>
            <div className="flex-1">
              <p className="text-[11px] text-gray-500 mb-0.5">Com quem? (colega, cliente...)</p>
              <input type="text" value={comQuem} onChange={e => setComQuem(e.target.value)}
                onKeyDown={e => { if (e.key === "Enter") e.currentTarget.blur(); }}
                placeholder="opcional" className="w-full border border-gray-200 rounded-lg px-2 py-1.5 text-sm" />
            </div>
          </div>
          {nPessoas > 1 && valorPorPessoa != null && (
            <p className="text-[10px] text-gray-400 mt-1">
              R$ {valorNum.toLocaleString("pt-BR", { minimumFractionDigits: 2 })} ÷ {nPessoas} pessoas = R$ {valorPorPessoa.toLocaleString("pt-BR", { minimumFractionDigits: 2 })} por pessoa
              {limiteAplicavel != null && ` (limite: R$ ${limiteAplicavel.toLocaleString("pt-BR", { minimumFractionDigits: 2 })})`}
            </p>
          )}
        </div>
      )}

      {excede && (
        <div className="rounded-lg p-3 mb-2.5" style={{ background: "#FEF3E2", border: "1px solid #F5C97A" }}>
          <p className="text-xs font-medium mb-2" style={{ color: "#854F0B" }}>
            ⚠ R$ {valorPorPessoa.toLocaleString("pt-BR", { minimumFractionDigits: 2 })} por pessoa passa do limite de R$ {limiteAplicavel.toLocaleString("pt-BR", { minimumFractionDigits: 2 })}. Como quer registrar?
          </p>
          <div className="flex gap-1.5 mb-2">
            <button onClick={() => setEscolhaExcedente("elegivel")}
              className="flex-1 rounded-lg py-1.5 text-xs font-medium"
              style={escolhaExcedente === "elegivel" ? { background: BTJ_BLUE, color: "#fff" } : { background: "#fff", border: "1px solid #F5C97A", color: "#854F0B" }}>
              Só o elegível (R$ {(limiteAplicavel * nPessoas).toLocaleString("pt-BR", { minimumFractionDigits: 2 })})
            </button>
            <button onClick={() => setEscolhaExcedente("total")}
              className="flex-1 rounded-lg py-1.5 text-xs font-medium"
              style={escolhaExcedente === "total" ? { background: BTJ_BLUE, color: "#fff" } : { background: "#fff", border: "1px solid #F5C97A", color: "#854F0B" }}>
              Valor total (R$ {valorNum.toLocaleString("pt-BR", { minimumFractionDigits: 2 })})
            </button>
          </div>
          {escolhaExcedente === "total" && (
            <input type="text" value={justificativa} onChange={e => setJustificativa(e.target.value)}
              onKeyDown={e => { if (e.key === "Enter") e.currentTarget.blur(); }}
              placeholder="Justificativa do valor acima do limite..."
              className="w-full border rounded-lg px-2 py-1.5 text-xs" style={{ borderColor: "#F5C97A" }} />
          )}
        </div>
      )}

      <div className="flex items-center justify-between mb-2">
        <p className="text-[11px]"><span style={{ color: "#C62A2F" }}>* </span><span className="text-gray-500">Comprovante (obrigatório)</span></p>
        <button onClick={() => setUsarIA(v => !v)} className="flex items-center gap-1.5">
          <span className="text-[11px] font-medium" style={{ color: usarIA ? BTJ_NAVY : "#9C9C96" }}>✨ Preencher com IA</span>
          <span className="relative inline-block w-8 h-[18px] rounded-full transition-colors" style={{ background: usarIA ? BTJ_BLUE : "#D8D6CE" }}>
            <span className="absolute top-[2px] w-[14px] h-[14px] rounded-full bg-white transition-all" style={{ left: usarIA ? "16px" : "2px" }} />
          </span>
        </button>
      </div>
      <div className="flex gap-1.5 mb-3">
        <button onClick={() => camRef.current?.click()} disabled={lendoCupom}
          className="flex-1 rounded-lg py-2.5 text-sm font-medium text-white disabled:opacity-60" style={{ background: BTJ_BLUE }}>
          {lendoCupom ? "⟳ lendo..." : "📷 Tirar foto"}
        </button>
        <button onClick={() => galRef.current?.click()} disabled={lendoCupom}
          className="flex-1 rounded-lg py-2.5 text-sm font-medium text-white disabled:opacity-60" style={{ background: BTJ_BLUE }}>
          {lendoCupom ? "⟳..." : "🖼️ Da galeria"}
        </button>
      </div>
      <input ref={camRef} type="file" accept="image/*" capture="environment" className="hidden" onChange={e => { pickComprovante(e.target.files[0], usarIA); e.target.value = ""; }} />
      <input ref={galRef} type="file" accept="image/*" className="hidden" onChange={e => { pickComprovante(e.target.files[0], usarIA); e.target.value = ""; }} />
      {compNome && <p className="text-[10px] mb-3" style={{ color: "#0F6E56" }}>🧾 {compNome} — será enviado ao Drive ao salvar</p>}
      {!compNome && <p className="text-[10px] mb-3" style={{ color: "#C62A2F" }}>⚠ Anexe a foto do comprovante pra poder salvar</p>}

      <div className="flex gap-2">
        <button onClick={salvar} disabled={saving || !compNome}
          className="flex-[2] rounded-xl py-2.5 text-sm font-semibold text-white disabled:opacity-60"
          style={{ background: BTJ_BLUE }}>
          {saving ? "Salvando..." : "Salvar despesa"}
        </button>
        <button onClick={onCancel} className="flex-1 rounded-xl py-2.5 text-sm text-gray-600 border border-gray-200">Cancelar</button>
      </div>
    </Card>
  );
}

// ─── Tela: Importar Extrato do Tag (pedágio/estacionamento cobrado no tag) ────
function ImportarExtrato({ carros, carroInicial, records, colaborador, travado, onNovoCarro, onDone, onCancel }) {
  const [carro, setCarro] = useState(carroInicial);
  const [mostrarCadastroCarro, setMostrarCadastroCarro] = useState(false);
  const [passagens, setPassagens] = useState(null); // null = ainda não leu
  const [sel, setSel] = useState({});
  const [loading, setLoading] = useState(false);
  const [saving, setSaving] = useState(false);
  const galRef = useRef(null);

  // Dias em que existe viagem de trabalho lançada para o carro deste extrato
  // (usado pra pré-marcar só as passagens que batem com um dia de trabalho).
  const diasComViagem = useMemo(() => {
    const s = new Set();
    (records || []).forEach(r => {
      if ((r.carro || CARRO_PADRAO) === carro) s.add(r.data);
    });
    return s;
  }, [records, carro]);

  async function lerExtrato(fileList) {
    const files = Array.from(fileList || []);
    if (!files.length) return;
    if (carro === NOVO_CARRO_VALUE) { avisar("Escolha ou cadastre um carro antes de importar."); return; }
    setLoading(true);
    try {
      let todas = [];
      for (const file of files) {
        const b64 = await fileToResizedBase64(file, 1600, 0.7);
        const lista = await apiOcrExtrato(b64);
        todas = todas.concat(lista);
      }
      // Sem dedupe entre passagens do mesmo print: passagens iguais (mesma data/local/valor)
      // podem ser genuinamente diferentes (duas passadas pela mesma praça no mesmo dia).
      todas.sort((a, b) => String(a.data).localeCompare(String(b.data)));
      if (!todas.length) { avisar("Não encontrei passagens nesses prints. Tente imagens mais nítidas."); return; }
      // Checa contra o que já está na planilha (essa sim é uma duplicata real).
      let comCheck = todas;
      try { comCheck = await apiCheckDuplicatas(todas); } catch { /* segue sem o check se falhar */ }
      setPassagens(comCheck);
      const inicial = {};
      comCheck.forEach((p, i) => {
        // Pré-marca só se: não foi lançada antes E existe viagem de trabalho
        // desse carro nesse dia. Sem viagem correspondente = vem desmarcada
        // (provável uso pessoal), pra você decidir.
        inicial[i] = !p.jaLancado && diasComViagem.has(p.data);
      });
      setSel(inicial);
    } catch (e) {
      avisar("Erro ao ler o extrato: " + (e.message || "tente novamente."));
    } finally {
      setLoading(false);
    }
  }

  const totalSel = passagens ? passagens.reduce((s, p, i) => s + (sel[i] ? Number(p.valor) || 0 : 0), 0) : 0;
  const qtdSel = passagens ? passagens.filter((_, i) => sel[i]).length : 0;

  async function lancar() {
    setSaving(true);
    try {
      const marcadas = passagens.filter((_, i) => sel[i]);
      const bloqueadas = travado ? marcadas.filter(p => travado(p.data)) : [];
      if (bloqueadas.length > 0) {
        avisar("Algumas passagens são de período já fechado e enviado — desmarque-as antes de lançar: " + bloqueadas.map(p => p.data).join(", "));
        setSaving(false);
        return;
      }
      for (const p of marcadas) {
        await apiSaveDespesa({
          data: p.data, carro, tipo: "Pedágio",
          valor: Number(p.valor) || 0, descricao: p.local || "", origem: "extrato",
          colaborador,
        });
      }
      avisar(`${marcadas.length} lançamento(s) do tag lançado(s) na planilha (R$ ${totalSel.toLocaleString("pt-BR", { minimumFractionDigits: 2 })}).`);
      onDone();
    } catch (e) {
      avisar("Erro ao lançar: " + (e.message || "tente novamente."));
    } finally {
      setSaving(false);
    }
  }

  return (
    <Card className="mt-2.5 p-3.5">
      <div className="mb-2.5">
        <p className="text-[11px] text-gray-500 mb-0.5">🚗 Carro deste extrato</p>
        <select value={carro} onChange={e => {
            if (e.target.value === NOVO_CARRO_VALUE) { setMostrarCadastroCarro(true); return; }
            setCarro(e.target.value); saveLastCar(e.target.value);
          }}
          className="w-full border border-gray-200 rounded-lg px-2 py-1.5 text-sm bg-white">
          {carros.map(c => <option key={c} value={c}>{c}</option>)}
          <option value={NOVO_CARRO_VALUE}>+ Outro (cadastrar novo)</option>
        </select>
        {mostrarCadastroCarro && (
          <CadastrarCarroModal
          colaborador={colaborador}
            onCancel={() => setMostrarCadastroCarro(false)}
            onSaved={(novoCarro) => { onNovoCarro(novoCarro); setCarro(novoCarro); saveLastCar(novoCarro); setMostrarCadastroCarro(false); }}
          />
        )}
      </div>

      {!passagens && (
        <>
          <p className="text-xs text-gray-500 mb-2">Mande o(s) print(s) do extrato do tag (C6, Veloe, Sem Parar etc.) — pode selecionar várias fotos de uma vez da galeria. Cobre pedágio e outras cobranças do tag (ex: estacionamento).</p>
          <button onClick={() => galRef.current?.click()} disabled={loading}
            className="w-full rounded-lg py-2.5 text-sm bg-amber-400">
            {loading ? "⟳ lendo..." : "🖼️ Escolher fotos da galeria (várias)"}
          </button>
          <input ref={galRef} type="file" accept="image/*" multiple className="hidden" onChange={e => { lerExtrato(e.target.files); e.target.value = ""; }} />
        </>
      )}

      {passagens && (
        <>
          <div className="flex justify-between items-center mb-2">
            <span className="text-[11px] text-gray-500">Desmarque as de uso pessoal</span>
            <span className="text-[11px] font-medium" style={{ color: "#185FA5" }}>{qtdSel} de {passagens.length}</span>
          </div>
          <div className="border border-gray-100 rounded-xl overflow-hidden mb-3">
            {passagens.map((p, i) => {
              const temViagem = diasComViagem.has(p.data);
              const semViagem = !p.jaLancado && !temViagem;
              return (
              <div key={i} className={`border-b border-gray-50 last:border-b-0 ${p.jaLancado ? "bg-red-50" : semViagem ? "bg-amber-50" : ""}`}>
                <button onClick={() => setSel(s => ({ ...s, [i]: !s[i] }))}
                  className="w-full flex items-center gap-2.5 px-3 py-2 text-left">
                  <div className="w-[18px] h-[18px] rounded-md flex items-center justify-center text-xs text-white shrink-0"
                    style={{ background: sel[i] ? BTJ_BLUE : "transparent", border: sel[i] ? "none" : "1.5px solid #CFCFC8" }}>
                    {sel[i] ? "✓" : ""}
                  </div>
                  <div className="flex-1">
                    <p className={`text-xs ${sel[i] ? "text-gray-800" : "text-gray-400 line-through"} ${p.jaLancado ? "!text-red-700" : semViagem ? "!text-amber-700" : ""}`}>
                      {p.local || "Pedágio"}{p.jaLancado ? " · já lançado antes" : semViagem ? " · sem viagem nesse dia" : ""}
                    </p>
                    <p className={`text-[10px] ${sel[i] ? "text-gray-500" : "text-gray-400"}`}>{formatDateShort(p.data)} · {weekdayAbrev(p.data)}</p>
                  </div>
                  <span className={`text-xs font-semibold ${sel[i] ? "" : "text-gray-400 line-through"}`} style={sel[i] ? { color: "#04342C" } : {}}>
                    R$ {(Number(p.valor) || 0).toLocaleString("pt-BR", { minimumFractionDigits: 2 })}
                  </span>
                </button>
                {p.jaLancado && (
                  <p className="text-[10px] text-red-700 px-3 pb-2 -mt-1">⚠ Uma passagem igual (mesma data, local e valor) já está na planilha. Marque só se for outra passagem real.</p>
                )}
                {semViagem && (
                  <p className="text-[10px] text-amber-700 px-3 pb-2 -mt-1">🚗 Não há viagem de trabalho lançada para {carro} nesse dia. Marque só se for mesmo uso de trabalho.</p>
                )}
              </div>
              );
            })}
          </div>
          <div className="flex justify-between items-center mb-3">
            <span className="text-sm font-semibold text-gray-800">Total selecionado</span>
            <span className="text-base font-bold" style={{ color: BTJ_NAVY }}>R$ {totalSel.toLocaleString("pt-BR", { minimumFractionDigits: 2 })}</span>
          </div>
          <div className="flex gap-2">
            <button onClick={lancar} disabled={saving || qtdSel === 0}
              className="flex-[2] rounded-xl py-2.5 text-sm font-semibold text-white disabled:opacity-60"
              style={{ background: BTJ_BLUE }}>
              {saving ? "Lançando..." : `Lançar ${qtdSel} lançamento(s)`}
            </button>
            <button onClick={onCancel} className="flex-1 rounded-xl py-2.5 text-sm text-gray-600 border border-gray-200">Cancelar</button>
          </div>
        </>
      )}
    </Card>
  );
}

// ─── Tela: Gestão de Despesas (Pedágios ou Outras) ───────────────────────────
function GestaoDespesas({ titulo, icone, despesas, carros, travado, onChange, onAdd, addLabel }) {
  // Agrupa por período 26→25, últimos 3
  const byPeriod = {};
  despesas.forEach(d => {
    const k = periodKey(d.data);
    (byPeriod[k] = byPeriod[k] || []).push(d);
  });
  const keys = Object.keys(byPeriod).sort().reverse().slice(0, 3);

  const [expanded, setExpanded] = useState(keys[0] || null);
  const [edit, setEdit] = useState(null); // { id, valor, descricao, data }
  const [busy, setBusy] = useState(false);

  async function salvarEdicao() {
    setBusy(true);
    try {
      await apiUpdateDespesa({ id: edit.id, valor: Number(edit.valor) || 0, descricao: edit.descricao, data: edit.data });
      setEdit(null);
      await onChange();
    } catch (e) { avisar("Erro ao salvar: " + (e.message || "")); }
    finally { setBusy(false); }
  }
  async function excluir(id) {
    if (!(await confirmar("Excluir esta despesa? Não dá pra desfazer."))) return;
    setBusy(true);
    try { await apiDeleteDespesa(id); await onChange(); }
    catch (e) { avisar("Erro ao excluir: " + (e.message || "")); }
    finally { setBusy(false); }
  }

  return (
    <Card className="mt-2.5 p-3.5">
      <div className="flex items-center justify-between mb-3">
        <h2 className="font-semibold text-gray-800">{icone} {titulo}</h2>
        <button onClick={onAdd} className="text-xs font-medium px-2.5 py-1.5 rounded-lg" style={{ background: BTJ_BLUE, color: "#fff" }}>
          {addLabel}
        </button>
      </div>

      {keys.length === 0 && <p className="text-sm text-gray-400 text-center py-6">Nenhuma despesa lançada ainda.</p>}

      <div className="space-y-2">
        {keys.map(k => {
          const itens = byPeriod[k].sort((a, b) => b.data.localeCompare(a.data));
          const total = itens.reduce((s, d) => s + (Number(d.valor) || 0), 0);
          const aberto = expanded === k;
          return (
            <div key={k} className="border border-gray-100 rounded-xl overflow-hidden">
              <button onClick={() => setExpanded(aberto ? null : k)} className="w-full flex items-center justify-between px-3.5 py-3 text-left">
                <div>
                  <p className="text-sm font-semibold" style={{ color: BTJ_NAVY }}>{periodLabel(k)}</p>
                  <p className="text-[11px] text-gray-400">{itens.length} lançamento(s)</p>
                </div>
                <div className="flex items-center gap-2">
                  <span className="text-sm font-semibold" style={{ color: BTJ_BLUE }}>R$ {total.toLocaleString("pt-BR", { minimumFractionDigits: 2 })}</span>
                  <span className="text-gray-400 text-xs">{aberto ? "▲" : "▼"}</span>
                </div>
              </button>
              {aberto && (
                <div className="border-t border-gray-100">
                  {itens.map(d => (
                    <div key={d.id} className="border-b border-gray-50 last:border-b-0">
                      {edit?.id === d.id ? (
                        <div className="px-3.5 py-3" style={{ background: "#F8FAFC" }}>
                          <p className="text-[11px] font-medium mb-2" style={{ color: "#185FA5" }}>Editando · {d.tipo}{d.tipo === "Pedágio" && d.carro ? ` · ${d.carro.split(" ")[0]}` : ""}</p>
                          <div className="flex gap-1.5 mb-2">
                            <div className="flex-1">
                              <p className="text-[9px] text-gray-500 mb-0.5">Data</p>
                              <input type="date" value={edit.data} onChange={e => setEdit({ ...edit, data: e.target.value })}
                                className="w-full border border-gray-200 rounded-lg px-2 py-1.5 text-xs" />
                            </div>
                            <div className="flex-1">
                              <p className="text-[9px] text-gray-500 mb-0.5">Valor</p>
                              <input type="number" inputMode="decimal" value={edit.valor} onChange={e => setEdit({ ...edit, valor: e.target.value })}
                                onKeyDown={e => { if (e.key === "Enter") e.currentTarget.blur(); }}
                                className="w-full border border-gray-200 rounded-lg px-2 py-1.5 text-xs" />
                            </div>
                          </div>
                          <input type="text" value={edit.descricao} placeholder="descrição" onChange={e => setEdit({ ...edit, descricao: e.target.value })}
                            onKeyDown={e => { if (e.key === "Enter") e.currentTarget.blur(); }}
                            className="w-full border border-gray-200 rounded-lg px-2 py-1.5 text-xs mb-2" />
                          <div className="flex gap-1.5">
                            <button onClick={salvarEdicao} disabled={busy} className="flex-[2] rounded-lg py-2 text-xs font-medium text-white" style={{ background: BTJ_BLUE }}>Salvar</button>
                            <button onClick={() => setEdit(null)} className="flex-1 rounded-lg py-2 text-xs text-gray-600 border border-gray-200">Cancelar</button>
                            <button onClick={() => excluir(d.id)} disabled={busy} className="flex-1 rounded-lg py-2 text-xs" style={{ background: "#FDECEC", color: "#C62A2F" }}>🗑</button>
                          </div>
                        </div>
                      ) : (
                        <button onClick={() => { if (travado && travado(d.data)) { avisar("Período já enviado — edição travada."); return; } setEdit({ id: d.id, valor: d.valor, descricao: d.descricao, data: d.data }); }}
                          className="w-full flex items-center justify-between px-3.5 py-2 text-left">
                          <div className="min-w-0">
                            <p className="text-xs text-gray-800">{formatDateShort(d.data)} · {weekdayAbrev(d.data)}{d.tipo === "Pedágio" && d.carro ? ` · ${d.carro.split(" ")[0]}` : ""}{d.descricao ? ` · ${d.descricao}` : ""}</p>
                            <div className="flex items-center gap-1.5 mt-1 flex-wrap">
                              <span className="text-[10px] px-1.5 py-0.5 rounded"
                                style={{ background: (TIPO_COR[d.tipo] || TIPO_COR["Outros"]).bg, color: (TIPO_COR[d.tipo] || TIPO_COR["Outros"]).fg }}>
                                {d.tipo}{d.tipo === "Pedágio" && d.origem === "extrato" ? " · importado" : ""}
                              </span>
                              {d.comprovante && <a href={String(d.comprovante).split(" | ")[0]} target="_blank" rel="noreferrer" onClick={e => e.stopPropagation()} className="text-[10px]" style={{ color: BTJ_BLUE }}>🧾 comprovante</a>}
                            </div>
                          </div>
                          <span className="text-xs font-semibold shrink-0" style={{ color: "#04342C" }}>R$ {(Number(d.valor) || 0).toLocaleString("pt-BR", { minimumFractionDigits: 2 })} ✎</span>
                        </button>
                      )}
                    </div>
                  ))}
                </div>
              )}
            </div>
          );
        })}
      </div>
    </Card>
  );
}

// ─── Tela: Gestão de Pedágios (agrupado por dia, com a rota do km daquele dia) ─
function GestaoPedagios({ despesas, records, travado, onChange, onAdd }) {
  const byPeriod = {};
  despesas.forEach(d => {
    const k = periodKey(d.data);
    (byPeriod[k] = byPeriod[k] || []).push(d);
  });
  const periodos = Object.keys(byPeriod).sort().reverse().slice(0, 3);

  const [expandedPeriod, setExpandedPeriod] = useState(periodos[0] || null);
  const [edit, setEdit] = useState(null);
  const [busy, setBusy] = useState(false);

  async function salvarEdicao() {
    setBusy(true);
    try {
      await apiUpdateDespesa({ id: edit.id, valor: Number(edit.valor) || 0, descricao: edit.descricao, data: edit.data });
      setEdit(null);
      await onChange();
    } catch (e) { avisar("Erro ao salvar: " + (e.message || "")); }
    finally { setBusy(false); }
  }
  async function excluir(id) {
    if (!(await confirmar("Excluir este pedágio? Não dá pra desfazer."))) return;
    setBusy(true);
    try { await apiDeleteDespesa(id); await onChange(); }
    catch (e) { avisar("Erro ao excluir: " + (e.message || "")); }
    finally { setBusy(false); }
  }

  // Rotas de km lançadas num dia (pode ter mais de um carro/viagem no mesmo dia).
  function rotasDoDia(data) {
    return (records || []).filter(r => r.data === data && !isOpen(r));
  }

  return (
    <Card className="mt-2.5 p-3.5">
      <div className="flex items-center justify-between mb-3">
        <h2 className="font-semibold text-gray-800">🛣️ Pedágios</h2>
        <button onClick={onAdd} className="text-xs font-medium px-2.5 py-1.5 rounded-lg" style={{ background: BTJ_BLUE, color: "#fff" }}>
          📄 Importar extrato do tag
        </button>
      </div>

      {periodos.length === 0 && <p className="text-sm text-gray-400 text-center py-6">Nenhum pedágio lançado ainda.</p>}

      <div className="space-y-2">
        {periodos.map(pk => {
          const itensPeriodo = byPeriod[pk];
          const totalPeriodo = itensPeriodo.reduce((s, d) => s + (Number(d.valor) || 0), 0);
          const abertoPeriodo = expandedPeriod === pk;

          // Agrupa os itens do período por dia
          const byDia = {};
          itensPeriodo.forEach(d => (byDia[d.data] = byDia[d.data] || []).push(d));
          const dias = Object.keys(byDia).sort().reverse();

          return (
            <div key={pk} className="border border-gray-100 rounded-xl overflow-hidden">
              <button onClick={() => setExpandedPeriod(abertoPeriodo ? null : pk)} className="w-full flex items-center justify-between px-3.5 py-3 text-left">
                <div>
                  <p className="text-sm font-semibold" style={{ color: BTJ_NAVY }}>{periodLabel(pk)}</p>
                  <p className="text-[11px] text-gray-400">{itensPeriodo.length} passagem(ns) · {dias.length} dia(s)</p>
                </div>
                <div className="flex items-center gap-2">
                  <span className="text-sm font-semibold" style={{ color: BTJ_BLUE }}>R$ {totalPeriodo.toLocaleString("pt-BR", { minimumFractionDigits: 2 })}</span>
                  <span className="text-gray-400 text-xs">{abertoPeriodo ? "▲" : "▼"}</span>
                </div>
              </button>

              {abertoPeriodo && (
                <div className="border-t border-gray-100">
                  {dias.map(dia => {
                    const itensDia = byDia[dia].sort((a, b) => (a.descricao || "").localeCompare(b.descricao || ""));
                    const totalDia = itensDia.reduce((s, d) => s + (Number(d.valor) || 0), 0);
                    const rotas = rotasDoDia(dia);
                    return (
                      <div key={dia} className="border-b border-gray-100 last:border-b-0 px-3.5 py-2.5">
                        <div className="flex items-center justify-between mb-1">
                          <p className="text-xs font-semibold text-gray-700">{formatDateShort(dia)} · {weekdayAbrev(dia)}</p>
                          <p className="text-xs font-semibold" style={{ color: "#854F0B" }}>R$ {totalDia.toLocaleString("pt-BR", { minimumFractionDigits: 2 })}</p>
                        </div>

                        {rotas.length > 0 ? (
                          rotas.map(r => (
                            <p key={r.id} className="text-[10px] text-gray-400 mb-1">
                              🚗 {(r.carro || CARRO_PADRAO).split(" ")[0]} · {r.origem || "?"}{r.destino ? ` → ${r.destino}` : ""} · {kmOf(r).toLocaleString("pt-BR")} km
                            </p>
                          ))
                        ) : (
                          <p className="text-[10px] text-gray-400 mb-1">sem viagem de km lançada nesse dia</p>
                        )}

                        <div className="space-y-1 mt-1.5">
                          {itensDia.map(d => (
                            edit?.id === d.id ? (
                              <div key={d.id} className="rounded-lg p-2.5" style={{ background: "#F8FAFC" }}>
                                <div className="flex gap-1.5 mb-2">
                                  <div className="flex-1">
                                    <p className="text-[9px] text-gray-500 mb-0.5">Data</p>
                                    <input type="date" value={edit.data} onChange={e => setEdit({ ...edit, data: e.target.value })}
                                      className="w-full border border-gray-200 rounded-lg px-2 py-1.5 text-xs" />
                                  </div>
                                  <div className="flex-1">
                                    <p className="text-[9px] text-gray-500 mb-0.5">Valor</p>
                                    <input type="number" inputMode="decimal" value={edit.valor} onChange={e => setEdit({ ...edit, valor: e.target.value })}
                                      onKeyDown={e => { if (e.key === "Enter") e.currentTarget.blur(); }}
                                      className="w-full border border-gray-200 rounded-lg px-2 py-1.5 text-xs" />
                                  </div>
                                </div>
                                <input type="text" value={edit.descricao} placeholder="praça/local" onChange={e => setEdit({ ...edit, descricao: e.target.value })}
                                  onKeyDown={e => { if (e.key === "Enter") e.currentTarget.blur(); }}
                                  className="w-full border border-gray-200 rounded-lg px-2 py-1.5 text-xs mb-2" />
                                <div className="flex gap-1.5">
                                  <button onClick={salvarEdicao} disabled={busy} className="flex-[2] rounded-lg py-2 text-xs font-medium text-white" style={{ background: BTJ_BLUE }}>Salvar</button>
                                  <button onClick={() => setEdit(null)} className="flex-1 rounded-lg py-2 text-xs text-gray-600 border border-gray-200">Cancelar</button>
                                  <button onClick={() => excluir(d.id)} disabled={busy} className="flex-1 rounded-lg py-2 text-xs" style={{ background: "#FDECEC", color: "#C62A2F" }}>🗑</button>
                                </div>
                              </div>
                            ) : (
                              <button key={d.id} onClick={() => { if (travado && travado(d.data)) { avisar("Período já enviado — edição travada."); return; } setEdit({ id: d.id, valor: d.valor, descricao: d.descricao, data: d.data }); }}
                                className="w-full flex items-center justify-between px-2.5 py-1.5 rounded-lg text-left" style={{ background: "#FAFAF8" }}>
                                <div className="min-w-0">
                                  <span className="text-xs text-gray-700">{d.descricao || "Pedágio"}</span>
                                  <span className="text-[10px] text-gray-400 ml-1.5">{d.origem === "extrato" ? "· importado" : "· manual"}</span>
                                </div>
                                <span className="text-xs font-semibold shrink-0" style={{ color: "#04342C" }}>R$ {(Number(d.valor) || 0).toLocaleString("pt-BR", { minimumFractionDigits: 2 })} ✎</span>
                              </button>
                            )
                          ))}
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
          );
        })}
      </div>
    </Card>
  );
}

// ─── Assinatura: desenhar na tela ou normalizar arquivo pra PNG 600×180 ─────
// A foto/scan vira só o TRAÇO (fundo do papel removido via transparência,
// como o Adobe faz) — no PDF a assinatura senta em cima da linha, sem o
// retângulo branco da foto por cima.
function arquivoParaAssinatura(file) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    const url = URL.createObjectURL(file);
    img.onload = () => {
      const cv = document.createElement("canvas");
      cv.width = 600; cv.height = 180;
      const ctx = cv.getContext("2d");
      const esc = Math.min(600 / img.width, 180 / img.height);
      const w = img.width * esc, h = img.height * esc;
      ctx.drawImage(img, (600 - w) / 2, (180 - h) / 2, w, h);
      // Papel some, traço fica: pixel claro -> transparente, meio-tom -> suave
      const id = ctx.getImageData(0, 0, 600, 180);
      const px = id.data;
      for (let i = 0; i < px.length; i += 4) {
        if (px[i + 3] === 0) continue; // já transparente (borda do enquadramento)
        const lum = 0.299 * px[i] + 0.587 * px[i + 1] + 0.114 * px[i + 2];
        if (lum > 200) px[i + 3] = 0;
        else if (lum > 140) px[i + 3] = Math.round(255 * (200 - lum) / 60);
      }
      ctx.putImageData(id, 0, 0);
      URL.revokeObjectURL(url);
      resolve(cv.toDataURL("image/png"));
    };
    img.onerror = () => { URL.revokeObjectURL(url); reject(new Error("Não consegui ler a imagem.")); };
    img.src = url;
  });
}

// Assinatura do solicitante fica salva no aparelho: envia uma vez, e nos
// próximos relatórios ela já vem carregada (dá pra trocar/remover na hora).
const KEY_ASSINATURA = "km_assinatura_v1";
function loadAssinatura() {
  try { return localStorage.getItem(KEY_ASSINATURA) || null; } catch { return null; }
}
function saveAssinatura(png) {
  try {
    if (png) localStorage.setItem(KEY_ASSINATURA, png);
    else localStorage.removeItem(KEY_ASSINATURA);
  } catch {}
}

function AssinaturaModal({ onConfirm, onCancel }) {
  const canvasRef = useRef(null);
  const desenhouRef = useRef(false);

  useEffect(() => {
    const cv = canvasRef.current;
    const ctx = cv.getContext("2d");
    // Pixels ficam TRANSPARENTES (o PNG exportado é só o traço); o fundo
    // branco que se vê é CSS do elemento, não faz parte da imagem.
    ctx.clearRect(0, 0, cv.width, cv.height);
    ctx.strokeStyle = "#0A2540"; ctx.lineWidth = 3; ctx.lineCap = "round"; ctx.lineJoin = "round";
    let drawing = false, last = null;
    const pos = (e) => {
      const r = cv.getBoundingClientRect();
      const t = e.touches ? e.touches[0] : e;
      return { x: (t.clientX - r.left) * (cv.width / r.width), y: (t.clientY - r.top) * (cv.height / r.height) };
    };
    const down = (e) => { e.preventDefault(); drawing = true; last = pos(e); };
    const move = (e) => {
      if (!drawing) return;
      e.preventDefault();
      const p = pos(e);
      ctx.beginPath(); ctx.moveTo(last.x, last.y); ctx.lineTo(p.x, p.y); ctx.stroke();
      last = p; desenhouRef.current = true;
    };
    const up = () => { drawing = false; };
    cv.addEventListener("mousedown", down);
    cv.addEventListener("mousemove", move);
    window.addEventListener("mouseup", up);
    cv.addEventListener("touchstart", down, { passive: false });
    cv.addEventListener("touchmove", move, { passive: false });
    cv.addEventListener("touchend", up);
    return () => {
      cv.removeEventListener("mousedown", down);
      cv.removeEventListener("mousemove", move);
      window.removeEventListener("mouseup", up);
      cv.removeEventListener("touchstart", down);
      cv.removeEventListener("touchmove", move);
      cv.removeEventListener("touchend", up);
    };
  }, []);

  function limpar() {
    const cv = canvasRef.current;
    cv.getContext("2d").clearRect(0, 0, cv.width, cv.height);
    desenhouRef.current = false;
  }
  function confirmar() {
    if (!desenhouRef.current) { avisar("Assine no quadro antes de confirmar."); return; }
    onConfirm(canvasRef.current.toDataURL("image/png"));
  }

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4" style={{ background: "rgba(0,0,0,0.45)" }}>
      <div className="bg-white rounded-xl p-4 w-full max-w-sm">
        <p className="font-semibold text-sm mb-2" style={{ color: BTJ_NAVY }}>✍️ Assine no quadro abaixo</p>
        <canvas ref={canvasRef} width={600} height={180}
          className="w-full rounded-lg mb-3" style={{ border: "1.5px dashed #C5C3BB", touchAction: "none", background: "#fff" }} />
        <div className="flex gap-2">
          <button onClick={confirmar} className="flex-[2] rounded-lg py-2 text-sm font-semibold text-white" style={{ background: BTJ_BLUE }}>Usar assinatura</button>
          <button onClick={limpar} className="flex-1 rounded-lg py-2 text-sm text-gray-600 border border-gray-200">Limpar</button>
          <button onClick={onCancel} className="flex-1 rounded-lg py-2 text-sm text-gray-600 border border-gray-200">Cancelar</button>
        </div>
      </div>
    </div>
  );
}

// ─── Tela: Aprovações (só pra quem é Aprovador no Cadastros, coluna I) ───────
function Aprovacoes({ usuario, onVoltar }) {
  const [lista, setLista] = useState(null); // null = carregando
  const [erro, setErro] = useState("");
  const [decidindo, setDecidindo] = useState(null); // id em decisão

  async function carregar() {
    setErro("");
    try {
      setLista(await apiListarReaberturas(usuario.nome));
    } catch (e) {
      setErro(e.message || "Erro ao carregar.");
      setLista([]);
    }
  }
  useEffect(() => { carregar(); }, []);

  async function decidir(id, decisao) {
    const verbo = decisao === "aprovar" ? "APROVAR" : "NEGAR";
    if (!(await confirmar(`${verbo} esta reabertura?`))) return;
    setDecidindo(id);
    try {
      await apiDecidirReabertura(id, decisao, usuario.nome);
      await carregar();
    } catch (e) {
      avisar(e.message || "Erro ao registrar a decisão.");
    } finally {
      setDecidindo(null);
    }
  }

  const pendentes = (lista || []).filter(r => r.status === "pendente");
  const decididas = (lista || []).filter(r => r.status !== "pendente");

  return (
    <Card className="p-4">
      <div className="flex items-center justify-between mb-3">
        <h2 className="font-bold text-base" style={{ color: BTJ_NAVY }}>✓ Aprovações</h2>
        <button onClick={onVoltar} className="text-xs text-gray-400">‹ Voltar</button>
      </div>

      {lista === null && <p className="text-center text-sm text-gray-400 py-6">Carregando…</p>}
      {erro && <p className="text-xs text-red-600 mb-2">{erro}</p>}

      {lista !== null && (
        <>
          <p className="text-[11px] font-semibold text-gray-500 mb-1.5">Reaberturas pendentes ({pendentes.length})</p>
          {pendentes.length === 0 && <p className="text-xs text-gray-400 mb-3">Nenhuma solicitação aguardando decisão.</p>}
          <div className="space-y-2 mb-4">
            {pendentes.map(r => (
              <div key={r.id} className="rounded-xl border p-3" style={{ borderColor: "#F2D9A0", background: "#FFFBEF" }}>
                <p className="text-sm font-semibold" style={{ color: BTJ_NAVY }}>
                  {monthLabelFromKey(r.periodo)} · {r.colaborador.split(" ").slice(0, 2).join(" ")}
                </p>
                <p className="text-xs text-gray-600 mt-0.5">“{r.motivo}”</p>
                <p className="text-[10px] text-gray-400 mt-0.5">solicitado em {r.solicitadoEm}</p>
                <div className="flex gap-1.5 mt-2">
                  <button onClick={() => decidir(r.id, "aprovar")} disabled={decidindo === r.id}
                    className="flex-1 rounded-lg py-1.5 text-xs font-semibold text-white"
                    style={{ background: decidindo === r.id ? "#9CA3AF" : "#0F766E" }}>
                    ✓ Aprovar
                  </button>
                  <button onClick={() => decidir(r.id, "negar")} disabled={decidindo === r.id}
                    className="flex-1 rounded-lg py-1.5 text-xs font-semibold text-white"
                    style={{ background: decidindo === r.id ? "#9CA3AF" : "#B3261E" }}>
                    ✗ Negar
                  </button>
                </div>
              </div>
            ))}
          </div>

          {decididas.length > 0 && (
            <>
              <p className="text-[11px] font-semibold text-gray-500 mb-1.5">Histórico</p>
              <div className="space-y-1.5">
                {decididas.map(r => (
                  <div key={r.id} className="rounded-lg border border-gray-100 px-3 py-2 bg-white">
                    <p className="text-xs text-gray-700">
                      {monthLabelFromKey(r.periodo)} · {r.colaborador.split(" ").slice(0, 2).join(" ")}{" "}
                      {r.status === "aprovado"
                        ? <span className="text-[10px] px-1.5 py-0.5 rounded-full" style={{ background: "#E1F5EE", color: "#085041" }}>aprovado</span>
                        : <span className="text-[10px] px-1.5 py-0.5 rounded-full" style={{ background: "#FDE8E8", color: "#9B1C1C" }}>negado</span>}
                    </p>
                    <p className="text-[10px] text-gray-400">“{r.motivo}” · decidido {r.decididoEm}{r.decididoPor ? ` · ${r.decididoPor}` : ""}</p>
                  </div>
                ))}
              </div>
            </>
          )}
        </>
      )}
    </Card>
  );
}

// ─── Tela: Revisão do Relatório do Mês (conferência antes de fechar/enviar) ──
function RevisaoRelatorio({ periodo, records, despesas, taxas, colaboradores, usuario, envio, reabertura, onVoltar, onEmitido, onPendente }) {
  const [assinatura, setAssinaturaState] = useState(envio?.assinaturaBase64 || loadAssinatura());
  // Toda mudança de assinatura também persiste no aparelho (null = esquecer)
  function setAssinatura(png) {
    setAssinaturaState(png);
    saveAssinatura(png);
  }
  const [mostrarAssinatura, setMostrarAssinatura] = useState(false);
  const [enviando, setEnviando] = useState(false);
  const [erro, setErro] = useState("");
  const arquivoRef = useRef(null);

  const jaEnviado = envio?.status === "enviado";
  const pendente = envio?.status === "pendente";
  const reaberto = envio?.status === "reaberto";
  const reabPendente = reabertura?.status === "pendente";
  const [pedindoReab, setPedindoReab] = useState(false);
  const [motivoReab, setMotivoReab] = useState("");
  const [reabEnviada, setReabEnviada] = useState(false);

  async function solicitarReabertura() {
    const motivo = motivoReab.trim();
    if (!motivo) { avisar("Explique o motivo da reabertura."); return; }
    try {
      await apiSolicitarReabertura(usuario.nome, periodo, motivo);
      setPedindoReab(false);
      setReabEnviada(true);
    } catch (e) {
      avisar(e.message || "Erro ao solicitar a reabertura.");
    }
  }
  const retorno = envio?.retorno || null;

  const recs = records.filter(r => periodKey(r.data) === periodo).sort((a, b) => a.data.localeCompare(b.data));
  const desps = despesas.filter(d => periodKey(d.data) === periodo).sort((a, b) => a.data.localeCompare(b.data));

  // ── Totais ──
  const kmTotal = recs.reduce((s, r) => s + kmOf(r), 0);
  const valorKm = recs.reduce((s, r) => s + kmOf(r) * taxaVigente(taxas, colaboradores, usuario.nome, r.data), 0);
  const pedagios = desps.filter(d => (d.tipo || "").toLowerCase().indexOf("ped") === 0);
  const outras = desps.filter(d => (d.tipo || "").toLowerCase().indexOf("ped") !== 0);
  const totalPedagio = pedagios.reduce((s, d) => s + (Number(d.valor) || 0), 0);
  const totalOutras = outras.reduce((s, d) => s + (Number(d.valor) || 0), 0);
  const totalGeral = valorKm + totalPedagio + totalOutras;

  // ── Caça às inconsistências ──
  const diasComViagem = new Set(recs.filter(r => !isOpen(r)).map(r => r.data));
  const problemas = [];
  recs.filter(isOpen).forEach(r => problemas.push({
    data: r.data,
    texto: `${formatDateShort(r.data)} (${weekdayAbrev(r.data)}): viagem sem km ${r.kmInicial == null ? "inicial" : "final"} (${r.origem || "?"} → ${r.destino || "?"})`,
  }));
  outras.filter(d => !d.comprovante && d.origem === "manual").forEach(d => problemas.push({
    data: d.data,
    texto: `${formatDateShort(d.data)} (${weekdayAbrev(d.data)}): ${d.tipo} de R$ ${(Number(d.valor) || 0).toLocaleString("pt-BR", { minimumFractionDigits: 2 })} sem comprovante`,
  }));
  const diasPedagioSemViagem = [...new Set(pedagios.filter(p => !diasComViagem.has(p.data)).map(p => p.data))];
  diasPedagioSemViagem.forEach(data => problemas.push({
    data,
    texto: `${formatDateShort(data)} (${weekdayAbrev(data)}): pedágio lançado, mas sem viagem de km nesse dia`,
  }));
  const avisos = outras.filter(d => d.justificativa).map(d => ({
    data: d.data,
    texto: `${formatDateShort(d.data)} (${weekdayAbrev(d.data)}): ${d.tipo} acima do limite — justificativa: "${d.justificativa}"`,
  }));

  const todasDatas = [...new Set([...recs.map(r => r.data), ...desps.map(d => d.data)])].sort();
  const fmt = v => v.toLocaleString("pt-BR", { minimumFractionDigits: 2 });

  async function enviar() {
    setErro("");
    if (!assinatura) {
      setErro("A assinatura do solicitante é obrigatória — assine na tela ou envie o arquivo antes de fechar.");
      return;
    }
    const acao = jaEnviado ? "reenviar" : "fechar e enviar";
    if (!(await confirmar(`Depois de ${acao}, os lançamentos deste período ficam travados pra edição. Continuar?`))) return;

    if (!navigator.onLine) {
      onPendente(assinatura);
      return;
    }
    setEnviando(true);
    try {
      const ret = await apiAprovarEEmitir(periodo, assinatura);
      onEmitido(ret, assinatura);
    } catch (e) {
      setErro(e.message || "Erro ao emitir. Tente de novo.");
    } finally {
      setEnviando(false);
    }
  }

  async function pegarArquivo(file) {
    if (!file) return;
    try {
      const png = await arquivoParaAssinatura(file);
      setAssinatura(png);
    } catch (e) {
      avisar(e.message || "Não consegui ler o arquivo.");
    }
  }

  return (
    <Card className="mt-2.5 p-3.5">
      <div className="flex items-center justify-between mb-1">
        <h2 className="font-semibold text-gray-800">📋 Revisão · {periodLabel(periodo)}</h2>
        <span className="text-[10px] font-medium px-2 py-0.5 rounded-full"
          style={jaEnviado ? { background: "#E1F5EE", color: "#085041" } : reaberto ? { background: "#FEF3E2", color: "#854F0B" } : pendente ? { background: "#FEF3E2", color: "#854F0B" } : { background: "#FEF3E2", color: "#854F0B" }}>
          {jaEnviado ? "✓ Enviado" : reaberto ? "🔓 Reaberto pra correção" : pendente ? "⟳ Envio pendente" : "Fechado · aguardando envio"}
        </span>
      </div>
      <p className="text-[11px] text-gray-400 mb-3">Confira tudo antes de fechar e enviar.</p>

      {/* ── Totais ── */}
      <div className="rounded-xl p-3 mb-3" style={{ background: BTJ_NAVY }}>
        <div className="flex justify-between items-baseline mb-1.5">
          <span className="text-[11px]" style={{ color: BTJ_LIGHT }}>Km rodado ({kmTotal.toLocaleString("pt-BR")} km)</span>
          <span className="text-sm text-white font-medium">R$ {fmt(valorKm)}</span>
        </div>
        <div className="flex justify-between items-baseline mb-1.5">
          <span className="text-[11px]" style={{ color: BTJ_LIGHT }}>Pedágios ({pedagios.length})</span>
          <span className="text-sm text-white font-medium">R$ {fmt(totalPedagio)}</span>
        </div>
        <div className="flex justify-between items-baseline mb-2">
          <span className="text-[11px]" style={{ color: BTJ_LIGHT }}>Outras despesas ({outras.length})</span>
          <span className="text-sm text-white font-medium">R$ {fmt(totalOutras)}</span>
        </div>
        <div className="flex justify-between items-baseline pt-2" style={{ borderTop: "1px solid rgba(255,255,255,0.15)" }}>
          <span className="text-xs font-semibold text-white">TOTAL A REEMBOLSAR</span>
          <span className="text-lg font-bold" style={{ color: BTJ_BLUE }}>R$ {fmt(totalGeral)}</span>
        </div>
      </div>

      {/* ── Inconsistências ── */}
      {problemas.length > 0 && (
        <div className="rounded-xl p-3 mb-3" style={{ background: "#FDECEC", border: "1px solid #F5B8B8" }}>
          <p className="text-xs font-semibold mb-1.5" style={{ color: "#C62A2F" }}>⚠ {problemas.length} pendência(s) — corrija antes de enviar</p>
          {problemas.map((p, i) => (
            <p key={i} className="text-[11px] mb-1" style={{ color: "#8A1F23" }}>• {p.texto}</p>
          ))}
        </div>
      )}
      {avisos.length > 0 && (
        <div className="rounded-xl p-3 mb-3" style={{ background: "#FEF3E2", border: "1px solid #F5C97A" }}>
          <p className="text-xs font-semibold mb-1.5" style={{ color: "#854F0B" }}>ℹ Vai junto no relatório</p>
          {avisos.map((a, i) => (
            <p key={i} className="text-[11px] mb-1" style={{ color: "#6B4409" }}>• {a.texto}</p>
          ))}
        </div>
      )}

      {/* ── Linha do tempo por dia ── */}
      <div className="space-y-1.5 mb-4">
        {todasDatas.map(data => {
          const viagens = recs.filter(r => r.data === data);
          const despsDia = desps.filter(d => d.data === data);
          const temProblema = problemas.some(p => p.data === data);
          return (
            <div key={data} className="rounded-lg px-3 py-2" style={{ background: temProblema ? "#FFF7F7" : "#FAFAF8", border: temProblema ? "1px solid #F5B8B8" : "1px solid transparent" }}>
              <p className="text-[11px] font-semibold text-gray-700 mb-0.5">{formatDateShort(data)} · {weekdayAbrev(data)}{temProblema ? " ⚠" : ""}</p>
              {viagens.map(r => (
                <div key={r.id} className="flex justify-between items-baseline">
                  <span className="text-[11px] text-gray-600">🚗 {(r.carro || "").split(" ")[0]} · {r.origem || "?"} → {r.destino || "?"} · {isOpen(r) ? "— km" : `${kmOf(r).toLocaleString("pt-BR")} km`}</span>
                  <span className="text-[11px] font-medium text-gray-700">{isOpen(r) ? "—" : `R$ ${fmt(kmOf(r) * taxaVigente(taxas, colaboradores, usuario.nome, r.data))}`}</span>
                </div>
              ))}
              {despsDia.map(d => (
                <div key={d.id} className="flex justify-between items-baseline">
                  <span className="text-[11px] text-gray-500">
                    {(d.tipo || "").toLowerCase().indexOf("ped") === 0 ? "🛣️" : "💳"} {d.tipo}{d.descricao ? ` · ${d.descricao}` : ""}
                    {!d.comprovante && d.origem === "manual" && (d.tipo || "").toLowerCase().indexOf("ped") !== 0 ? " ⚠" : ""}
                  </span>
                  <span className="text-[11px] font-medium text-gray-700">R$ {fmt(Number(d.valor) || 0)}</span>
                </div>
              ))}
            </div>
          );
        })}
      </div>

      {/* ── Assinatura ── */}
      <p className="text-[11px] text-gray-500 mb-1">Assinatura do solicitante <span className="font-semibold" style={{ color: "#B3261E" }}>(obrigatória)</span></p>
      <div className="flex gap-1.5 mb-2">
        <button onClick={() => setMostrarAssinatura(true)} className="flex-1 rounded-lg py-2 text-xs font-medium border" style={{ borderColor: BTJ_BLUE, color: BTJ_NAVY }}>✍️ Assinar na tela</button>
        <button onClick={() => arquivoRef.current?.click()} className="flex-1 rounded-lg py-2 text-xs font-medium border" style={{ borderColor: BTJ_BLUE, color: BTJ_NAVY }}>🖼️ Enviar arquivo</button>
      </div>
      <input ref={arquivoRef} type="file" accept="image/*" className="hidden" onChange={e => { pegarArquivo(e.target.files[0]); e.target.value = ""; }} />
      {assinatura && (
        <div className="flex items-center gap-2 mb-2">
          <img src={assinatura} alt="assinatura" className="h-10 rounded border border-gray-200 bg-white" />
          <div>
            <p className="text-[10px] text-gray-400">✓ salva neste aparelho — vale pros próximos relatórios</p>
            <button onClick={() => setAssinatura(null)} className="text-[10px] text-gray-400 underline">remover</button>
          </div>
        </div>
      )}
      {mostrarAssinatura && (
        <AssinaturaModal
          onConfirm={(png) => { setAssinatura(png); setMostrarAssinatura(false); }}
          onCancel={() => setMostrarAssinatura(false)}
        />
      )}

      {/* ── Retorno do envio ── */}
      {jaEnviado && retorno && (
        <div className="rounded-xl p-3 mb-2" style={{ background: "#E1F5EE", border: "1px solid #9BD8C3" }}>
          <p className="text-xs font-semibold mb-0.5" style={{ color: "#085041" }}>✓ Relatório enviado para {retorno.enviadoPara}</p>
          <p className="text-[11px]" style={{ color: "#0F6E56" }}>em {retorno.enviadoEm}{retorno.comAssinatura ? " · com assinatura" : ""}</p>
          {retorno.reemissao && <p className="text-[10px] mt-1" style={{ color: "#6B4409" }}>⚠ Este período já havia sido emitido — os arquivos anteriores foram substituídos.</p>}
        </div>
      )}
      {pendente && (
        <div className="rounded-xl p-3 mb-2" style={{ background: "#FEF3E2", border: "1px solid #F5C97A" }}>
          <p className="text-xs font-semibold" style={{ color: "#854F0B" }}>⟳ Fechado — o envio será concluído assim que houver conexão.</p>
        </div>
      )}
      {erro && (
        <div className="rounded-xl p-3 mb-2" style={{ background: "#FDECEC", border: "1px solid #F5B8B8" }}>
          <p className="text-xs" style={{ color: "#C62A2F" }}>⚠ {erro}</p>
        </div>
      )}

      {/* ── Reabertura ── */}
      {reaberto && (
        <div className="rounded-xl px-3 py-2.5 mb-2" style={{ background: "#FFF7E6", border: "1px solid #F2D9A0" }}>
          <p className="text-xs font-semibold" style={{ color: "#854F0B" }}>🔓 Reabertura aprovada — período liberado pra correção</p>
          <p className="text-[11px] mt-0.5" style={{ color: "#854F0B" }}>Edite os lançamentos de km e despesas e depois feche e reenvie. Os arquivos anteriores serão substituídos.</p>
        </div>
      )}
      {jaEnviado && (reabPendente || reabEnviada) && (
        <div className="rounded-xl px-3 py-2.5 mb-2" style={{ background: "#FFF7E6", border: "1px solid #F2D9A0" }}>
          <p className="text-xs font-semibold" style={{ color: "#854F0B" }}>⏳ Reabertura solicitada — aguardando aprovação</p>
          <p className="text-[11px] mt-0.5" style={{ color: "#854F0B" }}>O aprovador foi avisado por e-mail. Quando aprovar, o período destrava aqui no app pra você corrigir e reenviar.</p>
        </div>
      )}

      {/* ── Fechar e enviar / Reenviar ── */}
      <button onClick={enviar} disabled={enviando || (problemas.length > 0 && !jaEnviado)}
        className="w-full rounded-xl py-3 text-sm font-semibold text-white disabled:opacity-50 mb-1.5"
        style={{ background: jaEnviado ? "#0F6E56" : BTJ_BLUE }}>
        {enviando ? "⟳ Gerando e enviando (15-20s)..." : jaEnviado ? "↻ Reenviar relatório" : pendente ? "⟳ Tentar enviar agora" : reaberto ? "🔒 Fechar e reenviar relatório" : "🔒 Fechar e enviar relatório"}
      </button>
      {problemas.length > 0 && !jaEnviado && <p className="text-[10px] text-center" style={{ color: "#C62A2F" }}>Resolva as pendências acima pra liberar o envio</p>}
      {jaEnviado && retorno && (
        <div className="flex gap-1.5 mt-2">
          <button onClick={() => window.open(retorno.pdfUrl, "_blank")} className="flex-1 rounded-lg py-2 text-xs font-medium border border-gray-200 text-gray-600">📄 Ver PDF</button>
          <button onClick={() => window.open(retorno.excelUrl, "_blank")} className="flex-1 rounded-lg py-2 text-xs font-medium border border-gray-200 text-gray-600">📊 Ver Excel</button>
        </div>
      )}
      {jaEnviado && !reabPendente && !reabEnviada && (
        <button onClick={() => setPedindoReab(true)}
          className="w-full rounded-lg py-2 text-xs font-medium mt-1.5 border" style={{ borderColor: "#F2D9A0", color: "#854F0B", background: "#FFFBEF" }}>
          🔓 Solicitar reabertura pra corrigir
        </button>
      )}
      <button onClick={onVoltar} className="w-full rounded-lg py-2 text-xs text-gray-500 mt-1">‹ Voltar</button>

      {/* Modal do motivo da reabertura */}
      {pedindoReab && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4" style={{ background: "rgba(0,0,0,0.45)" }}>
          <div className="bg-white rounded-xl p-4 w-full max-w-sm">
            <p className="font-semibold text-sm mb-1" style={{ color: BTJ_NAVY }}>🔓 Solicitar reabertura</p>
            <p className="text-[11px] text-gray-500 mb-2">O período continua travado até o aprovador autorizar. Explique o que precisa ser corrigido:</p>
            <textarea value={motivoReab} onChange={e => setMotivoReab(e.target.value)} rows={3}
              placeholder="Ex: faltou lançar a viagem do dia 14 e o pedágio saiu duplicado"
              className="w-full border border-gray-200 rounded-lg px-2 py-1.5 text-sm mb-2" />
            <div className="flex gap-1.5">
              <button onClick={solicitarReabertura} className="flex-[2] rounded-lg py-2 text-sm font-semibold text-white" style={{ background: BTJ_BLUE }}>Enviar solicitação</button>
              <button onClick={() => setPedindoReab(false)} className="flex-1 rounded-lg py-2 text-xs text-gray-600 border border-gray-200">Cancelar</button>
            </div>
          </div>
        </div>
      )}
    </Card>
  );
}

// ─── Tela de Login ─────────────────────────────────────────────────────────
function LoginScreen({ onLogin }) {
  const [email, setEmail] = useState("");
  const [senha, setSenha] = useState("");
  const [erro, setErro] = useState("");
  const [entrando, setEntrando] = useState(false);

  async function entrar() {
    setErro("");
    if (!email.trim()) { setErro("Informe o e-mail corporativo."); return; }
    if (senha.trim().length !== 4) { setErro("A senha são os 4 primeiros números do seu CPF."); return; }
    setEntrando(true);
    try {
      const colaborador = await apiLogin(email.trim(), senha.trim());
      saveSessao(colaborador);
      onLogin(colaborador);
    } catch (e) {
      setErro(e.message || "Não foi possível entrar.");
    } finally {
      setEntrando(false);
    }
  }

  return (
    <div className="min-h-screen flex items-center justify-center px-5" style={{ background: BTJ_NAVY }}>
      <div className="w-full max-w-xs">
        <div className="flex justify-center mb-6">
          <div className="bg-white px-3 py-2 rounded-md">
            <img src={logoUrl} alt="BTJ" className="h-9" onError={e => { e.target.outerHTML = '<span style="color:#001F3E;font-weight:700;letter-spacing:1px;font-size:20px;">BTJ</span>'; }} />
          </div>
        </div>
        <p className="text-center text-white text-sm mb-6" style={{ color: BTJ_LIGHT }}>Km_BTJ · entre com seu e-mail corporativo</p>

        <div className="bg-white rounded-xl p-4">
          <p className="text-[11px] text-gray-500 mb-0.5">E-mail corporativo</p>
          <input type="email" value={email} onChange={e => setEmail(e.target.value)}
            onKeyDown={e => { if (e.key === "Enter") entrar(); }}
            placeholder="voce@empresa.com.br" autoCapitalize="none" autoCorrect="off"
            className="w-full border border-gray-200 rounded-lg px-2.5 py-2 text-sm mb-3" />

          <p className="text-[11px] text-gray-500 mb-0.5">Senha (4 primeiros números do seu CPF)</p>
          <input type="password" inputMode="numeric" maxLength={4} value={senha} onChange={e => setSenha(e.target.value.replace(/\D/g, ""))}
            onKeyDown={e => { if (e.key === "Enter") entrar(); }}
            placeholder="••••"
            className="w-full border border-gray-200 rounded-lg px-2.5 py-2 text-sm mb-3 tracking-widest" />

          {erro && <p className="text-[12px] mb-3" style={{ color: "#C62A2F" }}>⚠ {erro}</p>}

          <button onClick={entrar} disabled={entrando}
            className="w-full rounded-lg py-2.5 text-sm font-semibold text-white disabled:opacity-60"
            style={{ background: BTJ_BLUE }}>
            {entrando ? "Entrando..." : "Entrar"}
          </button>
        </div>
        <p className="text-center text-[11px] mt-4" style={{ color: BTJ_LIGHT }}>
          Depois de entrar, o app fica logado neste aparelho.
        </p>
      </div>
    </div>
  );
}

// ─── App ──────────────────────────────────────────────────────────────────────
export default function App() {
  const [usuario, setUsuario] = useState(loadSessao());
  const [records, setRecords] = useState([]);
  const [config, setConfig] = useState(loadCachedConfig());
  const [screen, setScreen] = useState("home"); // home | resumos | despesa | extrato | despMenu | despPedagio | despOutras
  const [despesas, setDespesas] = useState([]);
  const [online, setOnline] = useState(navigator.onLine);
  const [syncStatus, setSyncStatus] = useState(null);
  const [ultimaSync, setUltimaSync] = useState(null); // Date da última sincronização bem-sucedida
  const [, forceTick] = useState(0); // re-render periódico pra atualizar o "há X min"
  const [needRefresh, setNeedRefresh] = useState(false);
  const updateSWRef = useRef(null);
  const [showPending, setShowPending] = useState(false);
  const [expandedMonth, setExpandedMonth] = useState(periodKey(todayISO()));
  const [menuAberto, setMenuAberto] = useState(false);
  const [pedagioAberto, setPedagioAberto] = useState(null); // data (iso) do dia com pedágios expandidos
  const [inlineEdit, setInlineEdit] = useState(null); // { id, kmIni, kmFin, obs, err }

  // Formulário (apontamento do dia)
  const [editingId, setEditingId] = useState(null);
  const [fData, setFData] = useState(todayISO());
  const [fCarro, setFCarro] = useState(loadLastCar());
  const [mostrarCadastroCarroPrincipal, setMostrarCadastroCarroPrincipal] = useState(false);
  const [envios, setEnvios] = useState(loadEnvios());
  const [revisaoPeriodo, setRevisaoPeriodo] = useState(null);
  // Agrupamento do relatório de km: mês (ciclo 26→25, padrão) ou semana (dom–sáb).
  // PRECISA ficar aqui em cima, antes do gate de login — hook depois de um
  // return antecipado quebra o React (tela branca ao logar; erro #310).
  const [agrupamento, setAgrupamento] = useState("mes");

  const hojeKey = periodKey(todayISO());
  const perAnterior = periodoAnterior(hojeKey);

  function setEnvioPeriodo(periodo, dados) {
    setEnvios(prev => {
      const novo = { ...prev, [periodo]: dados };
      saveEnvios(novo);
      return novo;
    });
  }

  // Período travado = já fechado e enviado (ou fechado aguardando envio offline)
  function periodoTravado(dataISO) {
    const st = statusPeriodo(periodKey(dataISO), envios, hojeKey);
    return st === "enviado" || st === "pendente";
  }

  // Re-tenta envios pendentes quando a conexão volta
  useEffect(() => {
    if (!online) return;
    Object.entries(envios).forEach(async ([periodo, e]) => {
      if (e.status !== "pendente") return;
      try {
        const ret = await apiAprovarEEmitir(periodo, e.assinaturaBase64);
        setEnvioPeriodo(periodo, { status: "enviado", retorno: ret, assinaturaBase64: e.assinaturaBase64 });
      } catch { /* tenta de novo na próxima vez que ficar online */ }
    });
  }, [online]);

  // Atualiza o "App sincronizado há X min" a cada 30s (só quando está nesse estado)
  useEffect(() => {
    if (syncStatus !== "synced") return;
    const t = setInterval(() => forceTick(n => n + 1), 30000);
    return () => clearInterval(t);
  }, [syncStatus]);

  // ── Reaberturas: sincroniza minhas solicitações e destrava períodos aprovados ──
  const [minhasReab, setMinhasReab] = useState([]);
  useEffect(() => {
    if (!usuario || !online) return;
    apiStatusReaberturas(usuario.nome).then(reab => {
      setMinhasReab(reab);
      setEnvios(prev => {
        const nxt = { ...prev };
        let mudou = false;
        reab.forEach(r => {
          if (r.status !== "aprovado") return;
          const e = nxt[r.periodo];
          // Destrava só uma vez por solicitação: depois do reenvio, a mesma
          // aprovação antiga não reabre de novo (reaberturaId marca isso).
          if (e && e.status === "enviado" && e.reaberturaId !== r.id) {
            nxt[r.periodo] = { ...e, status: "reaberto", reaberturaId: r.id };
            mudou = true;
          }
        });
        if (mudou) saveEnvios(nxt);
        return mudou ? nxt : prev;
      });
    }).catch(() => {});
  }, [usuario, online, screen]);

  // Sou aprovador? (vem do Cadastros, coluna I — sempre fresco via config)
  const ehAprovador = !!(config.colaboradores || []).find(
    c => c.nome?.toLowerCase() === usuario?.nome?.toLowerCase() && c.aprovador
  );

  // Compartilhado por todas as telas com seletor de carro: adiciona o carro
  // recém-cadastrado na config local, sem precisar recarregar tudo.
  function handleNovoCarro(novoCarro) {
    setConfig(c => ({ ...c, carros: [...(c.carros || []), novoCarro] }));
  }
  const [fOrigem, setFOrigem] = useState("");
  const [fOrigemGps, setFOrigemGps] = useState(false);
  const [fDestino, setFDestino] = useState("");
  const [fKmIni, setFKmIni] = useState(null);
  const [fKmFin, setFKmFin] = useState(null);
  const [fObs, setFObs] = useState("");
  const [loadingIni, setLoadingIni] = useState(false);
  const [loadingFin, setLoadingFin] = useState(false);
  const [queuedIni, setQueuedIni] = useState(false);
  const [queuedFin, setQueuedFin] = useState(false);

  const fileIniCamRef = useRef(null);
  const fileIniGalRef = useRef(null);
  const fileFinCamRef = useRef(null);
  const fileFinGalRef = useRef(null);
  const [coherErr, setCoherErr] = useState(null); // { field, msg } | null

  // ── Init ──
  useEffect(() => {
    updateSWRef.current = registerSW({
      onNeedRefresh() { setNeedRefresh(true); },
    });
    const on = () => setOnline(true);
    const off = () => setOnline(false);
    window.addEventListener("online", on);
    window.addEventListener("offline", off);
    return () => { window.removeEventListener("online", on); window.removeEventListener("offline", off); };
  }, []);

  useEffect(() => {
    if (!usuario) return;
    setRecords(loadRecords());
    apiFetchConfig(usuario.nome)
      .then(c => { setConfig(c); localStorage.setItem(KEY_CONFIG, JSON.stringify(c)); })
      .catch(() => {});
    pullAndReconcile(); // baixa da planilha (fonte mestra) e alinha o local
    refreshDespesas();
  }, [usuario]);

  // Baixa os lançamentos da planilha e reconcilia com o local.
  // Regra: pendentes de envio (synced=false) têm prioridade e são preservados;
  // para todo o resto, a planilha é a fonte da verdade.
  async function refreshDespesas() {
    if (!navigator.onLine || !usuario) return;
    try { setDespesas(await apiListDespesas(usuario.nome)); } catch {}
  }

  async function pullAndReconcile() {
    if (!navigator.onLine || !usuario) return;
    try {
      setSyncStatus("syncing");
      const remote = (await apiList(usuario.nome)).filter(registroValido_);
      const local = loadRecords();
      // Rede de segurança: guarda uma cópia do estado local ANTES de qualquer
      // reconciliação. Se algo der errado, nada se perde de verdade.
      try { localStorage.setItem("km_registros_backup", JSON.stringify({ ts: Date.now(), records: local })); } catch {}
      // Proteção: se a base remota vier vazia mas houver dados locais,
      // NÃO sobrescreve — mantém o local e avisa.
      if (remote.length === 0 && local.length > 0) {
        setSyncStatus("error");
        return;
      }
      const pendentes = local.filter(r => r.synced === false);
      const keyOf = r => `${r.data}|${r.carro || CARRO_PADRAO}`;
      const pendByKey = new Map(pendentes.map(r => [keyOf(r), r]));

      const merged = [];
      for (const rem of remote) {
        const k = `${rem.data}|${rem.carro || CARRO_PADRAO}`;
        if (pendByKey.has(k)) {
          merged.push(pendByKey.get(k));
          pendByKey.delete(k);
        } else {
          merged.push({
            id: `sheet-${k}`,
            data: rem.data,
            carro: rem.carro || CARRO_PADRAO,
            origem: rem.origem,
            destino: rem.destino,
            kmInicial: rem.kmInicial,
            kmFinal: rem.kmFinal,
            observacao: rem.observacao,
            synced: true,
          });
        }
      }
      for (const p of pendByKey.values()) merged.push(p);

      merged.sort((a, b) => a.data.localeCompare(b.data));
      setRecords(merged);
      persistRecords(merged);
      // Sobe as pendências que sobreviveram (elas sobrescrevem a planilha)
      resyncUnsynced();
      setUltimaSync(new Date());
      setSyncStatus("synced"); // carga/sincronização — NÃO é uma gravação de apontamento
    } catch {
      setSyncStatus(null); // offline ou erro: mantém o local como está
    }
  }

  // ── Carrega o registro do dia+carro selecionado no formulário ──
  useEffect(() => {
    const rec = records.find(r => r.data === fData && (r.carro || CARRO_PADRAO) === fCarro);
    if (rec && rec.id !== editingId) {
      setEditingId(rec.id);
      setFOrigem(rec.origem || "");
      setFDestino(rec.destino || "");
      setFKmIni(rec.kmInicial ?? null);
      setFKmFin(rec.kmFinal ?? null);
      setFObs(rec.observacao || "");
      setFOrigemGps(false);
    } else if (!rec && editingId !== null) {
      setEditingId(null);
      setFOrigem(""); setFDestino(""); setFKmIni(null); setFKmFin(null); setFObs("");
      setFOrigemGps(false);
    }
    const q = loadPhotoQueue();
    setQueuedIni(q.some(p => p.date === fData && p.carro === fCarro && p.phase === "inicial"));
    setQueuedFin(q.some(p => p.date === fData && p.carro === fCarro && p.phase === "final"));
  }, [fData, fCarro, records]);

  // ── Coerência em tempo real (por carro) ──
  useEffect(() => {
    setCoherErr(checkCoherence(records, fData, fKmIni, fKmFin, editingId, fCarro));
  }, [fKmIni, fKmFin, fData, fCarro, records, editingId]);

  // ── GPS: preenche origem ao abrir (se vazia) ──
  useEffect(() => {
    if (fOrigem) return;
    let alive = true;
    gpsCidade().then(c => { if (alive && c) { setFOrigem(c); setFOrigemGps(true); } });
    return () => { alive = false; };
  }, [fData]);

  // ── Fila offline: processa quando volta a conexão ──
  useEffect(() => {
    if (!online) return;
    processPhotoQueue();
    resyncUnsynced();
  }, [online]);

  // Depois daqui não tem mais nenhum hook — pode sair sem violar a regra dos hooks.
  if (!usuario) {
    return <><DialogHost /><LoginScreen onLogin={setUsuario} /></>;
  }

  async function processPhotoQueue() {
    let q = loadPhotoQueue();
    if (!q.length) return;
    setSyncStatus("saving");
    for (const item of [...q]) {
      try {
        const km = await apiOcr(item.b64);
        applyKmToDate(item.date, item.carro || CARRO_PADRAO, item.phase, km);
        q = q.filter(p => !(p.date === item.date && (p.carro || CARRO_PADRAO) === (item.carro || CARRO_PADRAO) && p.phase === item.phase));
        persistPhotoQueue(q);
      } catch { /* tenta de novo na próxima conexão */ }
    }
    setQueuedIni(q.some(p => p.date === fData && (p.carro || CARRO_PADRAO) === fCarro && p.phase === "inicial"));
    setQueuedFin(q.some(p => p.date === fData && (p.carro || CARRO_PADRAO) === fCarro && p.phase === "final"));
    setUltimaSync(new Date());
    setSyncStatus("ok");
  }

  async function resyncUnsynced() {
    const pending = loadRecords().filter(r => r.synced === false);
    for (const r of pending) {
      try {
        await apiSave(r, usuario.nome);
        mutateRecords(recs => recs.map(x => x.id === r.id ? { ...x, synced: true } : x));
      } catch { /* fica pra próxima */ }
    }
  }

  function mutateRecords(fn) {
    setRecords(prev => {
      const next = fn(prev);
      persistRecords(next);
      return next;
    });
  }

  function applyKmToDate(dateISO, carro, phase, km) {
    if (km == null) return;
    let updatedRec = null;
    mutateRecords(recs => {
      const idx = recs.findIndex(r => r.data === dateISO && (r.carro || CARRO_PADRAO) === carro);
      if (idx === -1) return recs;
      const upd = { ...recs[idx], [phase === "inicial" ? "kmInicial" : "kmFinal"]: km, synced: false };
      updatedRec = upd;
      const next = [...recs];
      next[idx] = upd;
      return next;
    });
    if (dateISO === fData && carro === fCarro) {
      if (phase === "inicial") setFKmIni(km); else setFKmFin(km);
    }
    syncRecord(updatedRec);
  }

  async function syncRecord(rec) {
    if (!rec) return;
    try {
      setSyncStatus("saving");
      await apiSave(rec, usuario.nome);
      mutateRecords(recs => recs.map(x => x.id === rec.id ? { ...x, synced: true } : x));
      setUltimaSync(new Date());
      setSyncStatus("ok");
    } catch {
      setSyncStatus("error");
    }
  }

  // ── Foto (inicial/final) ──
  async function handlePhoto(phase, file) {
    if (!file) return;
    const setLoading = phase === "inicial" ? setLoadingIni : setLoadingFin;
    setLoading(true);
    try {
      const b64 = await fileToResizedBase64(file);
      if (!navigator.onLine) {
        const q = loadPhotoQueue().filter(p => !(p.date === fData && (p.carro || CARRO_PADRAO) === fCarro && p.phase === phase));
        q.push({ date: fData, carro: fCarro, phase, b64, ts: Date.now() });
        persistPhotoQueue(q);
        if (phase === "inicial") setQueuedIni(true); else setQueuedFin(true);
        avisar("Sem internet agora. A foto ficou guardada e será lida automaticamente quando a conexão voltar.");
        return;
      }
      const km = await apiOcr(b64);
      if (km == null) {
        avisar("Não consegui ler o odômetro. Digite o KM manualmente.");
        return;
      }
      if (phase === "inicial") {
        setFKmIni(km);
        offerBridge(km);
      } else {
        setFKmFin(km);
      }
    } catch (e) {
      avisar("Erro ao processar a imagem: " + (e.message || "tente novamente."));
    } finally {
      setLoading(false);
    }
  }

  // ── KM emendado: foto inicial de hoje pode fechar o dia aberto anterior (MESMO carro) ──
  async function offerBridge(km) {
    const open = records
      .filter(r => r.data < fData && (r.carro || CARRO_PADRAO) === fCarro && r.kmInicial != null && r.kmFinal == null)
      .sort((a, b) => b.data.localeCompare(a.data))[0];
    if (!open) return;
    const err = checkCoherence(records, open.data, open.kmInicial, km, open.id, fCarro);
    if (err) return;
    const ok = await confirmar(
      `O dia ${formatDateBR(open.data)} está sem KM final (${fCarro}).\n` +
      `Como o carro ficou parado, usar esta leitura (${km.toLocaleString("pt-BR")}) como o KM final daquele dia?`
    );
    if (ok) applyKmToDate(open.data, fCarro, "final", km);
  }

  // ── Reaproveitar KM final do último dia (MESMO carro) como inicial de hoje ──
  const prevClosed = records
    .filter(r => r.data < fData && (r.carro || CARRO_PADRAO) === fCarro && r.kmFinal != null)
    .sort((a, b) => b.data.localeCompare(a.data))[0];
  const showReuse = prevClosed && fKmIni == null;

  // ── Salvar apontamento (parcial permitido, mas coerente) ──
  function save() {
    if (fKmIni == null && fKmFin == null && !fObs && !fDestino) {
      avisar("Preencha ao menos um KM, destino ou observação.");
      return;
    }
    const err = checkCoherence(records, fData, fKmIni, fKmFin, editingId, fCarro);
    if (err) {
      setCoherErr(err);
      avisar("⛔ " + err.msg);
      return;
    }
    if (periodoTravado(fData)) {
      avisar("Este período já foi fechado e enviado — não dá mais pra lançar ou editar km nele.");
      return;
    }
    saveLastCar(fCarro);
    const existing = records.find(r => r.data === fData && (r.carro || CARRO_PADRAO) === fCarro);
    const rec = {
      id: existing?.id ?? editingId ?? Date.now(),
      data: fData,
      carro: fCarro,
      origem: fOrigem,
      destino: fDestino,
      kmInicial: fKmIni,
      kmFinal: fKmFin,
      observacao: fObs,
      synced: false,
    };
    mutateRecords(recs => {
      const exists = recs.some(r => r.data === rec.data && (r.carro || CARRO_PADRAO) === fCarro);
      return exists ? recs.map(r => (r.data === rec.data && (r.carro || CARRO_PADRAO) === fCarro) ? rec : r) : [...recs, rec];
    });
    setEditingId(rec.id);
    syncRecord(rec);
  }

  // ── Derivados ──
  const openDays = records.filter(isOpen).sort((a, b) => b.data.localeCompare(a.data));
  const curKey = periodKey(todayISO());
  const cur = monthSummary(records, curKey, config.taxas, config.colaboradores, usuario.nome, despesas);
  const todayRec = records.find(r => r.data === todayISO());
  const kmHoje = todayRec ? kmOf(todayRec) : 0;

  const monthKeys = [...new Set(records.map(r => periodKey(r.data)))].sort().reverse().slice(0, 3);
  const weekKeys = [...new Set(records.map(r => weekKey(r.data)))].sort().reverse();
  const groupKeys = agrupamento === "mes" ? monthKeys : weekKeys;
  const curGroupKey = agrupamento === "mes" ? periodKey(todayISO()) : weekKey(todayISO());
  const destinos = config.destinos || DEFAULT_CONFIG.destinos;



  // ── Edição inline na tela de resumos ──
  function openInline(r) {
    if (periodoTravado(r.data)) {
      avisar("Este período já foi fechado e enviado — os lançamentos estão travados. Use 'Reenviar' na revisão se precisar corrigir algo com o financeiro.");
      return;
    }
    setInlineEdit({ id: r.id, kmIni: r.kmInicial, kmFin: r.kmFinal, obs: r.observacao || "", err: null });
  }
  function changeInline(patch) {
    setInlineEdit(prev => {
      const next = { ...prev, ...patch };
      const rec = records.find(r => r.id === next.id);
      next.err = rec ? checkCoherence(records, rec.data, next.kmIni, next.kmFin, next.id, rec.carro || CARRO_PADRAO) : null;
      return next;
    });
  }
  function saveInline() {
    const rec = records.find(r => r.id === inlineEdit.id);
    if (!rec) return;
    const err = checkCoherence(records, rec.data, inlineEdit.kmIni, inlineEdit.kmFin, inlineEdit.id, rec.carro || CARRO_PADRAO);
    if (err) { setInlineEdit(prev => ({ ...prev, err })); avisar("⛔ " + err.msg); return; }
    const updated = { ...rec, kmInicial: inlineEdit.kmIni, kmFinal: inlineEdit.kmFin, observacao: inlineEdit.obs, synced: false };
    mutateRecords(recs => recs.map(x => x.id === rec.id ? updated : x));
    setInlineEdit(null);
    syncRecord(updated);
  }

  // Exclui o lançamento de km do dia: apaga do aparelho E da planilha (a linha
  // do Log é limpa, não deletada). Exige internet — sem fila de exclusão, pra
  // evitar o zumbi que volta na sincronização.
  async function excluirDia(rec) {
    if (periodoTravado(rec.data)) { avisar("Este período já foi fechado e enviado — não dá pra excluir lançamentos dele."); return; }
    const rotulo = `${formatDateBR(rec.data)} · ${(rec.carro || CARRO_PADRAO).split(" ")[0]}${rec.destino ? ` · ${rec.origem || "?"} → ${rec.destino}` : ""}`;
    if (!(await confirmar(`Excluir a viagem de ${rotulo}?\nApaga do aparelho e da planilha. Não dá pra desfazer.`))) return;
    if (!navigator.onLine) { avisar("Sem internet agora — conecte pra excluir (senão o lançamento voltaria na próxima sincronização)."); return; }
    try {
      await apiDeleteLancamento(rec.data, rec.carro || CARRO_PADRAO, usuario.nome);
      mutateRecords(recs => recs.filter(x => x.id !== rec.id));
      setInlineEdit(null);
      setUltimaSync(new Date());
      setSyncStatus("deleted");
    } catch (e) {
      avisar("Erro ao excluir: " + (e.message || "tente novamente."));
    }
  }

  // ═══ RENDER ═══
  return (
    <div className="min-h-screen font-sans" style={{ background: "#F4F6FA" }}>
      <DialogHost />

      {/* Cabeçalho fixo */}
      <div style={{ background: BTJ_NAVY }} className="text-white pt-6 pb-2.5 px-4 sticky top-0 z-20">
        <div className="max-w-lg mx-auto flex items-center justify-between">
          <div className="flex items-center gap-2">
            <div className="bg-white px-1.5 py-0.5 rounded-sm">
              <img src={logoUrl} alt="BTJ" className="h-5" onError={e => { e.target.outerHTML = '<span style="color:#001F3E;font-weight:700;letter-spacing:1px;">BTJ</span>'; }} />
            </div>
            <div>
              <p className="text-sm font-semibold leading-tight">{usuario.nome.split(" ").slice(0, 2).join(" ")}</p>
              <p className="text-[11px]" style={{ color: BTJ_LIGHT }}>
                {usuario.setor || SETOR} · R$ {taxaVigente(config.taxas, config.colaboradores, usuario.nome, todayISO()).toFixed(2).replace(".", ",")}/km
              </p>
            </div>
          </div>
          {/* Menu de conta (⋯), sempre à direita */}
          <div className="relative">
            <button onClick={() => setMenuAberto(v => !v)} aria-label="Menu"
              className="w-9 h-9 flex items-center justify-center rounded-full text-xl leading-none"
              style={{ background: "rgba(255,255,255,0.12)", color: "#fff" }}>
              ⋯
            </button>
            {menuAberto && (
              <>
                <div className="fixed inset-0" style={{ zIndex: 30 }} onClick={() => setMenuAberto(false)} />
                <div className="absolute right-0 mt-1 w-48 bg-white rounded-xl shadow-lg overflow-hidden" style={{ zIndex: 31 }}>
                  {ehAprovador && (
                    <button onClick={() => { setMenuAberto(false); setScreen("aprovacoes"); }}
                      className="w-full text-left px-4 py-2.5 text-sm border-b border-gray-100" style={{ color: BTJ_NAVY }}>
                      ✓ Aprovações
                    </button>
                  )}
                  <button onClick={async () => { setMenuAberto(false); if (await confirmar("Sair da conta?")) { limparSessao(); setUsuario(null); } }}
                    className="w-full text-left px-4 py-2.5 text-sm" style={{ color: "#B3261E" }}>
                    ⎋ Sair da conta
                  </button>
                </div>
              </>
            )}
          </div>
        </div>
        {/* Linha 2: na home = ações (Relatório/Despesas); nas telas internas = Voltar. Sempre à direita. */}
        <div className="max-w-lg mx-auto flex items-center justify-end mt-1.5">
          {screen === "home" ? (
            <div className="flex gap-1.5">
              <button onClick={() => setScreen("resumos")}
                className="flex items-center gap-1 px-3 py-1.5 rounded-lg font-semibold text-white text-xs"
                style={{ background: BTJ_BLUE }}>
                <span className="text-sm">📊</span> Relatório
              </button>
              <button onClick={() => setScreen("despMenu")}
                className="flex items-center gap-1 px-3 py-1.5 rounded-lg font-semibold text-xs bg-white"
                style={{ color: BTJ_NAVY, border: `0.5px solid ${BTJ_BLUE}` }}>
                <span className="text-sm">💳</span> Despesas
              </button>
            </div>
          ) : (
            <button
              onClick={() => {
                setInlineEdit(null);
                const voltarPraMenu = ["despPedagio", "despOutras", "despesa", "extrato"];
                setScreen(voltarPraMenu.indexOf(screen) !== -1 ? "despMenu" : "home");
              }}
              className="text-sm font-medium" style={{ color: BTJ_BLUE }}>
              ‹ Voltar
            </button>
          )}
        </div>
        {syncStatus && (
          <p className="max-w-lg mx-auto text-[11px] mt-0.5" style={{ color: syncStatus === "error" ? "#FFD9A0" : BTJ_LIGHT }}>
            {syncStatus === "syncing" && "☁ sincronizando..."}
            {syncStatus === "saving" && "☁ gravando apontamento..."}
            {syncStatus === "synced" && `🔄 App sincronizado ${tempoRelativo(ultimaSync)}`}
            {syncStatus === "ok" && "✅ Apontamento gravado na base de dados"}
            {syncStatus === "deleted" && "🗑 Lançamento excluído da base de dados"}
            {syncStatus === "error" && "📱 Apontamento gravado LOCAL (sem sinal) — envio automático quando houver conexão"}
          </p>
        )}
        {!online && !syncStatus && (
          <p className="max-w-lg mx-auto text-[11px] mt-0.5" style={{ color: BTJ_LIGHT }}>✈ modo offline — tudo fica salvo no aparelho</p>
        )}
      </div>

      <div className="max-w-lg mx-auto px-3 pb-6">

        {needRefresh && (
          <button
            onClick={() => { updateSWRef.current?.(true); setNeedRefresh(false); }}
            className="w-full mt-2.5 rounded-lg py-2 text-xs font-medium text-white"
            style={{ background: BTJ_BLUE }}
          >
            ↻ Nova versão disponível — toque para atualizar
          </button>
        )}

        {/* ═══ TELA PRINCIPAL ═══ */}
        {screen === "home" && (
          <>
            {/* Aviso: ciclo anterior fechou e ainda não foi enviado */}
            {(() => {
              const st = statusPeriodo(perAnterior, envios, hojeKey);
              const temDados = records.some(r => periodKey(r.data) === perAnterior) || despesas.some(d => periodKey(d.data) === perAnterior);
              if (st === "enviado" || !temDados) return null;
              return (
                <button onClick={() => { setRevisaoPeriodo(perAnterior); setScreen("revisao"); }}
                  className="w-full mt-2.5 flex items-center justify-between rounded-lg px-3 py-2.5 text-left"
                  style={{ background: "#E6F1FB", borderLeft: "3px solid #1A9BE0" }}>
                  <span className="text-xs" style={{ color: "#0C447C" }}>
                    {st === "pendente"
                      ? <>⟳ O relatório de <b>{periodLabel(perAnterior)}</b> está fechado — envio pendente de conexão</>
                      : <>📋 O ciclo de <b>{periodLabel(perAnterior)}</b> fechou dia 25 — revise e envie o relatório de reembolso</>}
                  </span>
                  <span className="text-xs font-semibold shrink-0 ml-2" style={{ color: "#1A9BE0" }}>Revisar ›</span>
                </button>
              );
            })()}

            {/* Alerta de pendências (sublista) */}
            {openDays.length > 0 && (
              <div className="mt-2.5">
                <button
                  onClick={() => setShowPending(s => !s)}
                  className="w-full flex items-center justify-between rounded-lg px-3 py-2 text-left"
                  style={{ background: "#FAECE7", borderLeft: "3px solid #D85A30" }}
                >
                  <span className="text-xs" style={{ color: "#712B13" }}>
                    ⚠ {openDays.length} {openDays.length === 1 ? "dia com apontamento pendente" : "dias com apontamento pendente"}
                  </span>
                  <span className="text-xs" style={{ color: "#993C1D" }}>{showPending ? "▲" : "▼"}</span>
                </button>
                {showPending && (
                  <div className="mt-1 space-y-1">
                    {openDays.map(r => (
                      <button
                        key={r.id}
                        onClick={() => { setFData(r.data); setFCarro(r.carro || CARRO_PADRAO); setShowPending(false); }}
                        className="w-full flex items-center justify-between bg-white rounded-lg px-3 py-2 border border-gray-100 text-left"
                      >
                        <span className="text-xs text-gray-700">
                          {formatDateBR(r.data)} · {(r.carro || CARRO_PADRAO).split(" ")[0]} · {r.kmInicial == null ? "sem KM inicial" : "sem KM final"}
                        </span>
                        <span className="text-xs font-medium" style={{ color: BTJ_BLUE }}>completar</span>
                      </button>
                    ))}
                  </div>
                )}
              </div>
            )}

            {/* Formulário do dia */}
            <Card className="mt-2.5 p-3.5">
              <div className="mb-2.5">
                <p className="text-[11px] text-gray-500 mb-0.5">🚗 Carro</p>
                <select
                  value={fCarro}
                  onChange={e => {
                    if (e.target.value === NOVO_CARRO_VALUE) { setMostrarCadastroCarroPrincipal(true); return; }
                    setFCarro(e.target.value); saveLastCar(e.target.value);
                  }}
                  className="w-full border border-gray-200 rounded-lg px-2 py-1.5 text-sm bg-white font-medium"
                  style={{ color: BTJ_NAVY }}
                >
                  {(config.carros || DEFAULT_CONFIG.carros).map(c => <option key={c} value={c}>{c}</option>)}
                  <option value={NOVO_CARRO_VALUE}>+ Outro (cadastrar novo)</option>
                </select>
                {mostrarCadastroCarroPrincipal && (
                  <CadastrarCarroModal
                    colaborador={usuario.nome}
                    onCancel={() => setMostrarCadastroCarroPrincipal(false)}
                    onSaved={(novoCarro) => {
                      handleNovoCarro(novoCarro);
                      setFCarro(novoCarro); saveLastCar(novoCarro);
                      setMostrarCadastroCarroPrincipal(false);
                    }}
                  />
                )}
              </div>
              <div className="flex gap-2 mb-2.5">
                <div className="flex-1">
                  <p className="text-[11px] text-gray-500 mb-0.5">Data</p>
                  <input
                    type="date"
                    value={fData}
                    max={todayISO()}
                    onChange={e => setFData(e.target.value)}
                    className="w-full border border-gray-200 rounded-lg px-2 py-1.5 text-sm"
                  />
                </div>
                <div className="flex-1">
                  <p className="text-[11px] text-gray-500 mb-0.5">
                    Saída {fOrigemGps && <span style={{ color: BTJ_BLUE }}>· via GPS</span>}
                  </p>
                  <input
                    type="text"
                    value={fOrigem}
                    onChange={e => { setFOrigem(e.target.value); setFOrigemGps(false); }}
                    onKeyDown={e => { if (e.key === "Enter") e.currentTarget.blur(); }}
                    placeholder="detectando..."
                    className="w-full border border-gray-200 rounded-lg px-2 py-1.5 text-sm"
                  />
                </div>
              </div>

              <div className="mb-2.5">
                <p className="text-[11px] text-gray-500 mb-0.5">Destino (opcional)</p>
                <select
                  value={fDestino}
                  onChange={e => setFDestino(e.target.value)}
                  className="w-full border border-gray-200 rounded-lg px-2 py-1.5 text-sm bg-white"
                >
                  <option value="">— selecionar —</option>
                  {destinos.map(d => <option key={d} value={d}>{d}</option>)}
                </select>
              </div>

              <div className="flex gap-2 mb-2.5">
                <KmBox
                  label="KM inicial"
                  value={fKmIni}
                  state={queuedIni ? "queued" : fKmIni != null ? "done" : "pending"}
                  loading={loadingIni}
                  error={coherErr?.field === "ini"}
                  onCamera={() => fileIniCamRef.current?.click()}
                  onGallery={() => fileIniGalRef.current?.click()}
                  onManual={v => setFKmIni(v)}
                />
                <KmBox
                  label="KM final"
                  value={fKmFin}
                  state={queuedFin ? "queued" : fKmFin != null ? "done" : "pending"}
                  loading={loadingFin}
                  error={coherErr?.field === "fin"}
                  onCamera={() => fileFinCamRef.current?.click()}
                  onGallery={() => fileFinGalRef.current?.click()}
                  onManual={v => setFKmFin(v)}
                />
              </div>
              <input ref={fileIniCamRef} type="file" accept="image/*" capture="environment" className="hidden"
                onChange={e => { handlePhoto("inicial", e.target.files[0]); e.target.value = ""; }} />
              <input ref={fileIniGalRef} type="file" accept="image/*" className="hidden"
                onChange={e => { handlePhoto("inicial", e.target.files[0]); e.target.value = ""; }} />
              <input ref={fileFinCamRef} type="file" accept="image/*" capture="environment" className="hidden"
                onChange={e => { handlePhoto("final", e.target.files[0]); e.target.value = ""; }} />
              <input ref={fileFinGalRef} type="file" accept="image/*" className="hidden"
                onChange={e => { handlePhoto("final", e.target.files[0]); e.target.value = ""; }} />

              {coherErr && (
                <div className="rounded-lg px-2.5 py-2 mb-2.5" style={{ background: "#FDECEC" }}>
                  <span className="text-[11px]" style={{ color: "#C62A2F" }}>⛔ {coherErr.msg}</span>
                </div>
              )}

              {fKmIni != null && fKmFin != null && fKmFin >= fKmIni && (
                <div className="rounded-lg px-2.5 py-2 mb-2" style={{ background: "#E6F1FB" }}>
                  <div className="flex items-center justify-between">
                    <span className="text-[11px]" style={{ color: "#185FA5" }}>Saldo do dia</span>
                    <span className="text-sm font-semibold" style={{ color: "#0C447C" }}>
                      {(fKmFin - fKmIni).toLocaleString("pt-BR")} km · R$ {((fKmFin - fKmIni) * taxaVigente(config.taxas, config.colaboradores, usuario.nome, fData)).toLocaleString("pt-BR", { minimumFractionDigits: 2 })}
                    </span>
                  </div>
                </div>
              )}

              {prevClosed && (
                <div className="rounded-lg px-2.5 py-2 mb-2.5" style={{ background: "#F5F7FA" }}>
                  <div className="flex items-center justify-between">
                    <span className="text-[10px] text-gray-400">
                      ↩ anterior · {formatDateShort(prevClosed.data)} · {(prevClosed.carro || CARRO_PADRAO).split(" ")[0]}{prevClosed.destino ? ` · ${prevClosed.origem || "?"} → ${prevClosed.destino}` : ""}
                    </span>
                    <span className="text-[11px] text-gray-500">
                      ini <b>{prevClosed.kmInicial?.toLocaleString("pt-BR") ?? "—"}</b> · fim <b>{prevClosed.kmFinal.toLocaleString("pt-BR")}</b>
                    </span>
                  </div>
                  {showReuse && (
                    <button
                      onClick={() => setFKmIni(prevClosed.kmFinal)}
                      className="w-full mt-1.5 rounded-md py-1.5 text-[11px] font-medium"
                      style={{ background: "#E6F1FB", color: "#185FA5", border: "0.5px solid #B8D9F5" }}
                    >
                      ↧ usar {prevClosed.kmFinal.toLocaleString("pt-BR")} como KM inicial de hoje
                    </button>
                  )}
                </div>
              )}

              <div className="mb-3">
                <p className="text-[11px] text-gray-500 mb-0.5">Observação do dia</p>
                <input
                  type="text"
                  value={fObs}
                  onChange={e => setFObs(e.target.value)}
                  onKeyDown={e => { if (e.key === "Enter") e.currentTarget.blur(); }}
                  placeholder="Sobre o que foi a viagem..."
                  className="w-full border border-gray-200 rounded-lg px-2 py-1.5 text-sm"
                />
              </div>

              <button
                onClick={save}
                disabled={!!coherErr}
                className="w-full rounded-xl py-2.5 text-sm font-semibold text-white disabled:cursor-not-allowed"
                style={{ background: coherErr ? "#E7E5DE" : BTJ_BLUE, color: coherErr ? "#A8A69E" : "#fff" }}
              >
                {coherErr ? "Salvar (corrija o KM)" : (fKmIni != null && fKmFin != null ? "Salvar apontamento" : "Salvar — completar depois")}
              </button>
            </Card>

            {/* Métricas do mês */}
            <div className="flex gap-2 mt-2.5">
              <div className="flex-1 bg-white border border-gray-100 rounded-xl p-2 text-center">
                <p className="text-[10px] text-gray-400">💼 Trabalho</p>
                <p className="text-sm font-semibold" style={{ color: BTJ_NAVY }}>{cur.trabalho.toLocaleString("pt-BR")} km</p>
              </div>
              <div className="flex-1 bg-white border border-gray-100 rounded-xl p-2 text-center">
                <p className="text-[10px] text-gray-400">🛣️ Pedágio</p>
                <p className="text-sm font-semibold" style={{ color: "#854F0B" }}>
                  R$ {cur.pedagioMes.toLocaleString("pt-BR", { minimumFractionDigits: 2 })}
                </p>
              </div>
              <div className="flex-1 rounded-xl p-2 text-center" style={{ background: BTJ_NAVY }}>
                <p className="text-[10px]" style={{ color: BTJ_LIGHT }}>A receber</p>
                <p className="text-sm font-semibold" style={{ color: BTJ_BLUE }}>
                  R$ {cur.receber.toLocaleString("pt-BR", { minimumFractionDigits: 2 })}
                </p>
              </div>
            </div>
            {kmHoje > 0 && (
              <p className="text-center text-[11px] text-gray-400 mt-1.5">hoje: {kmHoje.toLocaleString("pt-BR")} km</p>
            )}
          </>
        )}

        {/* ═══ MENU DE DESPESAS ═══ */}
        {/* ═══ REVISÃO DO RELATÓRIO ═══ */}
        {screen === "revisao" && revisaoPeriodo && (
          <RevisaoRelatorio
            periodo={revisaoPeriodo}
            records={records}
            despesas={despesas}
            taxas={config.taxas}
            colaboradores={config.colaboradores}
            usuario={usuario}
            envio={envios[revisaoPeriodo] || null}
            reabertura={minhasReab.filter(r => r.periodo === revisaoPeriodo).slice(-1)[0] || null}
            onVoltar={() => { setScreen("resumos"); setRevisaoPeriodo(null); }}
            onEmitido={(ret, assinatura) => setEnvioPeriodo(revisaoPeriodo, { status: "enviado", retorno: ret, assinaturaBase64: assinatura, reaberturaId: envios[revisaoPeriodo]?.reaberturaId })}
            onPendente={(assinatura) => setEnvioPeriodo(revisaoPeriodo, { status: "pendente", assinaturaBase64: assinatura, reaberturaId: envios[revisaoPeriodo]?.reaberturaId })}
          />
        )}

        {screen === "aprovacoes" && (
          <Aprovacoes usuario={usuario} onVoltar={() => setScreen("home")} />
        )}

        {screen === "despMenu" && (
          <Card className="mt-2.5 p-4">
            <h2 className="font-semibold text-gray-800 mb-0.5">Despesas</h2>
            <p className="text-[11px] text-gray-400 mb-3">Veja e edite os lançamentos já feitos</p>
            <div className="space-y-2">
              <button onClick={() => setScreen("despPedagio")}
                className="w-full flex items-center justify-between rounded-xl px-4 py-3 border border-gray-100 text-left">
                <span className="text-sm text-gray-800">🛣️ Ver pedágios <span className="text-gray-400 font-normal">· resumo</span></span>
                <span className="text-gray-300">›</span>
              </button>
              <button onClick={() => setScreen("despOutras")}
                className="w-full flex items-center justify-between rounded-xl px-4 py-3 border border-gray-100 text-left">
                <span className="text-sm text-gray-800">🧾 Ver outras despesas <span className="text-gray-400 font-normal">· resumo</span></span>
                <span className="text-gray-300">›</span>
              </button>
            </div>
            <div className="flex gap-2 mt-4">
              <button onClick={() => setScreen("despesa")}
                className="flex-1 rounded-xl py-2.5 text-sm font-medium text-white" style={{ background: BTJ_BLUE }}>
                + Despesa
              </button>
              <button onClick={() => setScreen("extrato")}
                className="flex-1 rounded-xl py-2.5 text-sm font-medium border" style={{ borderColor: BTJ_BLUE, color: BTJ_NAVY }}>
                📄 Importar extrato do tag
              </button>
            </div>
          </Card>
        )}

        {/* ═══ LISTA: PEDÁGIOS ═══ */}
        {screen === "despPedagio" && (
          <GestaoPedagios
            despesas={despesas.filter(d => (d.tipo || "").toLowerCase().indexOf("ped") === 0)}
            records={records}
            travado={periodoTravado}
            onChange={refreshDespesas}
            onAdd={() => setScreen("extrato")}
          />
        )}

        {/* ═══ LISTA: OUTRAS DESPESAS ═══ */}
        {screen === "despOutras" && (
          <GestaoDespesas
            titulo="Outras despesas" icone="🧾"
            despesas={despesas.filter(d => (d.tipo || "").toLowerCase().indexOf("ped") !== 0)}
            carros={config.carros || DEFAULT_CONFIG.carros}
            travado={periodoTravado}
            onChange={refreshDespesas}
            onAdd={() => setScreen("despesa")}
            addLabel="+ Despesa"
          />
        )}

        {/* ═══ NOVA DESPESA ═══ */}
        {screen === "despesa" && (
          <DespesaManual
            carros={config.carros || DEFAULT_CONFIG.carros}
            carroInicial={fCarro}
            limites={config.limites || DEFAULT_CONFIG.limites}
            faixa={faixaDoColaborador(config.colaboradores || DEFAULT_CONFIG.colaboradores, usuario.nome)}
            colaborador={usuario.nome}
            travado={periodoTravado}
            onNovoCarro={handleNovoCarro}
            onSaved={() => { setScreen("despMenu"); refreshDespesas(); }}
            onCancel={() => setScreen("despMenu")}
          />
        )}

        {/* ═══ IMPORTAR EXTRATO ═══ */}
        {screen === "extrato" && (
          <ImportarExtrato
            carros={config.carros || DEFAULT_CONFIG.carros}
            carroInicial={fCarro}
            records={records}
            colaborador={usuario.nome}
            travado={periodoTravado}
            onNovoCarro={handleNovoCarro}
            onDone={() => { setScreen("despMenu"); refreshDespesas(); }}
            onCancel={() => setScreen("despMenu")}
          />
        )}

        {/* ═══ TELA DE RESUMOS ═══ */}
        {screen === "resumos" && (
          <>
            <div className="rounded-xl p-3.5 mt-2.5" style={{ background: BTJ_NAVY }}>
              <p className="text-xs font-medium mb-2" style={{ color: BTJ_LIGHT }}>{monthLabelFromKey(curKey)}</p>
              <div className="flex gap-4 items-end">
                <div>
                  <p className="text-[10px]" style={{ color: BTJ_LIGHT }}>Trabalho</p>
                  <p className="text-base font-semibold text-white">{cur.trabalho.toLocaleString("pt-BR")} km</p>
                </div>
                <div>
                  <p className="text-[10px]" style={{ color: BTJ_LIGHT }}>Pedágio</p>
                  <p className="text-base font-semibold text-white">R$ {cur.pedagioMes.toLocaleString("pt-BR", { minimumFractionDigits: 2 })}</p>
                </div>
                <div className="ml-auto text-right">
                  <p className="text-[10px]" style={{ color: BTJ_LIGHT }}>A receber</p>
                  <p className="text-base font-semibold" style={{ color: BTJ_BLUE }}>
                    R$ {cur.receber.toLocaleString("pt-BR", { minimumFractionDigits: 2 })}
                  </p>
                </div>
              </div>
            </div>

            <div className="mt-2.5 space-y-2">
              {/* Seletor de agrupamento: Mês (ciclo 26→25) ou Semana (dom–sáb) */}
              <div className="flex justify-end">
                <div className="inline-flex rounded-lg overflow-hidden border" style={{ borderColor: BTJ_BLUE }}>
                  {[["mes", "Mês"], ["semana", "Semana"]].map(([v, rotulo]) => (
                    <button key={v}
                      onClick={() => { setAgrupamento(v); setExpandedMonth(v === "mes" ? periodKey(todayISO()) : weekKey(todayISO())); }}
                      className="px-3 py-1 text-xs font-medium"
                      style={agrupamento === v ? { background: BTJ_BLUE, color: "#fff" } : { background: "#fff", color: BTJ_NAVY }}>
                      {rotulo}
                    </button>
                  ))}
                </div>
              </div>

              {groupKeys.length === 0 && (
                <p className="text-center text-sm text-gray-400 mt-6">Nenhum apontamento ainda.</p>
              )}
              {groupKeys.map(key => {
                const s = monthSummary(records, key, config.taxas, config.colaboradores, usuario.nome, despesas, agrupamento === "mes" ? periodKey : weekKey);
                const isCur = key === curGroupKey;
                const opened = expandedMonth === key;
                return (
                  <div key={key} className="bg-white border border-gray-100 rounded-xl overflow-hidden">
                    <button
                      onClick={() => setExpandedMonth(opened ? null : key)}
                      className="w-full flex items-center justify-between px-3.5 py-3 text-left"
                    >
                      <div>
                        <p className="text-sm font-semibold" style={{ color: isCur ? BTJ_NAVY : "#2C2C2A" }}>
                          {agrupamento === "mes" ? monthLabelFromKey(key) : `Semana ${weekLabel(key)}`}{" "}
                          {agrupamento === "mes" && (() => {
                            const st = statusPeriodo(key, envios, hojeKey);
                            if (st === "aberto") return <span className="text-[10px] px-2 py-0.5 rounded-full" style={{ background: "#E6F1FB", color: "#185FA5" }}>⏳ aberto</span>;
                            if (st === "enviado") return <span className="text-[10px] px-2 py-0.5 rounded-full" style={{ background: "#E1F5EE", color: "#085041" }}>✓ enviado</span>;
                            if (st === "pendente") return <span className="text-[10px] px-2 py-0.5 rounded-full" style={{ background: "#FEF3E2", color: "#854F0B" }}>⟳ envio pendente</span>;
                            if (st === "reaberto") return <span className="text-[10px] px-2 py-0.5 rounded-full" style={{ background: "#FFF7E6", color: "#854F0B" }}>🔓 reaberto pra correção</span>;
                            return <span className="text-[10px] px-2 py-0.5 rounded-full" style={{ background: "#FEF3E2", color: "#854F0B" }}>🔓 fechado · não enviado</span>;
                          })()}
                          {agrupamento === "semana" && isCur && <span className="text-[10px] px-2 py-0.5 rounded-full" style={{ background: "#E6F1FB", color: "#185FA5" }}>atual</span>}
                        </p>
                        <p className="text-[11px] text-gray-400">
                          {s.viagens} viagens · {s.trabalho.toLocaleString("pt-BR")} km
                          {s.pessoal != null && ` · pessoal ${s.pessoal.toLocaleString("pt-BR")} km`}
                          {" · "}R$ {s.receber.toLocaleString("pt-BR", { minimumFractionDigits: 2 })}
                        </p>
                        {agrupamento === "mes" && !isCur && (
                          <button onClick={(e) => { e.stopPropagation(); setRevisaoPeriodo(key); setScreen("revisao"); }}
                            className="text-[11px] font-semibold mt-0.5" style={{ color: BTJ_BLUE }}>
                            📋 {(() => { const st = statusPeriodo(key, envios, hojeKey); return st === "enviado" ? "Ver envio / reenviar" : st === "reaberto" ? "Corrigir e reenviar" : "Revisar e enviar"; })()} ›
                          </button>
                        )}
                      </div>
                      <span className="text-gray-400 text-xs">{opened ? "▲" : "▼"}</span>
                    </button>
                    {opened && (
                      <div className="border-t border-gray-100">
                        {[...s.recs].reverse().map(r => (
                          <div key={r.id} className="border-b border-gray-50 last:border-b-0">
                            {inlineEdit?.id === r.id ? (
                              <div className="px-3.5 py-3" style={{ background: "#F8FAFC" }}>
                                <p className="text-[11px] font-medium mb-2" style={{ color: "#185FA5" }}>
                                  Editando {formatDateShort(r.data)} · {weekdayAbrev(r.data)} · {(r.carro || CARRO_PADRAO).split(" ")[0]}{r.destino ? ` · ${r.origem || "?"} → ${r.destino}` : ""}
                                </p>
                                <div className="flex gap-1.5 mb-2">
                                  <div className={`flex-1 rounded-lg p-1.5 text-center ${inlineEdit.err?.field === "ini" ? "border-2 border-red-500" : ""}`} style={{ background: "#E1F5EE" }}>
                                    <p className="text-[9px]" style={{ color: "#0F6E56" }}>KM inicial</p>
                                    <input type="number" inputMode="numeric" value={inlineEdit.kmIni ?? ""} placeholder="—"
                                      onChange={e => changeInline({ kmIni: e.target.value === "" ? null : Number(e.target.value) })}
                                      onKeyDown={e => { if (e.key === "Enter") e.currentTarget.blur(); }}
                                      className="w-full bg-transparent text-center text-sm font-semibold focus:outline-none" style={{ color: "#04342C" }} />
                                  </div>
                                  <div className={`flex-1 rounded-lg p-1.5 text-center ${inlineEdit.err?.field === "fin" ? "border-2 border-red-500" : ""}`} style={{ background: "#E1F5EE" }}>
                                    <p className="text-[9px]" style={{ color: "#0F6E56" }}>KM final</p>
                                    <input type="number" inputMode="numeric" value={inlineEdit.kmFin ?? ""} placeholder="—"
                                      onChange={e => changeInline({ kmFin: e.target.value === "" ? null : Number(e.target.value) })}
                                      onKeyDown={e => { if (e.key === "Enter") e.currentTarget.blur(); }}
                                      className="w-full bg-transparent text-center text-sm font-semibold focus:outline-none" style={{ color: "#04342C" }} />
                                  </div>
                                </div>
                                {inlineEdit.err && <p className="text-[10px] mb-2" style={{ color: "#C62A2F" }}>⛔ {inlineEdit.err.msg}</p>}
                                <input type="text" value={inlineEdit.obs} placeholder="observação..."
                                  onChange={e => changeInline({ obs: e.target.value })}
                                  onKeyDown={e => { if (e.key === "Enter") e.currentTarget.blur(); }}
                                  className="w-full border border-gray-200 rounded-lg px-2 py-1.5 text-xs mb-2" />
                                <div className="flex gap-1.5">
                                  <button onClick={saveInline} disabled={!!inlineEdit.err}
                                    className="flex-[2] rounded-lg py-2 text-xs font-medium text-white disabled:opacity-50"
                                    style={{ background: inlineEdit.err ? "#E7E5DE" : BTJ_BLUE }}>
                                    Salvar alterações
                                  </button>
                                  <button onClick={() => setInlineEdit(null)}
                                    className="flex-1 rounded-lg py-2 text-xs text-gray-600 border border-gray-200">
                                    Cancelar
                                  </button>
                                </div>
                                <button onClick={() => excluirDia(r)}
                                  className="w-full rounded-lg py-1.5 text-[11px] mt-1.5 border" style={{ borderColor: "#F4C7C3", color: "#B3261E" }}>
                                  🗑 Excluir este dia (apaga a viagem do {formatDateShort(r.data)})
                                </button>
                              </div>
                            ) : r.soPedagio ? (
                              <div className="w-full flex items-center justify-between px-3.5 py-2">
                                <span className="text-xs text-gray-400">{formatDateShort(r.data)} · {weekdayAbrev(r.data)} · sem viagem registrada</span>
                                <span className="text-xs font-medium" style={{ color: "#854F0B" }}>
                                  🛣️ R$ {(s.pedagioPorDia[r.data] || 0).toLocaleString("pt-BR", { minimumFractionDigits: 2 })} de pedágio
                                </span>
                              </div>
                            ) : (
                              <button onClick={() => openInline(r)} className="w-full flex items-center justify-between px-3.5 py-2 text-left">
                                {isOpen(r) ? (
                                  <>
                                    <span className="text-xs" style={{ color: "#D85A30" }}>
                                      ⚠ {formatDateShort(r.data)} · {weekdayAbrev(r.data)} · {(r.carro || CARRO_PADRAO).split(" ")[0]} · {r.kmInicial == null ? "KM inicial pendente" : "KM final pendente"}
                                    </span>
                                    <span className="text-xs font-medium" style={{ color: BTJ_BLUE }}>completar</span>
                                  </>
                                ) : (
                                  <div className="w-full">
                                    <div className="flex items-center justify-between">
                                      <span className="text-xs text-gray-700 font-medium">
                                        {formatDateShort(r.data)} · {weekdayAbrev(r.data)} · {r.origem || "?"}{r.destino ? ` → ${r.destino}` : ""}
                                      </span>
                                      <span className="text-xs text-gray-600 font-medium">
                                        {kmOf(r).toLocaleString("pt-BR")} km · R$ {(kmOf(r) * taxaVigente(config.taxas, config.colaboradores, usuario.nome, r.data)).toLocaleString("pt-BR", { minimumFractionDigits: 2 })} ✎
                                      </span>
                                    </div>
                                    <div className="flex items-center justify-between mt-0.5 gap-2">
                                      <span className="text-[10px] text-gray-400 truncate">
                                        🚗 {(r.carro || CARRO_PADRAO).split(" ")[0]} · {r.kmInicial?.toLocaleString("pt-BR") ?? "—"} → {r.kmFinal?.toLocaleString("pt-BR") ?? "—"}
                                      </span>
                                      {r.observacao && <span className="text-[10px] text-gray-400 truncate ml-2 max-w-[35%] shrink-0">{r.observacao}</span>}
                                    </div>
                                    {(s.pedagioItensPorDia[r.data] || []).length > 0 && (
                                      <div className="mt-1">
                                        <button
                                          onClick={(e) => { e.stopPropagation(); setPedagioAberto(pedagioAberto === r.data ? null : r.data); }}
                                          className="flex items-center gap-1"
                                        >
                                          <span className="text-[11px] font-medium" style={{ color: "#854F0B" }}>
                                            🛣️ R$ {s.pedagioPorDia[r.data].toLocaleString("pt-BR", { minimumFractionDigits: 2 })} · {s.pedagioItensPorDia[r.data].length} passage{s.pedagioItensPorDia[r.data].length > 1 ? "ns" : "m"}
                                          </span>
                                          <span className="text-[9px]" style={{ color: "#854F0B" }}>{pedagioAberto === r.data ? "▲" : "▼"}</span>
                                        </button>
                                        {pedagioAberto === r.data && (
                                          <div className="mt-1 ml-1 pl-2 space-y-0.5" style={{ borderLeft: "2px solid #F5C97A" }}>
                                            {s.pedagioItensPorDia[r.data].map((p, i) => (
                                              <div key={i} className="flex items-center justify-between">
                                                <span className="text-[10px] text-gray-500">{p.local}</span>
                                                <span className="text-[10px] text-gray-500">R$ {p.valor.toLocaleString("pt-BR", { minimumFractionDigits: 2 })}</span>
                                              </div>
                                            ))}
                                          </div>
                                        )}
                                      </div>
                                    )}
                                  </div>
                                )}
                              </button>
                            )}
                          </div>
                        ))}
                      </div>
                    )}
                  </div>
                );
              })}
            </div>
          </>
        )}
      </div>
    </div>
  );
}
