import { useState, useEffect, useRef } from "react";
import * as XLSX from "xlsx";

// ─── Config fixa (infraestrutura) ─────────────────────────────────────────────
const APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbwDv9SdQANzfKb5oWQAU_7evjHN5BKLbW3kxb2eZbhVM1Weku0xUmDyHup7KRVTF8bPdw/exec";
const SOLICITANTE = "Felipe Torquato Junqueira Franco";
const SETOR = "Diretor";
const CPF = "372.742.538-59";

// Identidade BTJ
const BTJ_NAVY = "#001F3E";
const BTJ_BLUE = "#1EABE3";
const BTJ_LIGHT = "#A1D6DC";

// Fallback caso a config da planilha não carregue (offline no primeiro uso)
const DEFAULT_CONFIG = {
  destinos: ["Sud", "Ilha", "RP", "SP", "Campinas", "Jundiaí", "Ribeirão", "VCP", "Foods", "Sud Foods", "Foods RP", "Foods Prudente", "Prudente"],
  taxas: [
    { colaborador: SOLICITANTE, taxa: 1.12, vigenteDesde: "2026-01-01" },
    { colaborador: "Geral", taxa: 0.88, vigenteDesde: "2026-01-01" },
  ],
};

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
function formatDateBR(s) {
  const [y, m, d] = s.split("-");
  return `${d}/${m}/${y}`;
}
function formatDateShort(s) {
  const [, m, d] = s.split("-");
  return `${d}/${m}`;
}
function weekdayPT(iso) {
  return WEEKDAYS_PT[parseISO(iso).getDay()];
}
function monthKey(iso) {
  return iso.slice(0, 7); // "2026-07"
}
function monthLabelFromKey(key) {
  const [y, m] = key.split("-").map(Number);
  return `${MONTHS_PT[m - 1]} ${y}`;
}

// ─── Taxa vigente (mesma regra do Apps Script, no cliente) ────────────────────
function taxaVigente(taxas, colaborador, isoDate) {
  const d = parseISO(isoDate);
  function melhor(nome) {
    let mTaxa = null, mData = null;
    for (const t of taxas) {
      if (String(t.colaborador).toLowerCase() !== nome.toLowerCase()) continue;
      const v = typeof t.vigenteDesde === "string" ? new Date(t.vigenteDesde) : new Date(t.vigenteDesde);
      if (v > d) continue;
      if (mData === null || v > mData) { mData = v; mTaxa = Number(t.taxa); }
    }
    return mTaxa;
  }
  return melhor(colaborador) ?? melhor("Geral") ?? 0.88;
}

// ─── Storage ──────────────────────────────────────────────────────────────────
const KEY_RECORDS = "km_registros_v3";
const KEY_RECORDS_OLD = "km_registros_v2";
const KEY_CONFIG = "km_config_cache";
const KEY_PHOTO_QUEUE = "km_fotos_pendentes";

function loadRecords() {
  try {
    const raw = localStorage.getItem(KEY_RECORDS);
    if (raw) return JSON.parse(raw);
    const old = localStorage.getItem(KEY_RECORDS_OLD);
    if (old) {
      const migrated = JSON.parse(old).map(r => ({ ...r, observacao: "", synced: true, logRow: null }));
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
  catch { alert("Memória do navegador cheia — a foto não pôde ser guardada offline. Digite o KM manualmente."); }
}

// ─── API (Apps Script) ────────────────────────────────────────────────────────
async function apiFetchConfig() {
  const res = await fetch(APPS_SCRIPT_URL, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({ action: "config" }),
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

async function apiSave(record) {
  const res = await fetch(APPS_SCRIPT_URL, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({
      data: record.data,
      tipo: "Viagem",
      origem: record.origem || "",
      destino: record.destino || "",
      kmInicial: record.kmInicial ?? "",
      kmFinal: record.kmFinal ?? "",
      observacao: record.observacao || "",
      colaborador: SOLICITANTE,
      logRow: record.logRow || undefined,
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
function exportToExcel(records, mesLabel, taxas) {
  const sorted = [...records].sort((a, b) => a.data.localeCompare(b.data));
  const wb = XLSX.utils.book_new();
  const rows = [
    ["", "RELATÓRIO DE DESPESAS"],
    [], [],
    ["", "Solicitante", "", SOLICITANTE],
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
    const taxa = taxaVigente(taxas, SOLICITANTE, r.data);
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

// ─── Cálculos de resumo ───────────────────────────────────────────────────────
function kmOf(r) {
  if (r.kmInicial == null || r.kmFinal == null) return 0;
  return Math.max(0, r.kmFinal - r.kmInicial);
}
function isOpen(r) {
  return r.kmInicial == null || r.kmFinal == null;
}
function monthSummary(records, key, taxas) {
  const recs = records.filter(r => monthKey(r.data) === key).sort((a, b) => a.data.localeCompare(b.data));
  const trabalho = recs.reduce((s, r) => s + kmOf(r), 0);
  const receber = recs.reduce((s, r) => s + kmOf(r) * taxaVigente(taxas, SOLICITANTE, r.data), 0);
  const odos = [];
  recs.forEach(r => {
    if (r.kmInicial != null) odos.push({ d: r.data + "A", v: r.kmInicial });
    if (r.kmFinal != null) odos.push({ d: r.data + "B", v: r.kmFinal });
  });
  odos.sort((a, b) => a.d.localeCompare(b.d));
  let pessoal = null;
  if (odos.length >= 2) {
    const span = odos[odos.length - 1].v - odos[0].v;
    pessoal = Math.max(0, span - trabalho);
  }
  return { recs, trabalho, receber, pessoal, viagens: recs.filter(r => kmOf(r) > 0).length };
}

// ─── UI base ──────────────────────────────────────────────────────────────────
function Card({ children, className = "" }) {
  return <div className={`bg-white rounded-2xl border border-gray-100 shadow-sm ${className}`}>{children}</div>;
}

function KmBox({ label, value, state, onPhoto, onManual, loading }) {
  // state: "done" | "pending" | "queued"
  const styles = {
    done: { bg: "bg-emerald-50", labelC: "text-emerald-700", valC: "text-emerald-900" },
    pending: { bg: "bg-amber-50", labelC: "text-amber-700", valC: "text-amber-600" },
    queued: { bg: "bg-sky-50", labelC: "text-sky-700", valC: "text-sky-600" },
  }[state];
  return (
    <div className={`flex-1 ${styles.bg} rounded-xl p-2.5 text-center`}>
      <p className={`text-[11px] ${styles.labelC}`}>{label}</p>
      <input
        type="number"
        inputMode="numeric"
        placeholder="—"
        value={value ?? ""}
        onChange={e => onManual(e.target.value === "" ? null : Number(e.target.value))}
        className={`w-full bg-transparent text-center text-lg font-semibold ${styles.valC} focus:outline-none`}
      />
      {state === "queued" ? (
        <div className="bg-sky-200 rounded-lg py-1 text-[11px] text-sky-900">☁ foto salva — aguardando conexão</div>
      ) : (
        <button
          onClick={onPhoto}
          disabled={loading}
          className={`w-full rounded-lg py-1 text-[11px] ${state === "done" ? "bg-white text-emerald-700" : "bg-amber-400 text-amber-950 font-medium"}`}
        >
          {loading ? "⟳ lendo..." : state === "done" ? "✓ foto lida · refazer" : "📷 tirar foto"}
        </button>
      )}
    </div>
  );
}

// ─── App ──────────────────────────────────────────────────────────────────────
export default function App() {
  const [records, setRecords] = useState([]);
  const [config, setConfig] = useState(loadCachedConfig());
  const [screen, setScreen] = useState("home"); // home | resumos
  const [online, setOnline] = useState(navigator.onLine);
  const [syncStatus, setSyncStatus] = useState(null);
  const [showPending, setShowPending] = useState(false);
  const [expandedMonth, setExpandedMonth] = useState(monthKey(todayISO()));

  // Formulário (apontamento do dia)
  const [editingId, setEditingId] = useState(null);
  const [fData, setFData] = useState(todayISO());
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

  const fileIniRef = useRef(null);
  const fileFinRef = useRef(null);

  // ── Init ──
  useEffect(() => {
    setRecords(loadRecords());
    apiFetchConfig()
      .then(c => { setConfig(c); localStorage.setItem(KEY_CONFIG, JSON.stringify(c)); })
      .catch(() => {});
    const on = () => setOnline(true);
    const off = () => setOnline(false);
    window.addEventListener("online", on);
    window.addEventListener("offline", off);
    return () => { window.removeEventListener("online", on); window.removeEventListener("offline", off); };
  }, []);

  // ── Carrega o registro do dia selecionado no formulário ──
  useEffect(() => {
    const rec = records.find(r => r.data === fData && r.id !== undefined);
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
    setQueuedIni(q.some(p => p.date === fData && p.phase === "inicial"));
    setQueuedFin(q.some(p => p.date === fData && p.phase === "final"));
  }, [fData, records]);

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

  async function processPhotoQueue() {
    let q = loadPhotoQueue();
    if (!q.length) return;
    setSyncStatus("syncing");
    for (const item of [...q]) {
      try {
        const km = await apiOcr(item.b64);
        applyKmToDate(item.date, item.phase, km);
        q = q.filter(p => !(p.date === item.date && p.phase === item.phase));
        persistPhotoQueue(q);
      } catch { /* tenta de novo na próxima conexão */ }
    }
    setQueuedIni(q.some(p => p.date === fData && p.phase === "inicial"));
    setQueuedFin(q.some(p => p.date === fData && p.phase === "final"));
    setSyncStatus("ok");
  }

  async function resyncUnsynced() {
    const pending = loadRecords().filter(r => r.synced === false);
    for (const r of pending) {
      try {
        const res = await apiSave(r);
        mutateRecords(recs => recs.map(x => x.id === r.id ? { ...x, synced: true, logRow: res.row } : x));
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

  function applyKmToDate(dateISO, phase, km) {
    if (km == null) return;
    mutateRecords(recs => {
      const idx = recs.findIndex(r => r.data === dateISO);
      if (idx === -1) return recs;
      const upd = { ...recs[idx], [phase === "inicial" ? "kmInicial" : "kmFinal"]: km, synced: false };
      const next = [...recs];
      next[idx] = upd;
      return next;
    });
    if (dateISO === fData) {
      if (phase === "inicial") setFKmIni(km); else setFKmFin(km);
    }
    syncRecordByDate(dateISO);
  }

  async function syncRecordByDate(dateISO) {
    const r = loadRecords().find(x => x.data === dateISO);
    if (!r) return;
    try {
      setSyncStatus("syncing");
      const res = await apiSave(r);
      mutateRecords(recs => recs.map(x => x.id === r.id ? { ...x, synced: true, logRow: res.row } : x));
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
        const q = loadPhotoQueue().filter(p => !(p.date === fData && p.phase === phase));
        q.push({ date: fData, phase, b64, ts: Date.now() });
        persistPhotoQueue(q);
        if (phase === "inicial") setQueuedIni(true); else setQueuedFin(true);
        alert("Sem internet agora. A foto ficou guardada e será lida automaticamente quando a conexão voltar.");
        return;
      }
      const km = await apiOcr(b64);
      if (km == null) {
        alert("Não consegui ler o odômetro. Digite o KM manualmente.");
        return;
      }
      if (phase === "inicial") {
        setFKmIni(km);
        offerBridge(km);
      } else {
        setFKmFin(km);
      }
    } catch (e) {
      alert("Erro ao processar a imagem: " + (e.message || "tente novamente."));
    } finally {
      setLoading(false);
    }
  }

  // ── KM emendado: foto inicial de hoje pode fechar o dia aberto anterior ──
  function offerBridge(km) {
    const open = records
      .filter(r => r.data < fData && r.kmInicial != null && r.kmFinal == null)
      .sort((a, b) => b.data.localeCompare(a.data))[0];
    if (!open) return;
    const ok = confirm(
      `O dia ${formatDateBR(open.data)} está sem KM final.\n` +
      `Como o carro ficou parado, usar esta leitura (${km.toLocaleString("pt-BR")}) como o KM final daquele dia?`
    );
    if (ok) applyKmToDate(open.data, "final", km);
  }

  // ── Salvar apontamento (parcial permitido) ──
  function save() {
    if (fKmIni == null && fKmFin == null && !fObs && !fDestino) {
      alert("Preencha ao menos um KM, destino ou observação.");
      return;
    }
    if (fKmIni != null && fKmFin != null && fKmFin < fKmIni) {
      alert("KM final menor que o inicial — confira os valores.");
      return;
    }
    const rec = {
      id: editingId ?? Date.now(),
      data: fData,
      origem: fOrigem,
      destino: fDestino,
      kmInicial: fKmIni,
      kmFinal: fKmFin,
      observacao: fObs,
      synced: false,
      logRow: editingId ? (records.find(r => r.id === editingId)?.logRow ?? null) : null,
    };
    mutateRecords(recs => {
      const exists = recs.some(r => r.id === rec.id);
      return exists ? recs.map(r => r.id === rec.id ? rec : r) : [...recs, rec];
    });
    setEditingId(rec.id);
    syncRecordByDate(fData);
  }

  // ── Derivados ──
  const openDays = records.filter(isOpen).sort((a, b) => b.data.localeCompare(a.data));
  const curKey = monthKey(todayISO());
  const cur = monthSummary(records, curKey, config.taxas);
  const todayRec = records.find(r => r.data === todayISO());
  const kmHoje = todayRec ? kmOf(todayRec) : 0;

  const monthKeys = [...new Set(records.map(r => monthKey(r.data)))].sort().reverse().slice(0, 3);
  const destinos = config.destinos || DEFAULT_CONFIG.destinos;

  const logoUrl = `${import.meta.env.BASE_URL}logo.png`;

  // ═══ RENDER ═══
  return (
    <div className="min-h-screen font-sans" style={{ background: "#F4F6FA" }}>

      {/* Cabeçalho */}
      <div style={{ background: BTJ_NAVY }} className="text-white pt-9 pb-4 px-4">
        <div className="max-w-lg mx-auto flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="bg-white px-2 py-1 rounded-sm">
              <img src={logoUrl} alt="BTJ" className="h-6" onError={e => { e.target.outerHTML = '<span style="color:#001F3E;font-weight:700;letter-spacing:1px;">BTJ</span>'; }} />
            </div>
            <div>
              <p className="text-sm font-semibold leading-tight">Felipe Torquato</p>
              <p className="text-[11px]" style={{ color: BTJ_LIGHT }}>
                {SETOR} · R$ {taxaVigente(config.taxas, SOLICITANTE, todayISO()).toFixed(2).replace(".", ",")}/km
              </p>
            </div>
          </div>
          <button
            onClick={() => setScreen(screen === "home" ? "resumos" : "home")}
            className="text-2xl px-1"
            style={{ color: BTJ_BLUE }}
            aria-label={screen === "home" ? "Ver resumos" : "Voltar"}
          >
            {screen === "home" ? "📊" : "←"}
          </button>
        </div>
        {syncStatus && (
          <p className="max-w-lg mx-auto text-[11px] mt-1" style={{ color: syncStatus === "error" ? "#F09595" : BTJ_LIGHT }}>
            {syncStatus === "syncing" && "☁ sincronizando com a planilha..."}
            {syncStatus === "ok" && "☁ sincronizado com a planilha"}
            {syncStatus === "error" && "⚠ sem conexão — salvo no aparelho, sincroniza depois"}
          </p>
        )}
        {!online && !syncStatus && (
          <p className="max-w-lg mx-auto text-[11px] mt-1" style={{ color: BTJ_LIGHT }}>✈ modo offline — tudo fica salvo no aparelho</p>
        )}
      </div>

      <div className="max-w-lg mx-auto px-3 pb-6">

        {/* ═══ TELA PRINCIPAL ═══ */}
        {screen === "home" && (
          <>
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
                        onClick={() => { setFData(r.data); setShowPending(false); }}
                        className="w-full flex items-center justify-between bg-white rounded-lg px-3 py-2 border border-gray-100 text-left"
                      >
                        <span className="text-xs text-gray-700">
                          {formatDateBR(r.data)} · {r.kmInicial == null ? "sem KM inicial" : "sem KM final"}
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
                  onPhoto={() => fileIniRef.current?.click()}
                  onManual={v => setFKmIni(v)}
                />
                <KmBox
                  label="KM final"
                  value={fKmFin}
                  state={queuedFin ? "queued" : fKmFin != null ? "done" : "pending"}
                  loading={loadingFin}
                  onPhoto={() => fileFinRef.current?.click()}
                  onManual={v => setFKmFin(v)}
                />
              </div>
              <input ref={fileIniRef} type="file" accept="image/*" className="hidden"
                onChange={e => { handlePhoto("inicial", e.target.files[0]); e.target.value = ""; }} />
              <input ref={fileFinRef} type="file" accept="image/*" className="hidden"
                onChange={e => { handlePhoto("final", e.target.files[0]); e.target.value = ""; }} />

              <div className="mb-3">
                <p className="text-[11px] text-gray-500 mb-0.5">Observação do dia</p>
                <input
                  type="text"
                  value={fObs}
                  onChange={e => setFObs(e.target.value)}
                  placeholder="Sobre o que foi a viagem..."
                  className="w-full border border-gray-200 rounded-lg px-2 py-1.5 text-sm"
                />
              </div>

              <button
                onClick={save}
                className="w-full rounded-xl py-2.5 text-sm font-semibold text-white"
                style={{ background: BTJ_BLUE }}
              >
                {fKmIni != null && fKmFin != null ? "Salvar apontamento" : "Salvar — completar depois"}
              </button>
            </Card>

            {/* Métricas do mês */}
            <div className="flex gap-2 mt-2.5">
              <div className="flex-1 bg-white border border-gray-100 rounded-xl p-2 text-center">
                <p className="text-[10px] text-gray-400">💼 Trabalho</p>
                <p className="text-sm font-semibold" style={{ color: BTJ_NAVY }}>{cur.trabalho.toLocaleString("pt-BR")} km</p>
              </div>
              <div className="flex-1 bg-white border border-gray-100 rounded-xl p-2 text-center">
                <p className="text-[10px] text-gray-400">👤 Pessoal</p>
                <p className="text-sm font-semibold text-gray-500">
                  {cur.pessoal == null ? "—" : `${cur.pessoal.toLocaleString("pt-BR")} km`}
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
                  <p className="text-[10px]" style={{ color: BTJ_LIGHT }}>Pessoal</p>
                  <p className="text-base font-semibold text-white">{cur.pessoal == null ? "—" : `${cur.pessoal.toLocaleString("pt-BR")} km`}</p>
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
              {monthKeys.length === 0 && (
                <p className="text-center text-sm text-gray-400 mt-6">Nenhum apontamento ainda.</p>
              )}
              {monthKeys.map(key => {
                const s = monthSummary(records, key, config.taxas);
                const isCur = key === curKey;
                const opened = expandedMonth === key;
                return (
                  <div key={key} className="bg-white border border-gray-100 rounded-xl overflow-hidden">
                    <button
                      onClick={() => setExpandedMonth(opened ? null : key)}
                      className="w-full flex items-center justify-between px-3.5 py-3 text-left"
                    >
                      <div>
                        <p className="text-sm font-semibold" style={{ color: isCur ? BTJ_NAVY : "#2C2C2A" }}>
                          {monthLabelFromKey(key)}{" "}
                          {isCur && <span className="text-[10px] px-2 py-0.5 rounded-full" style={{ background: "#E6F1FB", color: "#185FA5" }}>atual</span>}
                        </p>
                        <p className="text-[11px] text-gray-400">
                          {s.viagens} viagens · {s.trabalho.toLocaleString("pt-BR")} km
                          {s.pessoal != null && ` · pessoal ${s.pessoal.toLocaleString("pt-BR")} km`}
                          {" · "}R$ {s.receber.toLocaleString("pt-BR", { minimumFractionDigits: 2 })}
                        </p>
                      </div>
                      <span className="text-gray-400 text-xs">{opened ? "▲" : "▼"}</span>
                    </button>
                    {opened && (
                      <div className="border-t border-gray-100">
                        {s.recs.map(r => (
                          <div key={r.id} className="flex items-center justify-between px-3.5 py-2 border-b border-gray-50 last:border-b-0">
                            {isOpen(r) ? (
                              <>
                                <span className="text-xs" style={{ color: "#D85A30" }}>
                                  ⚠ {formatDateShort(r.data)} · {r.kmInicial == null ? "KM inicial pendente" : "KM final pendente"}
                                </span>
                                <button
                                  onClick={() => { setFData(r.data); setScreen("home"); }}
                                  className="text-xs font-medium"
                                  style={{ color: BTJ_BLUE }}
                                >
                                  completar
                                </button>
                              </>
                            ) : (
                              <>
                                <div>
                                  <span className="text-xs text-gray-700">
                                    {formatDateShort(r.data)} · {r.origem || "?"}{r.destino ? ` → ${r.destino}` : ""}
                                  </span>
                                  {r.observacao && <p className="text-[10px] text-gray-400">{r.observacao}</p>}
                                </div>
                                <span className="text-xs text-gray-500">
                                  {kmOf(r).toLocaleString("pt-BR")} km · R$ {(kmOf(r) * taxaVigente(config.taxas, SOLICITANTE, r.data)).toLocaleString("pt-BR", { minimumFractionDigits: 2 })}
                                </span>
                              </>
                            )}
                          </div>
                        ))}
                        <div className="px-3.5 py-2">
                          <button
                            onClick={() => exportToExcel(s.recs.filter(r => !isOpen(r)), monthLabelFromKey(key), config.taxas)}
                            className="w-full rounded-lg py-2 text-xs font-medium text-white"
                            style={{ background: "#16a34a" }}
                          >
                            ⬇ Baixar Excel de {monthLabelFromKey(key)}
                          </button>
                        </div>
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
