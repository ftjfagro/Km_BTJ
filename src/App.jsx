import { useState, useEffect, useRef } from "react";
import { registerSW } from "virtual:pwa-register";
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
  carros: ["Corolla FSZ8B48", "Outlander FXJ5336"],
  taxas: [
    { colaborador: SOLICITANTE, taxa: 1.12, vigenteDesde: "2026-01-01" },
    { colaborador: "Geral", taxa: 0.88, vigenteDesde: "2026-01-01" },
  ],
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
const KEY_LAST_CAR = "km_ultimo_carro";

function loadLastCar() {
  return localStorage.getItem(KEY_LAST_CAR) || CARRO_PADRAO;
}
function saveLastCar(c) {
  localStorage.setItem(KEY_LAST_CAR, c);
}
const KEY_PHOTO_QUEUE = "km_fotos_pendentes";

function loadRecords() {
  try {
    const raw = localStorage.getItem(KEY_RECORDS);
    if (raw) return JSON.parse(raw);
    const old = localStorage.getItem(KEY_RECORDS_OLD);
    if (old) {
      const migrated = JSON.parse(old).map(r => ({ ...r, observacao: "", synced: true }));
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

async function apiList() {
  const res = await fetch(APPS_SCRIPT_URL, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({ action: "list", colaborador: SOLICITANTE }),
  });
  const data = await res.json();
  if (!data.ok) throw new Error(data.error || "Erro ao listar lançamentos");
  return data.lancamentos || [];
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
function monthSummary(records, key, taxas) {
  const recs = records.filter(r => monthKey(r.data) === key).sort((a, b) => a.data.localeCompare(b.data));
  const trabalho = recs.reduce((s, r) => s + kmOf(r), 0);
  const receber = recs.reduce((s, r) => s + kmOf(r) * taxaVigente(taxas, SOLICITANTE, r.data), 0);

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
    if (odos.length >= 2) {
      const span = odos[odos.length - 1].v - odos[0].v;
      const trabC = recsC.reduce((s, r) => s + kmOf(r), 0);
      pessoal = (pessoal || 0) + Math.max(0, span - trabC);
    }
  }
  return { recs, trabalho, receber, pessoal, viagens: recs.filter(r => kmOf(r) > 0).length };
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
const TIPOS_DESPESA = ["Pedágio", "Alimentação", "Estacionamento", "Outros"];

function DespesaManual({ carros, carroInicial, onSaved, onCancel }) {
  const [data, setData] = useState(todayISO());
  const [carro, setCarro] = useState(carroInicial);
  const [carroOutro, setCarroOutro] = useState("");
  const [tipo, setTipo] = useState("Pedágio");
  const [valor, setValor] = useState("");
  const [descricao, setDescricao] = useState("");
  const [compB64, setCompB64] = useState(null);
  const [compNome, setCompNome] = useState("");
  const [saving, setSaving] = useState(false);
  const camRef = useRef(null);
  const galRef = useRef(null);

  const carroFinal = carro === "Outro (digitar)" ? carroOutro.trim() : carro;

  async function pickComprovante(file) {
    if (!file) return;
    try {
      const b64 = await fileToResizedBase64(file, 1600, 0.7);
      setCompB64(b64);
      setCompNome(file.name || "comprovante.jpg");
    } catch { alert("Não consegui processar a imagem do comprovante."); }
  }

  async function salvar() {
    if (!valor || Number(valor) <= 0) { alert("Informe o valor da despesa."); return; }
    if (carro === "Outro (digitar)" && !carroOutro.trim()) { alert("Digite o carro (apelido/placa)."); return; }
    setSaving(true);
    try {
      await apiSaveDespesa({
        data, carro: carroFinal, tipo, valor: Number(valor),
        descricao, comprovanteImage: compB64 || undefined, origem: "manual",
      });
      onSaved();
    } catch (e) {
      alert("Erro ao salvar: " + (e.message || "tente novamente."));
    } finally {
      setSaving(false);
    }
  }

  return (
    <Card className="mt-2.5 p-3.5">
      <div className="flex gap-2 mb-2.5">
        <div className="flex-1">
          <p className="text-[11px] text-gray-500 mb-0.5">Data</p>
          <input type="date" value={data} max={todayISO()} onChange={e => setData(e.target.value)}
            className="w-full border border-gray-200 rounded-lg px-2 py-1.5 text-sm" />
        </div>
        <div className="flex-1">
          <p className="text-[11px] text-gray-500 mb-0.5">🚗 Carro</p>
          <select value={carro} onChange={e => setCarro(e.target.value)}
            className="w-full border border-gray-200 rounded-lg px-2 py-1.5 text-sm bg-white">
            {carros.map(c => <option key={c} value={c}>{c}</option>)}
            <option value="Outro (digitar)">Outro (digitar)</option>
          </select>
        </div>
      </div>

      {carro === "Outro (digitar)" && (
        <input type="text" value={carroOutro} onChange={e => setCarroOutro(e.target.value)}
          onKeyDown={e => { if (e.key === "Enter") e.currentTarget.blur(); }}
          placeholder="Ex: Fiat Argo ABC1D23 (alugado)"
          className="w-full border border-gray-200 rounded-lg px-2 py-1.5 text-sm mb-2.5" />
      )}

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

      <div className="flex gap-2 mb-2.5">
        <div className="flex-1">
          <p className="text-[11px] text-gray-500 mb-0.5">Valor (R$)</p>
          <input type="number" inputMode="decimal" value={valor} onChange={e => setValor(e.target.value)}
            onKeyDown={e => { if (e.key === "Enter") e.currentTarget.blur(); }}
            placeholder="0,00" className="w-full border border-gray-200 rounded-lg px-2 py-1.5 text-sm" />
        </div>
        <div className="flex-[1.4]">
          <p className="text-[11px] text-gray-500 mb-0.5">Descrição (opcional)</p>
          <input type="text" value={descricao} onChange={e => setDescricao(e.target.value)}
            onKeyDown={e => { if (e.key === "Enter") e.currentTarget.blur(); }}
            className="w-full border border-gray-200 rounded-lg px-2 py-1.5 text-sm" />
        </div>
      </div>

      <p className="text-[11px] text-gray-500 mb-1">Comprovante (opcional)</p>
      <div className="flex gap-1.5 mb-3">
        <button onClick={() => camRef.current?.click()} className="flex-1 rounded-lg py-2 text-sm bg-amber-400" aria-label="Foto do comprovante">📷 Foto</button>
        <button onClick={() => galRef.current?.click()} className="flex-1 rounded-lg py-2 text-sm bg-amber-400" aria-label="Galeria">🖼️ Galeria</button>
      </div>
      <input ref={camRef} type="file" accept="image/*" capture="environment" className="hidden" onChange={e => { pickComprovante(e.target.files[0]); e.target.value = ""; }} />
      <input ref={galRef} type="file" accept="image/*" className="hidden" onChange={e => { pickComprovante(e.target.files[0]); e.target.value = ""; }} />
      {compNome && <p className="text-[10px] mb-3" style={{ color: "#0F6E56" }}>🧾 {compNome} — será enviado ao Drive ao salvar</p>}

      <div className="flex gap-2">
        <button onClick={salvar} disabled={saving}
          className="flex-[2] rounded-xl py-2.5 text-sm font-semibold text-white disabled:opacity-60"
          style={{ background: BTJ_BLUE }}>
          {saving ? "Salvando..." : "Salvar despesa"}
        </button>
        <button onClick={onCancel} className="flex-1 rounded-xl py-2.5 text-sm text-gray-600 border border-gray-200">Cancelar</button>
      </div>
    </Card>
  );
}

// ─── Tela: Importar Extrato de Pedágio ────────────────────────────────────────
function ImportarExtrato({ carros, carroInicial, onDone, onCancel }) {
  const [carro, setCarro] = useState(carroInicial);
  const [carroOutro, setCarroOutro] = useState("");
  const [passagens, setPassagens] = useState(null); // null = ainda não leu
  const [sel, setSel] = useState({});
  const [loading, setLoading] = useState(false);
  const [saving, setSaving] = useState(false);
  const camRef = useRef(null);
  const galRef = useRef(null);

  const carroFinal = carro === "Outro (digitar)" ? carroOutro.trim() : carro;

  async function lerExtrato(file) {
    if (!file) return;
    if (carro === "Outro (digitar)" && !carroOutro.trim()) { alert("Digite o carro antes de importar."); return; }
    setLoading(true);
    try {
      const b64 = await fileToResizedBase64(file, 1600, 0.7);
      const lista = await apiOcrExtrato(b64);
      if (!lista.length) { alert("Não encontrei passagens nesse print. Tente uma imagem mais nítida."); return; }
      setPassagens(lista);
      const inicial = {};
      lista.forEach((_, i) => { inicial[i] = true; }); // tudo marcado por padrão
      setSel(inicial);
    } catch (e) {
      alert("Erro ao ler o extrato: " + (e.message || "tente novamente."));
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
      for (const p of marcadas) {
        await apiSaveDespesa({
          data: p.data, carro: carroFinal, tipo: "Pedágio",
          valor: Number(p.valor) || 0, descricao: p.local || "", origem: "extrato",
        });
      }
      alert(`${marcadas.length} pedágio(s) lançado(s) na planilha (R$ ${totalSel.toLocaleString("pt-BR", { minimumFractionDigits: 2 })}).`);
      onDone();
    } catch (e) {
      alert("Erro ao lançar: " + (e.message || "tente novamente."));
    } finally {
      setSaving(false);
    }
  }

  return (
    <Card className="mt-2.5 p-3.5">
      <div className="mb-2.5">
        <p className="text-[11px] text-gray-500 mb-0.5">🚗 Carro deste extrato</p>
        <select value={carro} onChange={e => setCarro(e.target.value)}
          className="w-full border border-gray-200 rounded-lg px-2 py-1.5 text-sm bg-white">
          {carros.map(c => <option key={c} value={c}>{c}</option>)}
          <option value="Outro (digitar)">Outro (digitar)</option>
        </select>
        {carro === "Outro (digitar)" && (
          <input type="text" value={carroOutro} onChange={e => setCarroOutro(e.target.value)}
            onKeyDown={e => { if (e.key === "Enter") e.currentTarget.blur(); }}
            placeholder="Ex: Fiat Argo ABC1D23"
            className="w-full border border-gray-200 rounded-lg px-2 py-1.5 text-sm mt-2" />
        )}
      </div>

      {!passagens && (
        <>
          <p className="text-xs text-gray-500 mb-2">Mande o print do extrato do cartão de pedágio. Leio as passagens e você seleciona as de trabalho.</p>
          <div className="flex gap-1.5">
            <button onClick={() => camRef.current?.click()} disabled={loading} className="flex-1 rounded-lg py-2.5 text-sm bg-amber-400">
              {loading ? "⟳ lendo..." : "📷 Foto"}
            </button>
            <button onClick={() => galRef.current?.click()} disabled={loading} className="flex-1 rounded-lg py-2.5 text-sm bg-amber-400">🖼️ Galeria</button>
          </div>
          <input ref={camRef} type="file" accept="image/*" capture="environment" className="hidden" onChange={e => { lerExtrato(e.target.files[0]); e.target.value = ""; }} />
          <input ref={galRef} type="file" accept="image/*" className="hidden" onChange={e => { lerExtrato(e.target.files[0]); e.target.value = ""; }} />
        </>
      )}

      {passagens && (
        <>
          <div className="flex justify-between items-center mb-2">
            <span className="text-[11px] text-gray-500">Desmarque as de uso pessoal</span>
            <span className="text-[11px] font-medium" style={{ color: "#185FA5" }}>{qtdSel} de {passagens.length}</span>
          </div>
          <div className="border border-gray-100 rounded-xl overflow-hidden mb-3">
            {passagens.map((p, i) => (
              <button key={i} onClick={() => setSel(s => ({ ...s, [i]: !s[i] }))}
                className="w-full flex items-center gap-2.5 px-3 py-2 border-b border-gray-50 last:border-b-0 text-left">
                <div className="w-[18px] h-[18px] rounded-md flex items-center justify-center text-xs text-white shrink-0"
                  style={{ background: sel[i] ? BTJ_BLUE : "transparent", border: sel[i] ? "none" : "1.5px solid #CFCFC8" }}>
                  {sel[i] ? "✓" : ""}
                </div>
                <div className="flex-1">
                  <p className={`text-xs ${sel[i] ? "text-gray-800" : "text-gray-400 line-through"}`}>{p.local || "Pedágio"}</p>
                  <p className={`text-[10px] ${sel[i] ? "text-gray-500" : "text-gray-400"}`}>{formatDateShort(p.data)} · {weekdayPT(p.data)}</p>
                </div>
                <span className={`text-xs font-semibold ${sel[i] ? "" : "text-gray-400 line-through"}`} style={sel[i] ? { color: "#04342C" } : {}}>
                  R$ {(Number(p.valor) || 0).toLocaleString("pt-BR", { minimumFractionDigits: 2 })}
                </span>
              </button>
            ))}
          </div>
          <div className="flex justify-between items-center mb-3">
            <span className="text-sm font-semibold text-gray-800">Total selecionado</span>
            <span className="text-base font-bold" style={{ color: BTJ_NAVY }}>R$ {totalSel.toLocaleString("pt-BR", { minimumFractionDigits: 2 })}</span>
          </div>
          <div className="flex gap-2">
            <button onClick={lancar} disabled={saving || qtdSel === 0}
              className="flex-[2] rounded-xl py-2.5 text-sm font-semibold text-white disabled:opacity-60"
              style={{ background: BTJ_BLUE }}>
              {saving ? "Lançando..." : `Lançar ${qtdSel} pedágio(s)`}
            </button>
            <button onClick={onCancel} className="flex-1 rounded-xl py-2.5 text-sm text-gray-600 border border-gray-200">Cancelar</button>
          </div>
        </>
      )}
    </Card>
  );
}

// ─── App ──────────────────────────────────────────────────────────────────────
export default function App() {
  const [records, setRecords] = useState([]);
  const [config, setConfig] = useState(loadCachedConfig());
  const [screen, setScreen] = useState("home"); // home | resumos | despesa | extrato
  const [online, setOnline] = useState(navigator.onLine);
  const [syncStatus, setSyncStatus] = useState(null);
  const [needRefresh, setNeedRefresh] = useState(false);
  const updateSWRef = useRef(null);
  const [showPending, setShowPending] = useState(false);
  const [expandedMonth, setExpandedMonth] = useState(monthKey(todayISO()));
  const [inlineEdit, setInlineEdit] = useState(null); // { id, kmIni, kmFin, obs, err }

  // Formulário (apontamento do dia)
  const [editingId, setEditingId] = useState(null);
  const [fData, setFData] = useState(todayISO());
  const [fCarro, setFCarro] = useState(loadLastCar());
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
    setRecords(loadRecords());
    apiFetchConfig()
      .then(c => { setConfig(c); localStorage.setItem(KEY_CONFIG, JSON.stringify(c)); })
      .catch(() => {});
    pullAndReconcile(); // baixa da planilha (fonte mestra) e alinha o local
    const on = () => setOnline(true);
    const off = () => setOnline(false);
    window.addEventListener("online", on);
    window.addEventListener("offline", off);
    return () => { window.removeEventListener("online", on); window.removeEventListener("offline", off); };
  }, []);

  // Baixa os lançamentos da planilha e reconcilia com o local.
  // Regra: pendentes de envio (synced=false) têm prioridade e são preservados;
  // para todo o resto, a planilha é a fonte da verdade.
  async function pullAndReconcile() {
    if (!navigator.onLine) return;
    try {
      setSyncStatus("syncing");
      const remote = await apiList();
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
      setSyncStatus("ok");
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

  async function processPhotoQueue() {
    let q = loadPhotoQueue();
    if (!q.length) return;
    setSyncStatus("syncing");
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
    setSyncStatus("ok");
  }

  async function resyncUnsynced() {
    const pending = loadRecords().filter(r => r.synced === false);
    for (const r of pending) {
      try {
        await apiSave(r);
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
    mutateRecords(recs => {
      const idx = recs.findIndex(r => r.data === dateISO && (r.carro || CARRO_PADRAO) === carro);
      if (idx === -1) return recs;
      const upd = { ...recs[idx], [phase === "inicial" ? "kmInicial" : "kmFinal"]: km, synced: false };
      const next = [...recs];
      next[idx] = upd;
      return next;
    });
    if (dateISO === fData && carro === fCarro) {
      if (phase === "inicial") setFKmIni(km); else setFKmFin(km);
    }
    syncRecordByKey(dateISO, carro);
  }

  async function syncRecordByKey(dateISO, carro) {
    const r = loadRecords().find(x => x.data === dateISO && (x.carro || CARRO_PADRAO) === carro);
    if (!r) return;
    try {
      setSyncStatus("syncing");
      await apiSave(r);
      mutateRecords(recs => recs.map(x => x.id === r.id ? { ...x, synced: true } : x));
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

  // ── KM emendado: foto inicial de hoje pode fechar o dia aberto anterior (MESMO carro) ──
  function offerBridge(km) {
    const open = records
      .filter(r => r.data < fData && (r.carro || CARRO_PADRAO) === fCarro && r.kmInicial != null && r.kmFinal == null)
      .sort((a, b) => b.data.localeCompare(a.data))[0];
    if (!open) return;
    const err = checkCoherence(records, open.data, open.kmInicial, km, open.id, fCarro);
    if (err) return;
    const ok = confirm(
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
      alert("Preencha ao menos um KM, destino ou observação.");
      return;
    }
    const err = checkCoherence(records, fData, fKmIni, fKmFin, editingId, fCarro);
    if (err) {
      setCoherErr(err);
      alert("⛔ " + err.msg);
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
    syncRecordByKey(fData, fCarro);
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

  // ── Edição inline na tela de resumos ──
  function openInline(r) {
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
    if (err) { setInlineEdit(prev => ({ ...prev, err })); alert("⛔ " + err.msg); return; }
    const updated = { ...rec, kmInicial: inlineEdit.kmIni, kmFinal: inlineEdit.kmFin, observacao: inlineEdit.obs, synced: false };
    mutateRecords(recs => recs.map(x => x.id === rec.id ? updated : x));
    setInlineEdit(null);
    syncRecordByKey(rec.data, rec.carro || CARRO_PADRAO);
  }

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
          {screen === "home" ? (
            <button
              onClick={() => setScreen("resumos")}
              className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg font-semibold text-white text-xs"
              style={{ background: BTJ_BLUE }}
            >
              <span className="text-sm">📊</span> Relatório
            </button>
          ) : (
            <button
              onClick={() => { setScreen("home"); setInlineEdit(null); }}
              className="text-sm font-medium"
              style={{ color: BTJ_BLUE }}
            >
              ‹ Voltar
            </button>
          )}
        </div>
        {syncStatus && (
          <p className="max-w-lg mx-auto text-[11px] mt-1" style={{ color: syncStatus === "error" ? "#FFD9A0" : BTJ_LIGHT }}>
            {syncStatus === "syncing" && "☁ gravando na base de dados..."}
            {syncStatus === "ok" && "✅ Apontamento gravado na base de dados"}
            {syncStatus === "error" && "📱 Apontamento gravado LOCAL (sem sinal) — envio automático quando houver conexão"}
          </p>
        )}
        {!online && !syncStatus && (
          <p className="max-w-lg mx-auto text-[11px] mt-1" style={{ color: BTJ_LIGHT }}>✈ modo offline — tudo fica salvo no aparelho</p>
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
                  onChange={e => { setFCarro(e.target.value); saveLastCar(e.target.value); }}
                  className="w-full border border-gray-200 rounded-lg px-2 py-1.5 text-sm bg-white font-medium"
                  style={{ color: BTJ_NAVY }}
                >
                  {(config.carros || DEFAULT_CONFIG.carros).map(c => <option key={c} value={c}>{c}</option>)}
                </select>
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

            {/* Ações de despesa */}
            <div className="flex gap-2 mt-3">
              <button
                onClick={() => setScreen("despesa")}
                className="flex-1 rounded-xl py-2.5 text-sm font-medium border"
                style={{ borderColor: BTJ_BLUE, color: BTJ_NAVY, background: "#fff" }}
              >
                + Despesa
              </button>
              <button
                onClick={() => setScreen("extrato")}
                className="flex-1 rounded-xl py-2.5 text-sm font-medium border"
                style={{ borderColor: BTJ_BLUE, color: BTJ_NAVY, background: "#fff" }}
              >
                📄 Importar pedágio
              </button>
            </div>
          </>
        )}

        {/* ═══ NOVA DESPESA ═══ */}
        {screen === "despesa" && (
          <DespesaManual
            carros={config.carros || DEFAULT_CONFIG.carros}
            carroInicial={fCarro}
            onSaved={() => { setScreen("home"); pullAndReconcile(); }}
            onCancel={() => setScreen("home")}
          />
        )}

        {/* ═══ IMPORTAR EXTRATO ═══ */}
        {screen === "extrato" && (
          <ImportarExtrato
            carros={config.carros || DEFAULT_CONFIG.carros}
            carroInicial={fCarro}
            onDone={() => { setScreen("home"); pullAndReconcile(); }}
            onCancel={() => setScreen("home")}
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
                        {[...s.recs].reverse().map(r => (
                          <div key={r.id} className="border-b border-gray-50 last:border-b-0">
                            {inlineEdit?.id === r.id ? (
                              <div className="px-3.5 py-3" style={{ background: "#F8FAFC" }}>
                                <p className="text-[11px] font-medium mb-2" style={{ color: "#185FA5" }}>
                                  Editando {formatDateShort(r.data)} · {(r.carro || CARRO_PADRAO).split(" ")[0]}{r.destino ? ` · ${r.origem || "?"} → ${r.destino}` : ""}
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
                              </div>
                            ) : (
                              <button onClick={() => openInline(r)} className="w-full flex items-center justify-between px-3.5 py-2 text-left">
                                {isOpen(r) ? (
                                  <>
                                    <span className="text-xs" style={{ color: "#D85A30" }}>
                                      ⚠ {formatDateShort(r.data)} · {(r.carro || CARRO_PADRAO).split(" ")[0]} · {r.kmInicial == null ? "KM inicial pendente" : "KM final pendente"}
                                    </span>
                                    <span className="text-xs font-medium" style={{ color: BTJ_BLUE }}>completar</span>
                                  </>
                                ) : (
                                  <div className="w-full">
                                    <div className="flex items-center justify-between">
                                      <span className="text-xs text-gray-700 font-medium">
                                        {formatDateShort(r.data)} · {r.origem || "?"}{r.destino ? ` → ${r.destino}` : ""}
                                      </span>
                                      <span className="text-xs text-gray-600 font-medium">
                                        {kmOf(r).toLocaleString("pt-BR")} km · R$ {(kmOf(r) * taxaVigente(config.taxas, SOLICITANTE, r.data)).toLocaleString("pt-BR", { minimumFractionDigits: 2 })} ✎
                                      </span>
                                    </div>
                                    <div className="flex items-center justify-between mt-0.5">
                                      <span className="text-[10px] text-gray-400">
                                        🚗 {(r.carro || CARRO_PADRAO).split(" ")[0]} · {r.kmInicial?.toLocaleString("pt-BR") ?? "—"} → {r.kmFinal?.toLocaleString("pt-BR") ?? "—"}
                                      </span>
                                      {r.observacao && <span className="text-[10px] text-gray-400 truncate ml-2 max-w-[45%]">{r.observacao}</span>}
                                    </div>
                                  </div>
                                )}
                              </button>
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
