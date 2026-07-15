import { useState, useEffect } from "react";
import * as XLSX from "xlsx";

// ─── Config ───────────────────────────────────────────────────────────────────
const RATE_PER_KM = 0.88; // R$/km
const SOLICITANTE = "Felipe Torquato Junqueira Franco";
const SETOR = "Diretor";
const CPF = "372.742.538-59";

// URL fixa do Apps Script (grava lançamentos na planilha + faz a leitura do
// odômetro via OpenAI). É infraestrutura do app, não configuração do usuário —
// por isso fica embutida aqui. A chave da OpenAI mora dentro do Apps Script.
const APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbwDv9SdQANzfKb5oWQAU_7evjHN5BKLbW3kxb2eZbhVM1Weku0xUmDyHup7KRVTF8bPdw/exec";

const CITIES = ["Sud", "Ilha", "RP", "SP", "Campinas", "Jundiaí", "Outra"];

const WEEKDAYS_PT = ["domingo", "segunda-feira", "terça-feira", "quarta-feira", "quinta-feira", "sexta-feira", "sábado"];

const MONTHS_PT = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho","Agosto","Setembro","Outubro","Novembro","Dezembro"];

function todayISO() {
  return new Date().toISOString().split("T")[0];
}

function parseISO(s) {
  const [y, m, d] = s.split("-").map(Number);
  return new Date(y, m - 1, d);
}

function formatDateBR(s) {
  const [y, m, d] = s.split("-");
  return `${d}/${m}/${y}`;
}

function weekdayPT(iso) {
  return WEEKDAYS_PT[parseISO(iso).getDay()];
}

function monthLabel(iso) {
  const d = parseISO(iso);
  return `${MONTHS_PT[d.getMonth()]} ${d.getFullYear()}`;
}

// ─── Storage helpers ──────────────────────────────────────────────────────────
const KEY = "km_registros_v2";

function loadRecords() {
  try {
    const raw = localStorage.getItem(KEY);
    return raw ? JSON.parse(raw) : [];
  } catch {
    return [];
  }
}

function saveRecords(recs) {
  localStorage.setItem(KEY, JSON.stringify(recs));
}

// ─── OCR / AI helpers ─────────────────────────────────────────────────────────
// A leitura do odômetro roda no Apps Script (que guarda a chave da OpenAI),
// não no navegador — assim nenhum aparelho precisa da chave.
async function extractOdometerFromImage(base64) {
  const res = await fetch(APPS_SCRIPT_URL, {
    method: "POST",
    headers: { "Content-Type": "text/plain;charset=utf-8" },
    body: JSON.stringify({ action: "ocr", image: base64 }),
  });
  const data = await res.json();
  if (!data.ok) throw new Error(data.error || "Erro ao ler o odômetro.");
  return data.km ?? null;
}

// ─── Google Sheets sync ───────────────────────────────────────────────────────
async function syncToSheet(record) {
  try {
    await fetch(APPS_SCRIPT_URL, {
      method: "POST",
      mode: "no-cors", // Apps Script não retorna CORS headers; gravação ocorre mesmo assim
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify({
        data: record.data,
        tipo: "Viagem",
        origem: record.origem,
        destino: record.destino,
        kmInicial: record.kmInicial,
        kmFinal: record.kmFinal,
        categoria: "",
        descricao: "",
        valor: "",
        observacao: "Lançado via app Km_BTJ",
      }),
    });
    // com mode:"no-cors" não dá pra ler a resposta, então assumimos sucesso se não deu exceção
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e };
  }
}

// ─── Excel export ─────────────────────────────────────────────────────────────
function exportToExcel(records, mesRef) {
  const sorted = [...records].sort((a, b) => a.data.localeCompare(b.data));

  const wb = XLSX.utils.book_new();

  // Header rows
  const rows = [
    ["", "RELATÓRIO DE DESPESAS"],
    [],
    [],
    ["", "Solicitante", "", SOLICITANTE],
    ["", "Inserir CPF", "", CPF.replace(/\D/g, "")],
    ["", "Setor", "", SETOR],
    ["", "Referência", "", mesRef],
    [],
    ["", "INSERIR DESPESAS DE VIAGEM", "", "", "", "", "", "INSERIR  KM RODADO"],
    ["", "Data", "Dia da semana", "TIPO DE DESPESA", "DESCRIÇÃO DA DESPESA", "VALOR", "", "CIDADE DE ORIGEM", "CIDADE DESTINO", "Km Inicial", "KM Final", "KM RODADO", "R$ KM", "R$ KM TOTAL"],
  ];

  sorted.forEach(r => {
    const km = (r.kmFinal || 0) - (r.kmInicial || 0);
    const total = km * RATE_PER_KM;
    rows.push([
      "",
      formatDateBR(r.data),
      weekdayPT(r.data),
      "",
      "",
      "",
      "",
      r.origem || "",
      r.destino || "",
      r.kmInicial || "",
      r.kmFinal || "",
      km > 0 ? km : 0,
      km > 0 ? RATE_PER_KM : "",
      km > 0 ? total : 0,
    ]);
  });

  const totalKm = sorted.reduce((s, r) => s + Math.max(0, (r.kmFinal || 0) - (r.kmInicial || 0)), 0);
  const totalReais = totalKm * RATE_PER_KM;

  rows.push(
    [],
    ["", "", "", "", "", "", "", "", "", "", "", "Valor das despesas", "", 0],
    ["", "", "", "", "", "", "", "", "", "", "", "Outros", "", 0],
    ["", "", "", "", "", "", "", "", "", "", "", "Total a Receber", "", totalReais],
  );

  const ws = XLSX.utils.aoa_to_sheet(rows);
  XLSX.utils.book_append_sheet(wb, ws, "Despesas");
  XLSX.writeFile(wb, `Relatorio_KM_${mesRef.replace(/\s/g,"_")}.xlsx`);
}

// ─── Components ───────────────────────────────────────────────────────────────

function Badge({ children, color = "blue" }) {
  const colors = {
    blue: "bg-blue-100 text-blue-800",
    green: "bg-green-100 text-green-800",
    amber: "bg-amber-100 text-amber-800",
    red: "bg-red-100 text-red-800",
  };
  return (
    <span className={`inline-flex items-center px-2 py-0.5 rounded text-xs font-medium ${colors[color]}`}>
      {children}
    </span>
  );
}

function Card({ children, className = "" }) {
  return (
    <div className={`bg-white rounded-2xl shadow-sm border border-gray-100 ${className}`}>
      {children}
    </div>
  );
}

// ─── New Entry Form ────────────────────────────────────────────────────────────
function EntryForm({ onSave, onCancel, initial = null }) {
  const [data, setData] = useState(initial?.data ?? todayISO());
  const [origem, setOrigem] = useState(initial?.origem ?? "Sud");
  const [origemCustom, setOrigemCustom] = useState("");
  const [destino, setDestino] = useState(initial?.destino ?? "Ilha");
  const [destinoCustom, setDestinoCustom] = useState("");
  const [kmInicial, setKmInicial] = useState(initial?.kmInicial ?? "");
  const [kmFinal, setKmFinal] = useState(initial?.kmFinal ?? "");
  const [loading, setLoading] = useState(false);
  const [imgPhase, setImgPhase] = useState(null); // 'inicial' | 'final'

  const kmRodado = Math.max(0, (Number(kmFinal) || 0) - (Number(kmInicial) || 0));
  const total = kmRodado * RATE_PER_KM;

  async function handlePhoto(phase, file) {
    setLoading(true);
    try {
      const b64 = await new Promise((res, rej) => {
        const r = new FileReader();
        r.onload = () => res(r.result.split(",")[1]);
        r.onerror = rej;
        r.readAsDataURL(file);
      });
      const km = await extractOdometerFromImage(b64);
      if (km) {
        if (phase === "inicial") setKmInicial(String(km));
        else setKmFinal(String(km));
      } else {
        alert("Não consegui ler o odômetro. Por favor, insira manualmente.");
      }
    } catch (e) {
      alert("Erro ao processar a imagem: " + (e.message || "tente novamente."));
    }
    setLoading(false);
    setImgPhase(null);
  }

  function submit() {
    const o = origem === "Outra" ? origemCustom : origem;
    const d = destino === "Outra" ? destinoCustom : destino;
    if (!o || !d || !kmInicial || !kmFinal) {
      alert("Preencha todos os campos obrigatórios.");
      return;
    }
    onSave({
      id: initial?.id ?? Date.now(),
      data,
      origem: o,
      destino: d,
      kmInicial: Number(kmInicial),
      kmFinal: Number(kmFinal),
    });
  }

  return (
    <div className="space-y-4">
      {/* Date */}
      <div>
        <label className="block text-sm font-medium text-gray-700 mb-1">Data</label>
        <input
          type="date"
          value={data}
          onChange={e => setData(e.target.value)}
          className="w-full rounded-xl border border-gray-200 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
        />
        <p className="text-xs text-gray-400 mt-1">{weekdayPT(data)}</p>
      </div>

      {/* Origem / Destino */}
      <div className="grid grid-cols-2 gap-3">
        <div>
          <label className="block text-sm font-medium text-gray-700 mb-1">Origem</label>
          <select
            value={origem}
            onChange={e => setOrigem(e.target.value)}
            className="w-full rounded-xl border border-gray-200 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
          >
            {CITIES.map(c => <option key={c}>{c}</option>)}
          </select>
          {origem === "Outra" && (
            <input
              placeholder="Nome da cidade"
              value={origemCustom}
              onChange={e => setOrigemCustom(e.target.value)}
              className="mt-1 w-full rounded-xl border border-gray-200 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
            />
          )}
        </div>
        <div>
          <label className="block text-sm font-medium text-gray-700 mb-1">Destino</label>
          <select
            value={destino}
            onChange={e => setDestino(e.target.value)}
            className="w-full rounded-xl border border-gray-200 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
          >
            {CITIES.map(c => <option key={c}>{c}</option>)}
          </select>
          {destino === "Outra" && (
            <input
              placeholder="Nome da cidade"
              value={destinoCustom}
              onChange={e => setDestinoCustom(e.target.value)}
              className="mt-1 w-full rounded-xl border border-gray-200 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
            />
          )}
        </div>
      </div>

      {/* KM Inicial */}
      <div>
        <label className="block text-sm font-medium text-gray-700 mb-1">KM Inicial (ODO saída)</label>
        <div className="flex gap-2">
          <input
            type="number"
            placeholder="Ex: 224209"
            value={kmInicial}
            onChange={e => setKmInicial(e.target.value)}
            className="flex-1 rounded-xl border border-gray-200 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
          />
          <label className="flex items-center gap-1 px-3 py-2 bg-blue-50 text-blue-700 rounded-xl text-sm cursor-pointer hover:bg-blue-100 transition">
            {loading && imgPhase === "inicial" ? (
              <span className="animate-spin">⟳</span>
            ) : (
              <>📷 Foto</>
            )}
            <input
              type="file"
              accept="image/*"
              className="hidden"
              onChange={e => { if (e.target.files[0]) { setImgPhase("inicial"); handlePhoto("inicial", e.target.files[0]); } }}
            />
          </label>
        </div>
      </div>

      {/* KM Final */}
      <div>
        <label className="block text-sm font-medium text-gray-700 mb-1">KM Final (ODO chegada)</label>
        <div className="flex gap-2">
          <input
            type="number"
            placeholder="Ex: 224342"
            value={kmFinal}
            onChange={e => setKmFinal(e.target.value)}
            className="flex-1 rounded-xl border border-gray-200 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
          />
          <label className="flex items-center gap-1 px-3 py-2 bg-blue-50 text-blue-700 rounded-xl text-sm cursor-pointer hover:bg-blue-100 transition">
            {loading && imgPhase === "final" ? (
              <span className="animate-spin">⟳</span>
            ) : (
              <>📷 Foto</>
            )}
            <input
              type="file"
              accept="image/*"
              className="hidden"
              onChange={e => { if (e.target.files[0]) { setImgPhase("final"); handlePhoto("final", e.target.files[0]); } }}
            />
          </label>
        </div>
      </div>

      {/* Preview */}
      {kmRodado > 0 && (
        <div className="flex justify-between items-center bg-green-50 rounded-xl px-4 py-3">
          <span className="text-sm text-green-700 font-medium">{kmRodado} km rodados</span>
          <span className="text-base font-bold text-green-800">
            R$ {total.toLocaleString("pt-BR", { minimumFractionDigits: 2 })}
          </span>
        </div>
      )}

      <div className="flex gap-2 pt-1">
        <button
          onClick={submit}
          disabled={loading}
          className="flex-1 bg-blue-600 hover:bg-blue-700 text-white font-medium py-2.5 rounded-xl text-sm transition disabled:opacity-50"
        >
          {initial ? "Salvar alterações" : "Adicionar registro"}
        </button>
        {onCancel && (
          <button
            onClick={onCancel}
            className="px-4 py-2.5 rounded-xl text-sm text-gray-600 border border-gray-200 hover:bg-gray-50 transition"
          >
            Cancelar
          </button>
        )}
      </div>
    </div>
  );
}

// ─── Main App ─────────────────────────────────────────────────────────────────
export default function App() {
  const [records, setRecords] = useState([]);
  const [view, setView] = useState("list"); // list | new | edit | export
  const [editRec, setEditRec] = useState(null);
  const [filterMonth, setFilterMonth] = useState("all");
  const [syncStatus, setSyncStatus] = useState(null); // null | 'syncing' | 'ok' | 'error'
  const [mesRef, setMesRef] = useState(() => {
    const d = new Date();
    return `${MONTHS_PT[d.getMonth()]} ${d.getFullYear()}`;
  });

  useEffect(() => {
    setRecords(loadRecords());
  }, []);

  function save(recs) {
    setRecords(recs);
    saveRecords(recs);
  }

  function addRecord(r) {
    save([...records, r]);
    setView("list");
    setSyncStatus("syncing");
    syncToSheet(r).then(res => setSyncStatus(res.ok ? "ok" : "error"));
  }

  function updateRecord(r) {
    save(records.map(x => x.id === r.id ? r : x));
    setView("list");
    setEditRec(null);
    setSyncStatus("syncing");
    syncToSheet(r).then(res => setSyncStatus(res.ok ? "ok" : "error"));
  }

  function deleteRecord(id) {
    if (!confirm("Excluir este registro?")) return;
    save(records.filter(x => x.id !== id));
  }

  // Filter
  const months = [...new Set(records.map(r => monthLabel(r.data)))].sort();
  const filtered = filterMonth === "all" ? records : records.filter(r => monthLabel(r.data) === filterMonth);
  const sorted = [...filtered].sort((a, b) => a.data.localeCompare(b.data));

  const totalKm = sorted.reduce((s, r) => s + Math.max(0, (r.kmFinal || 0) - (r.kmInicial || 0)), 0);
  const totalReais = totalKm * RATE_PER_KM;

  return (
    <div className="min-h-screen bg-gray-50 font-sans">
      {/* Header */}
      <div className="bg-blue-700 text-white pt-10 pb-6 px-4">
        <div className="max-w-lg mx-auto">
          <div>
            <p className="text-blue-200 text-xs mb-0.5">Relatório de Quilometragem</p>
            <h1 className="text-xl font-bold leading-tight">Felipe Torquato</h1>
            <p className="text-blue-300 text-xs mt-0.5">{SETOR} · R$ {RATE_PER_KM.toFixed(2)}/km</p>
          </div>

          {syncStatus && (
            <p className={`text-xs mt-1 ${syncStatus === "error" ? "text-red-300" : "text-blue-300"}`}>
              {syncStatus === "syncing" && "☁ sincronizando com a planilha..."}
              {syncStatus === "ok" && "☁ sincronizado com a planilha"}
              {syncStatus === "error" && "⚠ falha ao sincronizar (registro salvo localmente)"}
            </p>
          )}

          {/* Summary */}
          <div className="mt-4 grid grid-cols-2 gap-3">
            <div className="bg-blue-800 rounded-xl px-4 py-3">
              <p className="text-blue-300 text-xs">KM rodados</p>
              <p className="text-2xl font-bold">{totalKm.toLocaleString("pt-BR")}</p>
            </div>
            <div className="bg-blue-800 rounded-xl px-4 py-3">
              <p className="text-blue-300 text-xs">Total a receber</p>
              <p className="text-2xl font-bold">
                R$ {totalReais.toLocaleString("pt-BR", { minimumFractionDigits: 2 })}
              </p>
            </div>
          </div>
        </div>
      </div>

      <div className="max-w-lg mx-auto px-4 -mt-2 pb-24">

        {view === "new" && (
          <Card className="p-5 mt-4">
            <h2 className="font-semibold text-gray-800 mb-4">Novo registro</h2>
            <EntryForm onSave={addRecord} onCancel={() => setView("list")} />
          </Card>
        )}

        {view === "edit" && editRec && (
          <Card className="p-5 mt-4">
            <h2 className="font-semibold text-gray-800 mb-4">Editar registro</h2>
            <EntryForm
              initial={editRec}
              onSave={updateRecord}
              onCancel={() => { setView("list"); setEditRec(null); }}
            />
          </Card>
        )}

        {view === "export" && (
          <Card className="p-5 mt-4">
            <h2 className="font-semibold text-gray-800 mb-4">Exportar Excel</h2>
            <div className="space-y-4">
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-1">Mês de referência</label>
                <select
                  value={mesRef}
                  onChange={e => setMesRef(e.target.value)}
                  className="w-full rounded-xl border border-gray-200 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
                >
                  {MONTHS_PT.map((m, i) => {
                    const yr = new Date().getFullYear();
                    return <option key={m}>{m} {yr}</option>;
                  })}
                </select>
              </div>
              <div className="bg-gray-50 rounded-xl p-4 text-sm text-gray-600 space-y-1">
                <p>📋 {sorted.length} registros ({filterMonth === "all" ? "todos" : filterMonth})</p>
                <p>🚗 {totalKm.toLocaleString("pt-BR")} km rodados</p>
                <p className="font-semibold text-gray-800">
                  💰 R$ {totalReais.toLocaleString("pt-BR", { minimumFractionDigits: 2 })} a receber
                </p>
              </div>
              <button
                onClick={() => exportToExcel(sorted, mesRef)}
                className="w-full bg-green-600 hover:bg-green-700 text-white font-medium py-3 rounded-xl text-sm transition"
              >
                ⬇ Baixar .xlsx
              </button>
              <button
                onClick={() => setView("list")}
                className="w-full py-2.5 rounded-xl text-sm text-gray-600 border border-gray-200 hover:bg-gray-50 transition"
              >
                Voltar
              </button>
            </div>
          </Card>
        )}

        {view === "list" && (
          <>
            {/* Actions */}
            <div className="flex gap-2 mt-4">
              <button
                onClick={() => setView("new")}
                className="flex-1 bg-blue-600 hover:bg-blue-700 text-white font-medium py-3 rounded-xl text-sm transition flex items-center justify-center gap-1.5"
              >
                + Novo registro
              </button>
              <button
                onClick={() => setView("export")}
                className="px-4 py-3 bg-green-600 hover:bg-green-700 text-white rounded-xl text-sm font-medium transition"
              >
                ⬇ Excel
              </button>
            </div>

            {/* Filter */}
            {months.length > 1 && (
              <div className="mt-3 flex gap-2 overflow-x-auto pb-1">
                <button
                  onClick={() => setFilterMonth("all")}
                  className={`flex-none px-3 py-1.5 rounded-full text-xs font-medium transition ${
                    filterMonth === "all" ? "bg-blue-600 text-white" : "bg-white text-gray-600 border border-gray-200"
                  }`}
                >
                  Todos
                </button>
                {months.map(m => (
                  <button
                    key={m}
                    onClick={() => setFilterMonth(m)}
                    className={`flex-none px-3 py-1.5 rounded-full text-xs font-medium transition ${
                      filterMonth === m ? "bg-blue-600 text-white" : "bg-white text-gray-600 border border-gray-200"
                    }`}
                  >
                    {m}
                  </button>
                ))}
              </div>
            )}

            {/* Records */}
            <div className="mt-3 space-y-2">
              {sorted.length === 0 ? (
                <Card className="p-8 text-center">
                  <p className="text-4xl mb-2">🚗</p>
                  <p className="text-gray-500 text-sm">Nenhum registro ainda.</p>
                  <p className="text-gray-400 text-xs mt-1">Toque em "+ Novo registro" para começar.</p>
                </Card>
              ) : (
                sorted.map(r => {
                  const km = Math.max(0, (r.kmFinal || 0) - (r.kmInicial || 0));
                  const tot = km * RATE_PER_KM;
                  return (
                    <Card key={r.id} className="p-4">
                      <div className="flex items-start justify-between gap-2">
                        <div className="flex-1 min-w-0">
                          <div className="flex items-center gap-2 mb-0.5">
                            <span className="text-sm font-semibold text-gray-800">{formatDateBR(r.data)}</span>
                            <Badge color="blue">{weekdayPT(r.data)}</Badge>
                          </div>
                          <p className="text-sm text-gray-600">
                            {r.origem} → {r.destino}
                          </p>
                          <p className="text-xs text-gray-400 mt-0.5">
                            ODO: {r.kmInicial?.toLocaleString("pt-BR")} → {r.kmFinal?.toLocaleString("pt-BR")}
                          </p>
                        </div>
                        <div className="text-right flex-none">
                          <p className="text-base font-bold text-green-700">
                            R$ {tot.toLocaleString("pt-BR", { minimumFractionDigits: 2 })}
                          </p>
                          <p className="text-xs text-gray-400">{km.toLocaleString("pt-BR")} km</p>
                        </div>
                      </div>
                      <div className="flex gap-2 mt-3 pt-2 border-t border-gray-50">
                        <button
                          onClick={() => { setEditRec(r); setView("edit"); }}
                          className="flex-1 text-xs text-blue-600 py-1.5 rounded-lg hover:bg-blue-50 transition font-medium"
                        >
                          ✏ Editar
                        </button>
                        <button
                          onClick={() => deleteRecord(r.id)}
                          className="flex-1 text-xs text-red-500 py-1.5 rounded-lg hover:bg-red-50 transition font-medium"
                        >
                          🗑 Excluir
                        </button>
                      </div>
                    </Card>
                  );
                })
              )}
            </div>
          </>
        )}
      </div>
    </div>
  );
}

