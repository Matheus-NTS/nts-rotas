import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import {
  AreaChart, Area, BarChart, Bar, CartesianGrid, ResponsiveContainer,
  Tooltip, XAxis, YAxis,
} from "recharts";
import {
  Bike, CalendarDays, CheckCircle2, ChevronDown, ChevronRight, Download, FileCheck2,
  FileSpreadsheet, Gauge, Info, Pencil, Route, Search, Upload, Users, X,
} from "lucide-react";
import * as XLSX from "xlsx";
import {
  calculateClosing, exportClosing, formatBRL, formatDateBR, formatKm,
  parseWorkbookRows, type ClosingResult, type DriverSummary, type ProcessedRecord,
} from "../lib/closing";

type SortKey = keyof Pick<DriverSummary, "name" | "deliveries" | "daysWorked" | "validKm" | "dailyAverage" | "bonus">;

const RATE_KEY = "nts-rotas-rate";
const PROCESSING_STEPS = ["Lendo rotas", "Calculando quilômetros", "Calculando bônus", "Montando dashboard"];
const pause = (duration: number) => new Promise((resolve) => window.setTimeout(resolve, duration));

export default function Home() {
  const inputRef = useRef<HTMLInputElement>(null);
  const [dragging, setDragging] = useState(false);
  const [fileName, setFileName] = useState("");
  const [result, setResult] = useState<ClosingResult | null>(null);
  const [rate, setRate] = useState(0.25);
  const [draftRate, setDraftRate] = useState("0,25");
  const [rateModal, setRateModal] = useState(false);
  const [selected, setSelected] = useState<DriverSummary | null>(null);
  const [auditOpen, setAuditOpen] = useState(false);
  const [replaceModal, setReplaceModal] = useState(false);
  const [processingStep, setProcessingStep] = useState(-1);
  const [exportSuccess, setExportSuccess] = useState(false);
  const [error, setError] = useState("");
  const [sort, setSort] = useState<{ key: SortKey; direction: "asc" | "desc" }>({
    key: "validKm", direction: "desc",
  });
  const [query, setQuery] = useState("");

  useEffect(() => {
    const stored = localStorage.getItem(RATE_KEY);
    const saved = stored === null ? Number.NaN : Number(stored);
    if (Number.isFinite(saved) && saved >= 0) setRate(saved);
  }, []);

  const recalculate = useCallback((base: ClosingResult, nextRate: number) => {
    setResult(calculateClosing(base.sourceRows, nextRate, base.fileName));
  }, []);

  const readFile = useCallback(async (file?: File) => {
    if (!file) return;
    setError("");
    if (!/\.(xlsx|xls)$/i.test(file.name)) {
      setError("Formato não aceito. Selecione um arquivo Excel .xlsx ou .xls.");
      return;
    }
    try {
      setProcessingStep(0);
      const buffer = await file.arrayBuffer();
      await pause(260);
      setProcessingStep(1);
      const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
      const rows = parseWorkbookRows(workbook);
      await pause(260);
      setProcessingStep(2);
      const next = calculateClosing(rows, rate, file.name);
      await pause(260);
      setProcessingStep(3);
      await pause(300);
      setFileName(file.name);
      setResult(next);
    } catch (reason) {
      setResult(null);
      setFileName("");
      setError(reason instanceof Error ? reason.message : "Não foi possível ler o arquivo.");
    } finally {
      setProcessingStep(-1);
    }
  }, [rate]);

  const handleExport = () => {
    if (!result) return;
    exportClosing(result, rate);
    setExportSuccess(true);
    window.setTimeout(() => setExportSuccess(false), 3600);
  };

  const replaceFile = () => {
    setReplaceModal(false);
    setResult(null);
    setFileName("");
    setError("");
    setQuery("");
    setSelected(null);
  };

  const applyRate = () => {
    const parsed = Number(draftRate.replace(/\./g, "").replace(",", "."));
    if (!Number.isFinite(parsed) || parsed < 0) return;
    setRate(parsed);
    localStorage.setItem(RATE_KEY, String(parsed));
    if (result) recalculate(result, parsed);
    setRateModal(false);
  };

  const sortedDrivers = useMemo(() => {
    if (!result) return [];
    const direction = sort.direction === "asc" ? 1 : -1;
    return [...result.drivers]
      .filter((driver) => driver.name.toLocaleLowerCase("pt-BR").includes(query.toLocaleLowerCase("pt-BR")))
      .sort((a, b) => {
        const left = a[sort.key];
        const right = b[sort.key];
        return (typeof left === "string"
          ? left.localeCompare(String(right), "pt-BR")
          : Number(left) - Number(right)) * direction;
      });
  }, [query, result, sort]);

  const changeSort = (key: SortKey) => {
    setSort((current) => ({
      key,
      direction: current.key === key && current.direction === "desc" ? "asc" : "desc",
    }));
  };

  if (!result) {
    return (
      <main className="min-h-screen bg-[#f4f7fb] text-[#10223e]">
        <Topbar rate={rate} onEdit={() => { setDraftRate(rate.toLocaleString("pt-BR", { minimumFractionDigits: 2 })); setRateModal(true); }} />
        <section className="mx-auto flex min-h-[calc(100vh-82px)] max-w-6xl items-center px-5 py-12">
          <div className="grid w-full items-center gap-12 lg:grid-cols-[0.8fr_1.2fr]">
            <div>
              <span className="eyebrow"><span className="status-dot" /> Fechamento de rotas</span>
              <h1 className="mt-6 text-4xl font-semibold tracking-[-0.04em] text-[#0b1d36] sm:text-5xl">NTS Rotas</h1>
              <h2 className="mt-3 text-xl font-semibold text-[#28527f] sm:text-2xl">Gestão Inteligente de Quilometragem</h2>
              <p className="mt-5 max-w-xl text-base leading-7 text-[#607089] sm:text-lg sm:leading-8">
                Importe o relatório exportado do Spoke para calcular automaticamente a quilometragem válida, o bônus dos motoboys e gerar o fechamento do período em poucos segundos.
              </p>
              <div className="mt-7 grid gap-3 text-sm text-[#53647d] sm:grid-cols-3 lg:grid-cols-1 xl:grid-cols-3">
                <span className="feature"><CheckCircle2 size={17} /> Processamento automático</span>
                <span className="feature"><Gauge size={17} /> Cálculo de bônus</span>
                <span className="feature"><FileCheck2 size={17} /> Auditoria automática</span>
              </div>
              <LogisticsIllustration />
            </div>
            <div
              className={`upload-zone ${dragging ? "is-dragging" : ""}`}
              onDragOver={(event) => { event.preventDefault(); setDragging(true); }}
              onDragLeave={() => setDragging(false)}
              onDrop={(event) => { event.preventDefault(); setDragging(false); void readFile(event.dataTransfer.files[0]); }}
            >
              <input ref={inputRef} className="hidden" type="file" accept=".xlsx,.xls" onChange={(e) => void readFile(e.target.files?.[0])} />
              <div className="upload-icon"><Upload size={28} /></div>
              <h2 className="mt-6 text-xl font-semibold">Importe o fechamento do Spoke</h2>
              <p className="mt-2 text-sm leading-6 text-[#6b7a90]">Arraste o arquivo para esta área ou selecione no computador.<br />Formatos aceitos: .xlsx e .xls</p>
              <button className="primary-btn mt-7" onClick={() => inputRef.current?.click()}>
                <FileSpreadsheet size={18} /> Selecionar Excel
              </button>
              <p className="mt-5 flex items-center justify-center gap-2 text-xs text-[#8b98aa]"><Info size={14} /> Seus dados não saem deste navegador.</p>
              {error && <div className="error-box mt-5" role="alert">{error}</div>}
            </div>
          </div>
        </section>
        {rateModal && <RateModal value={draftRate} onChange={setDraftRate} onCancel={() => setRateModal(false)} onApply={applyRate} />}
        {processingStep >= 0 && <ProcessingOverlay step={processingStep} />}
      </main>
    );
  }

  const audit = result.audit;
  return (
    <main className="min-h-screen bg-[#f4f7fb] text-[#10223e]">
      <Topbar rate={rate} onEdit={() => { setDraftRate(rate.toLocaleString("pt-BR", { minimumFractionDigits: 2 })); setRateModal(true); }} />
      <div className="mx-auto max-w-[1480px] px-4 py-8 sm:px-6 lg:px-8">
        <div className="mb-7 flex flex-col justify-between gap-4 md:flex-row md:items-end">
          <div>
            <span className="eyebrow"><span className="status-dot" /> Arquivo processado</span>
            <h1 className="mt-3 text-3xl font-semibold tracking-[-0.035em]">Fechamento de rotas</h1>
            <p className="mt-2 flex flex-wrap items-center gap-2 text-sm text-[#6a7890]">
              <FileSpreadsheet size={15} /> {fileName} <span>•</span> {result.totalRows.toLocaleString("pt-BR")} registros
              <span className="schema-badge">{result.schemaLabel}</span>
            </p>
          </div>
          <div className="flex flex-wrap gap-3">
            <button className="secondary-btn" onClick={() => setReplaceModal(true)}>
              <Upload size={17} /> Trocar arquivo
            </button>
            <button className="primary-btn" onClick={handleExport}>
              <Download size={17} /> Exportar fechamento
            </button>
          </div>
        </div>

        <section className="metrics-grid">
          <Metric icon={<Route />} label="Quilometragem válida" value={formatKm(result.totalKm)} helper={`${audit.included.toLocaleString("pt-BR")} trechos incluídos`} />
          <Metric icon={<span className="currency-icon">R$</span>} label="Bônus total" value={formatBRL(result.totalBonus)} helper={`${formatBRL(rate)} por km`} accent />
          <Metric icon={<Users />} label="Motoboys" value={String(result.drivers.length)} helper="com km válida" />
          <Metric icon={<Bike />} label="Entregas concluídas" value={result.totalDeliveries.toLocaleString("pt-BR")} helper="paradas entregues" />
          <Metric icon={<CalendarDays />} label="Período analisado" value={result.periodLabel} helper={`${result.workDays} dias com rota`} wide />
        </section>

        <section className="mt-6 grid gap-6 xl:grid-cols-2">
          <ChartCard title="Ranking de quilômetros" subtitle="Comparativo por motoboy">
            <ResponsiveContainer width="100%" height={290}>
              <BarChart data={result.drivers} layout="vertical" margin={{ left: 8, right: 25 }}>
                <CartesianGrid stroke="#e8edf4" horizontal={false} />
                <XAxis type="number" tickFormatter={(v) => `${v} km`} tick={{ fill: "#75839a", fontSize: 12 }} axisLine={false} tickLine={false} />
                <YAxis type="category" dataKey="name" width={140} tick={{ fill: "#34445e", fontSize: 12 }} axisLine={false} tickLine={false} />
                <Tooltip
                  formatter={(value) => [formatKm(Number(value)), "Quilometragem válida"]}
                  labelFormatter={(label) => `Motoboy: ${label}`}
                  cursor={{ fill: "#f4f7fb" }}
                />
                <Bar dataKey="validKm" fill="#1267d6" radius={[0, 6, 6, 0]} barSize={22} />
              </BarChart>
            </ResponsiveContainer>
          </ChartCard>
          <ChartCard title="Evolução diária" subtitle="Quilometragem válida de todas as rotas">
            <ResponsiveContainer width="100%" height={290}>
              <AreaChart data={result.daily}>
                <defs><linearGradient id="kmFill" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stopColor="#2b7de1" stopOpacity={0.28} /><stop offset="100%" stopColor="#2b7de1" stopOpacity={0.02} /></linearGradient></defs>
                <CartesianGrid stroke="#e8edf4" vertical={false} />
                <XAxis dataKey="label" tick={{ fill: "#75839a", fontSize: 12 }} axisLine={false} tickLine={false} />
                <YAxis tickFormatter={(v) => `${v}`} tick={{ fill: "#75839a", fontSize: 12 }} axisLine={false} tickLine={false} width={45} />
                <Tooltip
                  formatter={(value) => [formatKm(Number(value)), "Quilometragem válida"]}
                  labelFormatter={(label) => `Data: ${label}`}
                />
                <Area type="monotone" dataKey="km" stroke="#1267d6" strokeWidth={2.5} fill="url(#kmFill)" />
              </AreaChart>
            </ResponsiveContainer>
          </ChartCard>
        </section>

        <section className="panel mt-6 overflow-hidden">
          <div className="flex flex-col justify-between gap-4 border-b border-[#e6ebf2] p-5 sm:flex-row sm:items-center">
            <div><h2 className="panel-title">Resumo por motoboy</h2><p className="panel-subtitle">Clique em uma linha para abrir o detalhamento.</p></div>
            <label className="search-box"><Search size={16} /><input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Buscar motoboy" /></label>
          </div>
          <div className="overflow-x-auto">
            <table>
              <thead><tr>
                <th>#</th>
                <Sortable label="Motoboy" column="name" sort={sort} onSort={changeSort} />
                <Sortable label="Entregas" column="deliveries" sort={sort} onSort={changeSort} />
                <Sortable label="Dias trabalhados" column="daysWorked" sort={sort} onSort={changeSort} />
                <Sortable label="Km válidos" column="validKm" sort={sort} onSort={changeSort} />
                <Sortable label="Média diária" column="dailyAverage" sort={sort} onSort={changeSort} />
                <th>Valor/km</th>
                <Sortable label="Bônus" column="bonus" sort={sort} onSort={changeSort} />
                <th />
              </tr></thead>
              <tbody>{sortedDrivers.map((driver, index) => (
                <tr key={driver.name} onClick={() => setSelected(driver)}>
                  <td><span className={`rank ${index < 3 ? "top" : ""}`}>{index + 1}</span></td>
                  <td className="font-medium text-[#172943]">{driver.name}</td>
                  <td>{driver.deliveries.toLocaleString("pt-BR")}</td>
                  <td>{driver.daysWorked.toLocaleString("pt-BR")}</td>
                  <td className="font-semibold text-[#172943]">{formatKm(driver.validKm)}</td>
                  <td>{formatKm(driver.dailyAverage)}</td>
                  <td>{formatBRL(rate)}</td>
                  <td className="font-semibold text-[#0c60cb]">{formatBRL(driver.bonus)}</td>
                  <td><ChevronRight size={17} /></td>
                </tr>
              ))}</tbody>
            </table>
          </div>
        </section>

        <section className="panel mt-6">
          <button className="flex w-full items-center justify-between p-5 text-left" onClick={() => setAuditOpen(!auditOpen)}>
            <div className="flex items-center gap-3"><span className="audit-icon"><FileCheck2 size={18} /></span><div><h2 className="panel-title">Auditoria da importação</h2><p className="panel-subtitle">{audit.included} incluídos · {audit.totalExcluded} desconsiderados</p></div></div>
            <ChevronDown className={`transition ${auditOpen ? "rotate-180" : ""}`} size={20} />
          </button>
          {auditOpen && <div className="audit-grid border-t border-[#e6ebf2] p-5">
            <AuditItem label="Registros incluídos" value={audit.included} tone="green" />
            <AuditItem label="Excluídos por status" value={audit.statusExcluded} />
            <AuditItem label="Distância vazia" value={audit.emptyDistance} />
            <AuditItem label="Distância inválida" value={audit.invalidDistance} />
            <AuditItem label="Distância negativa" value={audit.negativeDistance} />
            <AuditItem label="Registros duplicados" value={audit.duplicateRecords} />
          </div>}
        </section>
      </div>
      {rateModal && <RateModal value={draftRate} onChange={setDraftRate} onCancel={() => setRateModal(false)} onApply={applyRate} />}
      {replaceModal && <ReplaceModal onCancel={() => setReplaceModal(false)} onConfirm={replaceFile} />}
      {selected && <DriverDrawer driver={selected} rate={rate} records={result.includedRecords.filter((r) => r.driver === selected.name)} onClose={() => setSelected(null)} />}
      {exportSuccess && <div className="success-toast" role="status"><CheckCircle2 size={20} /><span>Fechamento exportado com sucesso.</span></div>}
    </main>
  );
}

function Topbar({ rate, onEdit }: { rate: number; onEdit: () => void }) {
  return <header className="topbar"><div className="mx-auto flex h-[82px] max-w-[1480px] items-center justify-between px-5 lg:px-8">
    <div className="flex items-center gap-3"><div className="brand-mark"><Route size={24} /></div><div><div className="brand-name">NTS <span>Rotas</span></div><div className="brand-sub">Gestão de quilometragem</div></div></div>
    <div className="topbar-actions"><div className="release-meta"><strong>Versão 1.0.0</strong><span>Atualizado em 27/07/2026</span></div><div className="rate-pill"><div><span>Valor vigente</span><strong>{formatBRL(rate)} <small>por km</small></strong></div><button onClick={onEdit} aria-label="Editar valor por quilômetro"><Pencil size={16} /><span className="hidden sm:inline">Editar valor</span></button></div></div>
  </div></header>;
}

function Metric({ icon, label, value, helper, accent = false, wide = false }: { icon: React.ReactNode; label: string; value: string; helper: string; accent?: boolean; wide?: boolean }) {
  return <div className={`metric-card ${accent ? "accent" : ""} ${wide ? "wide" : ""}`}><div className="metric-icon">{icon}</div><div><p>{label}</p><strong>{value}</strong><span>{helper}</span></div></div>;
}

function ChartCard({ title, subtitle, children }: { title: string; subtitle: string; children: React.ReactNode }) {
  return <div className="panel p-5"><h2 className="panel-title">{title}</h2><p className="panel-subtitle mb-6">{subtitle}</p>{children}</div>;
}

function Sortable({ label, column, sort, onSort }: { label: string; column: SortKey; sort: { key: SortKey; direction: string }; onSort: (key: SortKey) => void }) {
  return <th><button className="sort-btn" onClick={() => onSort(column)}>{label}<ChevronDown size={13} className={sort.key === column && sort.direction === "asc" ? "rotate-180" : ""} /></button></th>;
}

function AuditItem({ label, value, tone }: { label: string; value: number; tone?: "green" }) {
  return <div className="audit-item"><span className={tone === "green" ? "success-dot" : "neutral-dot"} /> <div><strong>{value.toLocaleString("pt-BR")}</strong><span>{label}</span></div></div>;
}

function LogisticsIllustration() {
  return <div className="logistics-illustration" aria-hidden="true">
    <span className="route-line" />
    <span className="illustration-node node-start"><FileSpreadsheet size={18} /></span>
    <span className="illustration-node node-middle"><Bike size={23} /></span>
    <span className="illustration-node node-end"><CheckCircle2 size={19} /></span>
    <span className="illustration-label label-start">Relatório</span>
    <span className="illustration-label label-middle">Rotas</span>
    <span className="illustration-label label-end">Fechamento</span>
  </div>;
}

function ProcessingOverlay({ step }: { step: number }) {
  return <div className="processing-backdrop" role="status" aria-live="polite">
    <div className="processing-card">
      <div className="processing-spinner"><Route size={24} /></div>
      <h2>Processando relatório...</h2>
      <p>{PROCESSING_STEPS[step]}</p>
      <div className="processing-progress"><span style={{ width: `${((step + 1) / PROCESSING_STEPS.length) * 100}%` }} /></div>
      <div className="processing-steps">
        {PROCESSING_STEPS.map((label, index) => <span key={label} className={index <= step ? "active" : ""}><CheckCircle2 size={14} />{label}</span>)}
      </div>
    </div>
  </div>;
}

function ReplaceModal({ onCancel, onConfirm }: { onCancel: () => void; onConfirm: () => void }) {
  return <div className="modal-backdrop" onMouseDown={onCancel}><div className="modal-card" onMouseDown={(event) => event.stopPropagation()} role="dialog" aria-modal="true" aria-labelledby="replace-title">
    <div className="confirm-icon"><Upload size={21} /></div>
    <h2 id="replace-title" className="mt-4">Deseja substituir o relatório atual?</h2>
    <p>O dashboard atual será fechado para que você selecione um novo arquivo.</p>
    <div className="mt-6 flex justify-end gap-3"><button className="secondary-btn" onClick={onCancel}>Cancelar</button><button className="primary-btn" onClick={onConfirm}>Substituir</button></div>
  </div></div>;
}

function RateModal({ value, onChange, onCancel, onApply }: { value: string; onChange: (value: string) => void; onCancel: () => void; onApply: () => void }) {
  return <div className="modal-backdrop" onMouseDown={onCancel}><div className="modal-card" onMouseDown={(e) => e.stopPropagation()} role="dialog" aria-modal="true">
    <div className="flex items-start justify-between"><div><h2>Editar valor por quilômetro</h2><p>O bônus de todos os motoboys será recalculado.</p></div><button className="icon-btn" onClick={onCancel} aria-label="Fechar edição"><X size={19} /></button></div>
    <label className="money-field"><span>Valor por km</span><div><b>R$</b><input autoFocus inputMode="decimal" value={value} onChange={(e) => onChange(e.target.value)} /></div></label>
    <div className="mt-6 flex justify-end gap-3"><button className="secondary-btn" onClick={onCancel}>Cancelar</button><button className="primary-btn" onClick={onApply}>Aplicar</button></div>
  </div></div>;
}

function DriverDrawer({ driver, rate, records, onClose }: { driver: DriverSummary; rate: number; records: ProcessedRecord[]; onClose: () => void }) {
  return <div className="drawer-backdrop" onMouseDown={onClose}><aside className="drawer" onMouseDown={(e) => e.stopPropagation()}>
    <div className="drawer-header"><div><span>Detalhamento do motoboy</span><h2>{driver.name}</h2></div><button className="icon-btn" onClick={onClose} aria-label="Fechar detalhamento"><X size={20} /></button></div>
    <div className="drawer-body">
      <div className="drawer-metrics"><div><span>Km válidos</span><strong>{formatKm(driver.validKm)}</strong></div><div className="blue"><span>Bônus</span><strong>{formatBRL(driver.bonus)}</strong></div><div><span>Entregas</span><strong>{driver.deliveries}</strong></div><div><span>Dias trabalhados</span><strong>{driver.daysWorked}</strong></div><div><span>Média diária</span><strong>{formatKm(driver.dailyAverage)}</strong></div><div><span>Valor/km</span><strong>{formatBRL(rate)}</strong></div></div>
      <h3>Evolução diária</h3>
      <div className="h-56"><ResponsiveContainer width="100%" height="100%"><AreaChart data={driver.daily}><CartesianGrid stroke="#e8edf4" vertical={false} /><XAxis dataKey="label" tick={{ fontSize: 11 }} axisLine={false} tickLine={false} /><YAxis tick={{ fontSize: 11 }} axisLine={false} tickLine={false} width={38} /><Tooltip formatter={(v) => [formatKm(Number(v)), "Quilometragem válida"]} labelFormatter={(label) => `Data: ${label}`} /><Area type="monotone" dataKey="km" stroke="#1267d6" strokeWidth={2} fill="#dcecff" /></AreaChart></ResponsiveContainer></div>
      <h3 className="mt-7">Registros considerados</h3>
      <div className="drawer-table"><table><thead><tr><th>Data</th><th>Rota</th><th>Parada</th><th>Km</th></tr></thead><tbody>{records.map((record) => <tr key={record.rowNumber}><td>{formatDateBR(record.date)}</td><td>{record.route || "—"}</td><td>{record.stopNumber ?? "Trecho de base"}</td><td>{formatKm(record.distance)}</td></tr>)}</tbody></table></div>
    </div>
  </aside></div>;
}
