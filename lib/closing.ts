import * as XLSX from "xlsx";

export type RawRow = Record<string, unknown>;

export type ProcessedRecord = {
  rowNumber: number;
  driver: string;
  status: string;
  distance: number;
  date: Date | null;
  dateKey: string;
  route: string;
  stopNumber: string | number | null;
  trackingCode: string;
  original: RawRow;
};

export type DriverSummary = {
  name: string;
  deliveries: number;
  daysWorked: number;
  validKm: number;
  dailyAverage: number;
  bonus: number;
  daily: { dateKey: string; label: string; km: number }[];
};

export type ClosingResult = {
  fileName: string;
  sourceRows: RawRow[];
  totalRows: number;
  totalKm: number;
  totalBonus: number;
  totalDeliveries: number;
  workDays: number;
  periodLabel: string;
  schemaLabel: string;
  drivers: DriverSummary[];
  daily: { dateKey: string; label: string; km: number }[];
  includedRecords: ProcessedRecord[];
  excludedRecords: (ProcessedRecord & { exclusionReason: string })[];
  audit: {
    included: number;
    totalExcluded: number;
    statusExcluded: number;
    emptyDistance: number;
    invalidDistance: number;
    negativeDistance: number;
  };
};

const DATE_COLUMNS = ["date_attempted_local", "date_departed_local", "date_arrived_local", "date"];

export function parseWorkbookRows(workbook: XLSX.WorkBook): RawRow[] {
  const firstSheet = workbook.SheetNames[0];
  if (!firstSheet) throw new Error("O Excel não possui nenhuma aba.");
  return XLSX.utils.sheet_to_json<RawRow>(workbook.Sheets[firstSheet], { defval: null, raw: true });
}

function asText(value: unknown): string {
  return value === null || value === undefined ? "" : String(value).trim();
}

function parseDistance(value: unknown): number | null {
  if (typeof value === "number") return Number.isFinite(value) ? value : null;
  const text = asText(value);
  if (!text) return null;
  const normalized = text.includes(",") ? text.replace(/\./g, "").replace(",", ".") : text;
  const number = Number(normalized);
  return Number.isFinite(number) ? number : Number.NaN;
}

function parseDate(value: unknown): Date | null {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  if (typeof value === "number" && value > 20000) {
    const parsed = XLSX.SSF.parse_date_code(value);
    return parsed ? new Date(parsed.y, parsed.m - 1, parsed.d) : null;
  }
  const text = asText(value);
  if (!text) return null;
  const slash = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (slash) {
    // O Spoke exporta as datas locais em MM/DD/YYYY, como confirmado no arquivo de referência.
    return new Date(Number(slash[3]), Number(slash[1]) - 1, Number(slash[2]));
  }
  const date = new Date(text);
  return Number.isNaN(date.getTime()) ? null : date;
}

function dateKey(date: Date | null) {
  if (!date) return "";
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function resolveSchema(rows: RawRow[]) {
  const headers = new Set(Object.keys(rows[0] ?? {}));
  const missing = ["distance_km", "driver"].filter((column) => !headers.has(column));
  const statusColumn = headers.has("status") ? "status" : headers.has("stop_state") ? "stop_state" : "";
  if (!statusColumn) missing.push("status (ou stop_state no export atual do Spoke)");
  if (missing.length) throw new Error(`Arquivo inválido. Colunas obrigatórias ausentes: ${missing.join(", ")}.`);
  const dateColumn = DATE_COLUMNS.find((column) => headers.has(column)) ?? "";
  return {
    statusColumn,
    acceptedStatus: statusColumn === "status" ? "completed" : "delivered",
    dateColumn,
    label: statusColumn === "status" ? "status: completed" : "Spoke atual: stop_state = delivered",
  };
}

export function calculateClosing(rows: RawRow[], rate: number, fileName = "fechamento.xlsx"): ClosingResult {
  if (!rows.length) throw new Error("O Excel está vazio.");
  const schema = resolveSchema(rows);
  const audit = { included: 0, totalExcluded: 0, statusExcluded: 0, emptyDistance: 0, invalidDistance: 0, negativeDistance: 0 };
  const includedRecords: ProcessedRecord[] = [];
  const excludedRecords: (ProcessedRecord & { exclusionReason: string })[] = [];

  rows.forEach((row, index) => {
    const status = asText(row[schema.statusColumn]);
    const distanceValue = row.distance_km;
    const distance = parseDistance(distanceValue);
    const date = parseDate(schema.dateColumn ? row[schema.dateColumn] : null);
    const record: ProcessedRecord = {
      rowNumber: index + 2,
      driver: asText(row.driver) || "Motoboy não informado",
      status,
      distance: typeof distance === "number" && Number.isFinite(distance) ? distance : 0,
      date,
      dateKey: dateKey(date),
      route: asText(row.route),
      stopNumber: row.stop_number === null || row.stop_number === undefined || row.stop_number === "" ? null : (row.stop_number as string | number),
      trackingCode: asText(row.tracking_code),
      original: row,
    };
    let reason = "";
    if (status !== schema.acceptedStatus) {
      audit.statusExcluded++;
      reason = `Status diferente de ${schema.acceptedStatus}`;
    } else if (distanceValue === null || distanceValue === undefined || asText(distanceValue) === "") {
      audit.emptyDistance++;
      reason = "Distância vazia";
    } else if (distance === null || Number.isNaN(distance)) {
      audit.invalidDistance++;
      reason = "Distância inválida";
    } else if (distance < 0) {
      audit.negativeDistance++;
      reason = "Distância negativa";
    }
    if (reason) excludedRecords.push({ ...record, exclusionReason: reason });
    else includedRecords.push(record);
  });

  audit.included = includedRecords.length;
  audit.totalExcluded = excludedRecords.length;
  const byDriver = new Map<string, ProcessedRecord[]>();
  includedRecords.forEach((record) => byDriver.set(record.driver, [...(byDriver.get(record.driver) ?? []), record]));

  const drivers = [...byDriver.entries()].map(([name, records]) => {
    const days = new Map<string, number>();
    records.forEach((record) => {
      if (record.dateKey) days.set(record.dateKey, (days.get(record.dateKey) ?? 0) + record.distance);
    });
    const daily = [...days.entries()].sort(([a], [b]) => a.localeCompare(b)).map(([key, km]) => ({
      dateKey: key, label: formatShortDate(key), km: round(km),
    }));
    const validKm = round(records.reduce((sum, record) => sum + record.distance, 0));
    const deliveries = new Set(records.filter((record) => record.trackingCode).map((record) => record.trackingCode)).size;
    return {
      name,
      deliveries,
      daysWorked: days.size,
      validKm,
      dailyAverage: round(days.size ? validKm / days.size : 0),
      bonus: roundMoney(validKm * rate),
      daily,
    };
  }).sort((a, b) => b.validKm - a.validKm);

  const allDays = new Map<string, number>();
  includedRecords.forEach((record) => {
    if (record.dateKey) allDays.set(record.dateKey, (allDays.get(record.dateKey) ?? 0) + record.distance);
  });
  const daily = [...allDays.entries()].sort(([a], [b]) => a.localeCompare(b)).map(([key, km]) => ({
    dateKey: key, label: formatShortDate(key), km: round(km),
  }));
  const dates = includedRecords.map((record) => record.date).filter((date): date is Date => Boolean(date)).sort((a, b) => a.getTime() - b.getTime());
  const totalKm = round(includedRecords.reduce((sum, record) => sum + record.distance, 0));

  return {
    fileName,
    sourceRows: rows,
    totalRows: rows.length,
    totalKm,
    totalBonus: roundMoney(totalKm * rate),
    totalDeliveries: drivers.reduce((sum, driver) => sum + driver.deliveries, 0),
    workDays: allDays.size,
    periodLabel: dates.length ? `${formatDateBR(dates[0])} – ${formatDateBR(dates[dates.length - 1])}` : "Período não identificado",
    schemaLabel: schema.label,
    drivers,
    daily,
    includedRecords,
    excludedRecords,
    audit,
  };
}

function round(value: number) { return Math.round((value + Number.EPSILON) * 1000) / 1000; }
function roundMoney(value: number) { return Math.round((value + Number.EPSILON) * 100) / 100; }
function formatShortDate(key: string) {
  const [, month, day] = key.split("-");
  return `${day}/${month}`;
}
export function formatDateBR(date: Date | null) { return date ? new Intl.DateTimeFormat("pt-BR").format(date) : "Sem data"; }
export function formatKm(value: number) { return `${value.toLocaleString("pt-BR", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} km`; }
export function formatBRL(value: number) { return value.toLocaleString("pt-BR", { style: "currency", currency: "BRL" }); }

export function exportClosing(result: ClosingResult, rate: number) {
  const workbook = XLSX.utils.book_new();
  const summary = result.drivers.map((driver, index) => ({
    Posição: index + 1, Motoboy: driver.name, "Entregas concluídas": driver.deliveries,
    "Dias trabalhados": driver.daysWorked, "Quilômetros válidos": driver.validKm,
    "Média diária": driver.dailyAverage, "Valor por quilômetro": rate, Bônus: driver.bonus,
  }));
  const daily = result.drivers.flatMap((driver) => driver.daily.map((day) => ({
    Data: day.dateKey ? new Date(`${day.dateKey}T12:00:00`) : "", Motoboy: driver.name,
    "Quilômetros válidos": day.km, "Valor por quilômetro": rate, Bônus: roundMoney(day.km * rate),
  })));
  const considered = result.includedRecords.map((record) => ({
    Linha: record.rowNumber, Data: record.date ?? "", Motoboy: record.driver, Rota: record.route,
    Parada: record.stopNumber ?? "Trecho de base", Status: record.status,
    "Distância (km)": record.distance, "Código de rastreio": record.trackingCode,
  }));
  const disregarded = result.excludedRecords.map((record) => ({
    Linha: record.rowNumber, Data: record.date ?? "", Motoboy: record.driver, Rota: record.route,
    Status: record.status, "Distância original": record.original.distance_km ?? "", Motivo: record.exclusionReason,
  }));
  [
    ["Resumo por Motoboy", summary],
    ["Detalhamento Diário", daily],
    ["Registros Considerados", considered],
    ["Registros Desconsiderados", disregarded],
  ].forEach(([name, data]) => {
    const worksheet = XLSX.utils.json_to_sheet(data as object[]);
    worksheet["!autofilter"] = { ref: worksheet["!ref"] ?? "A1:A1" };
    worksheet["!freeze"] = { xSplit: 0, ySplit: 1 };
    worksheet["!cols"] = Object.keys((data as object[])[0] ?? { Dados: "" }).map((header) => ({ wch: Math.min(34, Math.max(12, header.length + 3)) }));
    Object.keys(worksheet).forEach((address) => {
      if (address.startsWith("!")) return;
      const cell = worksheet[address];
      if (cell.t === "d") cell.z = "dd/mm/yyyy";
      const header = worksheet[`${address.replace(/\d+/, "")}1`]?.v;
      if (header === "Bônus" || header === "Valor por quilômetro") cell.z = '"R$" #,##0.00';
      if (String(header).includes("Quilômetros") || header === "Média diária" || header === "Distância (km)") cell.z = '#,##0.000';
    });
    XLSX.utils.book_append_sheet(workbook, worksheet, name as string);
  });
  XLSX.writeFile(workbook, `NTS_Rotas_${new Date().toISOString().slice(0, 10)}.xlsx`, { cellDates: true });
}
