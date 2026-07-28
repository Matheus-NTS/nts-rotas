import { describe, expect, it } from "vitest";
import * as XLSX from "xlsx";
import { buildClosingWorkbook, calculateClosing, getDriverRateKey } from "./closing";

const base = {
  driver: "Ana", route: "Rota 1", date_attempted_local: "07/15/2026",
  stop_number: 1, tracking_code: "ABC",
};

describe("calculateClosing", () => {
  it("inclui apenas status completed e calcula bônus", () => {
    const result = calculateClosing([
      { ...base, status: "completed", distance_km: 10 },
      { ...base, tracking_code: "DEF", status: "cancelled", distance_km: 99 },
    ], 0.25);
    expect(result.totalKm).toBe(10);
    expect(result.totalBonus).toBe(2.5);
    expect(result.audit.statusExcluded).toBe(1);
  });

  it("aceita somente stop_state exatamente igual a delivered", () => {
    const result = calculateClosing([
      { ...base, stop_state: "delivered", distance_km: 12.603 },
      { ...base, tracking_code: "DEF", stop_state: "out_for_delivery", distance_km: 20 },
      { ...base, tracking_code: "GHI", stop_state: "failed", distance_km: 30 },
      { ...base, tracking_code: "JKL", stop_state: "Delivered", distance_km: 40 },
    ], 0.25);
    expect(result.totalKm).toBe(12.603);
    expect(result.audit.statusExcluded).toBe(3);
    expect(result.schemaLabel).toContain("stop_state");
  });

  it("audita distância vazia, inválida e negativa", () => {
    const result = calculateClosing([
      { ...base, status: "completed", distance_km: null },
      { ...base, status: "completed", distance_km: "abc" },
      { ...base, status: "completed", distance_km: -1 },
    ], 0.25);
    expect(result.audit).toMatchObject({ emptyDistance: 1, invalidDistance: 1, negativeDistance: 1, included: 0 });
  });

  it("soma trechos individuais da rota, inclusive retorno à base", () => {
    const result = calculateClosing([
      { ...base, stop_state: "delivered", distance_km: 6.5 },
      { ...base, stop_number: null, tracking_code: "", stop_state: "delivered", distance_km: 4.25 },
    ], 0.5);
    expect(result.totalKm).toBe(10.75);
    expect(result.drivers[0].deliveries).toBe(1);
    expect(result.totalBonus).toBe(5.38);
  });

  it("recalcula todos os bônus quando o valor por km muda", () => {
    const rows = [
      { ...base, driver: "Ana", stop_state: "delivered", distance_km: 10 },
      { ...base, driver: "Bruno", tracking_code: "DEF", stop_state: "delivered", distance_km: 20 },
    ];
    const original = calculateClosing(rows, 0.25);
    const updated = calculateClosing(rows, 0.4);
    expect(original.drivers.map((driver) => driver.bonus)).toEqual([5, 2.5]);
    expect(updated.drivers.map((driver) => driver.bonus)).toEqual([8, 4]);
    expect(updated.totalBonus).toBe(12);
  });

  it("aplica valor personalizado somente ao motoboy correspondente", () => {
    const rows = [
      { ...base, driver: "Ana", stop_state: "delivered", distance_km: 10 },
      { ...base, driver: "Bruno", tracking_code: "DEF", stop_state: "delivered", distance_km: 20 },
    ];
    const overrides = { [getDriverRateKey("Ana")]: 0.5 };
    const result = calculateClosing(rows, 0.25, "fechamento.xlsx", overrides);

    expect(result.drivers.find((driver) => driver.name === "Ana")).toMatchObject({
      appliedRate: 0.5,
      rateType: "Personalizado",
      bonus: 5,
    });
    expect(result.drivers.find((driver) => driver.name === "Bruno")).toMatchObject({
      appliedRate: 0.25,
      rateType: "Padrão",
      bonus: 5,
    });
    expect(result.totalBonus).toBe(10);
  });

  it("mantém o valor personalizado ao alterar o valor global", () => {
    const rows = [
      { ...base, driver: "Ana", stop_state: "delivered", distance_km: 10 },
      { ...base, driver: "Bruno", tracking_code: "DEF", stop_state: "delivered", distance_km: 20 },
    ];
    const overrides = { [getDriverRateKey("  ANA ")]: 0.5 };
    const result = calculateClosing(rows, 0.4, "fechamento.xlsx", overrides);

    expect(result.drivers.find((driver) => driver.name === "Ana")?.bonus).toBe(5);
    expect(result.drivers.find((driver) => driver.name === "Bruno")?.bonus).toBe(8);
    expect(result.totalBonus).toBe(13);
  });

  it("reconcilia os totais individuais com o total geral", () => {
    const result = calculateClosing([
      { ...base, driver: "Ana", stop_state: "delivered", distance_km: 1.005 },
      { ...base, driver: "Bruno", tracking_code: "DEF", stop_state: "delivered", distance_km: 1.005 },
    ], 0.25);
    expect(result.drivers.reduce((sum, driver) => sum + driver.validKm, 0)).toBe(result.totalKm);
    expect(result.drivers.reduce((sum, driver) => sum + driver.bonus, 0)).toBe(result.totalBonus);
  });

  it("exclui duplicidades pelo tracking_code e pela identidade do trecho de base", () => {
    const delivered = { ...base, stop_state: "delivered", distance_km: 10 };
    const baseLeg = { ...base, stop_number: null, tracking_code: "", stop_state: "delivered", distance_km: 4, address: "Base NTS" };
    const result = calculateClosing([delivered, { ...delivered }, baseLeg, { ...baseLeg }], 0.25);
    expect(result.totalKm).toBe(14);
    expect(result.audit.duplicateRecords).toBe(2);
    expect(result.excludedRecords.filter((record) => record.exclusionReason === "Registro duplicado")).toHaveLength(2);
  });

  it("unifica nomes com espaços, caixa e acentuação equivalentes", () => {
    const result = calculateClosing([
      { ...base, driver: "  José   da Silva ", stop_state: "delivered", distance_km: 5 },
      { ...base, driver: "jose da silva", tracking_code: "DEF", stop_state: "delivered", distance_km: 7 },
    ], 0.25);
    expect(result.drivers).toHaveLength(1);
    expect(result.drivers[0].name).toBe("José da Silva");
    expect(result.drivers[0].validKm).toBe(12);
  });

  it("gera o Excel com quatro abas, filtros e formatos definidos", () => {
    const result = calculateClosing(
      [{ ...base, stop_state: "delivered", distance_km: 12.6 }],
      0.25,
      "fechamento.xlsx",
      { [getDriverRateKey("Ana")]: 0.5 },
    );
    const workbook = buildClosingWorkbook(result, 0.25);
    expect(workbook.SheetNames).toEqual([
      "Resumo por Motoboy",
      "Detalhamento Diário",
      "Registros Considerados",
      "Registros Desconsiderados",
    ]);
    const summary = workbook.Sheets["Resumo por Motoboy"];
    expect(summary["!autofilter"]).toBeTruthy();
    expect(summary.G2.z).toBe('"R$" #,##0.00');
    expect(summary.I2.z).toBe('"R$" #,##0.00');
    expect(summary.E2.z).toBe("#,##0.00");
    expect(XLSX.utils.sheet_to_json<Record<string, unknown>>(summary)).toEqual([
      expect.objectContaining({
        "Valor/km aplicado": 0.5,
        "Tipo de valor": "Personalizado",
      }),
    ]);
    expect(XLSX.utils.sheet_to_json<Record<string, unknown>>(workbook.Sheets["Detalhamento Diário"])[0]).toMatchObject({
      "Valor/km aplicado": 0.5,
      "Tipo de valor": "Personalizado",
      Bônus: 6.3,
    });
    expect(XLSX.utils.sheet_to_json<Record<string, unknown>>(workbook.Sheets["Registros Considerados"])[0]).toMatchObject({
      "Valor/km aplicado": 0.5,
      "Tipo de valor": "Personalizado",
    });
  });
});
