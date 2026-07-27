import { describe, expect, it } from "vitest";
import * as XLSX from "xlsx";
import { buildClosingWorkbook, calculateClosing } from "./closing";

const referencePath = process.env.NTS_REFERENCE_XLSX;

describe.runIf(Boolean(referencePath))("homologação com o Excel real", () => {
  it("reconcilia todos os indicadores aprovados", () => {
    const workbook = XLSX.readFile(referencePath!, { cellDates: true });
    const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(
      workbook.Sheets[workbook.SheetNames[0]],
      { defval: null, raw: true },
    );
    const result = calculateClosing(rows, 0.25, referencePath);
    if (process.env.NTS_HOMOLOGATION_EXPORT) {
      XLSX.writeFile(buildClosingWorkbook(result, 0.25), process.env.NTS_HOMOLOGATION_EXPORT, {
        cellDates: true,
      });
    }
    const driverKm = result.drivers.reduce((sum, driver) => sum + driver.validKm, 0);
    const driverBonus = result.drivers.reduce((sum, driver) => sum + driver.bonus, 0);
    const trackingCodes = result.includedRecords.filter((record) => record.trackingCode);
    const baseLegs = result.includedRecords.filter((record) => !record.trackingCode);

    console.info("NTS_HOMOLOGACAO", JSON.stringify({
      colunas: {
        motoboy: "driver",
        entrega: "tracking_code",
        rota: "route",
        data: "date_attempted_local",
        distancia: "distance_km",
        estadoConcluido: "stop_state = delivered",
      },
      totais: {
        registros: result.totalRows,
        quilometragemValida: result.totalKm,
        bonus: result.totalBonus,
        motoboys: result.drivers.length,
        entregas: result.totalDeliveries,
        trechosIncluidos: result.audit.included,
        registrosDesconsiderados: result.audit.totalExcluded,
        trechosComEntrega: trackingCodes.length,
        trechosDeBase: baseLegs.length,
      },
      auditoria: result.audit,
      duplicidades: {
        trackingCodesRepetidos: trackingCodes.length - new Set(trackingCodes.map((record) => record.trackingCode)).size,
        registrosExcluidosComoDuplicados: result.audit.duplicateRecords,
      },
      reconciliacao: {
        somaKmMotoboys: driverKm,
        somaBonusMotoboys: driverBonus,
      },
      porMotoboy: result.drivers.map((driver) => ({
        motoboy: driver.name,
        entregas: driver.deliveries,
        dias: driver.daysWorked,
        km: driver.validKm,
        bonus: driver.bonus,
      })),
    }, null, 2));

    expect(result.totalRows).toBe(799);
    expect(result.totalKm).toBe(10948.515);
    expect(result.totalBonus).toBe(2737.13);
    expect(result.drivers).toHaveLength(5);
    expect(result.totalDeliveries).toBe(352);
    expect(result.audit.included).toBe(567);
    expect(result.audit.totalExcluded).toBe(232);
    expect(result.audit.statusExcluded).toBe(226);
    expect(result.audit.emptyDistance).toBe(6);
    expect(result.audit.invalidDistance).toBe(0);
    expect(result.audit.negativeDistance).toBe(0);
    expect(result.audit.duplicateRecords).toBe(0);
    expect(trackingCodes).toHaveLength(352);
    expect(baseLegs).toHaveLength(215);
    expect(Math.round(driverKm * 1000) / 1000).toBe(result.totalKm);
    expect(driverBonus).toBe(result.totalBonus);
  });
});
