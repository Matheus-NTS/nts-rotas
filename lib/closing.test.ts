import { describe, expect, it } from "vitest";
import { calculateClosing } from "./closing";

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

  it("aceita o schema real do Spoke com stop_state delivered", () => {
    const result = calculateClosing([{ ...base, stop_state: "delivered", distance_km: 12.603 }], 0.25);
    expect(result.totalKm).toBe(12.603);
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
});
