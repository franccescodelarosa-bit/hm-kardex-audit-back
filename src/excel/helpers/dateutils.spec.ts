
import { DateUtils } from "./dateutils";

describe("DateUtils.monthName", () => {
    it("devuelve el nombre del mes para valores 1-12", () => {
        expect(DateUtils.monthName(1)).toBe("Enero");
        expect(DateUtils.monthName(12)).toBe("Diciembre");
    });

    it("para un número fuera de rango (que no sea un mes real) cae al toString, como antes", () => {
        expect(DateUtils.monthName(13)).toBe("13");
        expect(DateUtils.monthName(0)).toBe("0");
    });

    it("NO explota con undefined -- devuelve 'Sin período' (reproduce el crash real de RULE_008: 'Cannot read properties of undefined (reading toString)')", () => {
        expect(() => DateUtils.monthName(undefined as unknown as number)).not.toThrow();
        expect(DateUtils.monthName(undefined as unknown as number)).toBe("Sin período");
    });

    it("tampoco explota con null", () => {
        expect(DateUtils.monthName(null as unknown as number)).toBe("Sin período");
    });
});
