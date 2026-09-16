
import { Rule008Exporter } from "./Rule008Exporter";
import { ReportHeader } from "./base/ReportHeader";

const header: ReportHeader = {
    companyName: "COMERCIAL L&M EIRL",
    ruc: "20451412508",
    year: 2024
};

describe("Rule008Exporter", () => {
    it("BUG REAL reproducido: un hallazgo 'producto en Inventario pero no en Kardex' (rule-008.ts -> validateInventory, que nunca setea metadata.month) NO debe explotar al exportar -- antes tiraba 'Cannot read properties of undefined (reading toString)'", async () => {
        const exporter = new Rule008Exporter();
        const results = [
            {
                product_code: "000999",
                product_name: "PRODUCTO SOLO EN INVENTARIO",
                risk_level: "ALTO",
                metadata: { source: "INVENTARIO_INICIAL" }
            }
        ];

        await expect(exporter.export(results, header)).resolves.toBeDefined();
    });

    it("el hallazgo sin mes se muestra como 'Sin período', no rompe ni queda en blanco", async () => {
        const exporter = new Rule008Exporter();
        const results = [
            {
                product_code: "000999",
                product_name: "PRODUCTO SOLO EN INVENTARIO",
                risk_level: "ALTO",
                metadata: { source: "INVENTARIO_INICIAL" }
            }
        ];

        const workbook = await exporter.export(results, header);
        const sheet = workbook.worksheets[0];

        expect(sheet.getRow(5).getCell(1).value).toBe("Sin período");
    });

    it("el otro camino de RULE_008 (validateKardex) SÍ trae month -- sigue mostrando el mes real, sin regresión", async () => {
        const exporter = new Rule008Exporter();
        const results = [
            {
                product_code: "000888",
                product_name: "PRODUCTO SOLO EN KARDEX",
                risk_level: "ALTO",
                metadata: { source: "Inventario", month: 3, date: "2024-03-15", document: "F001-001" }
            }
        ];

        const workbook = await exporter.export(results, header);
        const sheet = workbook.worksheets[0];

        expect(sheet.getRow(5).getCell(1).value).toBe("Marzo");
    });
});
