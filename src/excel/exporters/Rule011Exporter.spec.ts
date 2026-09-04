import { Rule011Exporter } from "./Rule011Exporter";
import { ReportHeader } from "./base/ReportHeader";

const header: ReportHeader = {
    companyName: "COMERCIAL L&M EIRL",
    ruc: "20451412508",
    year: 2024
};

describe("Rule011Exporter", () => {
    it("la trazabilidad usa 'Costo Unitario Final/Inicial' (misma convención que RULE_003), no 'Costo Anterior/Actual'", async () => {
        const exporter = new Rule011Exporter();
        const results = [
            {
                product_code: "000002",
                product_name: "CUADERNO BLOCK ANILLADO ESTAMPADO M/L&Q",
                risk_level: "MEDIO",
                metadata: {
                    date: "2024-02-01",
                    month: 2,
                    document: "00 Saldo Inicial",
                    operation: "16",
                    previousCost: 5,
                    currentCost: 0,
                    variationPercent: 100
                }
            }
        ];

        const workbook = await exporter.export(results, header);
        const sheet = workbook.worksheets[0];
        const trace = String(sheet.getRow(5).getCell(10).value);

        expect(trace).toContain("Costo Unitario Final: 5");
        expect(trace).toContain("Costo Unitario Inicial: 0");
        expect(trace).not.toContain("Costo Anterior");
        expect(trace).not.toContain("Costo Actual");
    });

    it("la 'Ocurrencia' junta Documento + Operación en una sola línea legible", async () => {
        const exporter = new Rule011Exporter();
        const results = [
            {
                product_code: "000057",
                product_name: "032-BOLIGRAFO 50-1 COLOR AZUL/ROJO/NEGRO M/FABER CASTELL",
                risk_level: "MEDIO",
                metadata: {
                    date: "2024-02-01",
                    month: 2,
                    document: "00 Saldo Inicial",
                    operation: "16",
                    previousCost: 48,
                    currentCost: 20.98,
                    variationPercent: 56.29
                }
            }
        ];

        const workbook = await exporter.export(results, header);
        const sheet = workbook.worksheets[0];
        const trace = String(sheet.getRow(5).getCell(10).value);

        expect(trace).toContain("Ocurrencia: 00 Saldo Inicial (Operación 16)");
        expect(trace).not.toContain("Documento:");
        expect(trace).not.toContain("Operación:");
    });
});
