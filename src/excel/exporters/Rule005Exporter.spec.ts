
import { Rule005Exporter } from "./Rule005Exporter";
import { ReportHeader } from "./base/ReportHeader";

const header: ReportHeader = {
    companyName: "COMERCIAL L&M EIRL",
    ruc: "20451412508",
    year: 2024
};

describe("Rule005Exporter", () => {
    it("reproduce el caso real: Costo Unitario y Costo Total negativos en el Saldo -> 2 filas, cada una con su propio valor y su propia trazabilidad", async () => {
        const exporter = new Rule005Exporter();
        const results = [
            {
                product_code: "006749",
                product_name: "PEGAMENTO E/LATA DE 1/32 T/MED.M/AFRICANO",
                risk_level: "CRITICO",
                metadata: {
                    source: "KARDEX",
                    month: 1,
                    document: "00 Saldo Inicial",
                    operation: "16",
                    date: "2024-01-01T00:00:00.000Z",
                    balanceQuantity: 10,
                    balanceUnitCost: -4.06,
                    balanceTotalCost: -40.6,
                    negatives: ["Costo Unitario de Saldo", "Costo Total de Saldo"]
                }
            }
        ];

        const workbook = await exporter.export(results, header);
        const sheet = workbook.worksheets[0];

        expect(sheet.rowCount).toBe(6); // 4 de header + 2 filas de datos

        // Fila 1: Costo Unitario de Saldo -- "Valor encontrado" YA NO debe salir vacío
        expect(sheet.getRow(5).getCell(4).value).toBe("Saldo Negativo (Costo Unitario de Saldo)");
        expect(sheet.getRow(5).getCell(6).value).toBe(-4.06);
        const trace1 = String(sheet.getRow(5).getCell(10).value);
        expect(trace1).toContain("Campos Negativos: Costo Unitario de Saldo");
        expect(trace1).not.toContain("Costo Total de Saldo");

        // Fila 2: Costo Total de Saldo -- su propio valor, su propia trazabilidad
        expect(sheet.getRow(6).getCell(4).value).toBe("Saldo Negativo (Costo Total de Saldo)");
        expect(sheet.getRow(6).getCell(6).value).toBe(-40.6);
        const trace2 = String(sheet.getRow(6).getCell(10).value);
        expect(trace2).toContain("Campos Negativos: Costo Total de Saldo");
        expect(trace2).not.toContain("Costo Unitario de Saldo");
    });

    it("sigue reconociendo las etiquetas VIEJAS (auditorias ya corridas antes del fix de rule-005.ts) para no romper reportes existentes", async () => {
        const exporter = new Rule005Exporter();
        const results = [
            {
                product_code: "000144",
                product_name: "ALCOHOL YODADO",
                risk_level: "CRITICO",
                metadata: {
                    source: "KARDEX",
                    month: 3,
                    document: "01 FT001-00027045",
                    operation: "01",
                    date: "2024-03-05T00:00:00.000Z",
                    exitTotalCost: -15,
                    negatives: ["Total Salida"] // etiqueta vieja
                }
            }
        ];

        const workbook = await exporter.export(results, header);
        const sheet = workbook.worksheets[0];

        expect(sheet.rowCount).toBe(5);
        expect(sheet.getRow(5).getCell(6).value).toBe(-15); // "Valor encontrado" resuelto igual
    });
});
