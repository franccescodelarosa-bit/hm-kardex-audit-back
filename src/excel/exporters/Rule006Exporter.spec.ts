/**
 * Reproduce el bug real: la columna "Periodo" del Excel exportado de
 * RULE_006 siempre mostraba "-" fijo, sin importar el mes real del
 * hallazgo -- aunque los duplicados detectados en el Kardex SÍ tienen
 * mes (metadata.month), a diferencia de los duplicados de Inventario
 * Inicial/Final, que legítimamente no tienen un mes (son una foto única).
 */
import { Rule006Exporter } from "./Rule006Exporter";
import { ReportHeader } from "./base/ReportHeader";

const header: ReportHeader = {
    companyName: "COMERCIAL L&M EIRL",
    ruc: "20451412508",
    year: 2024
};

describe("Rule006Exporter", () => {
    it("muestra el mes real en la columna Periodo para duplicados del KARDEX (antes mostraba siempre '-')", async () => {
        const exporter = new Rule006Exporter();
        const results = [
            {
                product_code: "006146",
                product_name: "PEGAMENTO DE MADERA",
                risk_level: "MEDIO",
                metadata: {
                    source: "KARDEX",
                    month: 1,
                    occurrences: 2,
                    duplicateOccurrences: [
                        { date: "2024-01-01", document: "00 Saldo Inicial", movementCount: 1 },
                        { date: "2024-01-01", document: "00 Saldo Inicial", movementCount: 2 }
                    ]
                }
            }
        ];

        const workbook = await exporter.export(results, header);
        const sheet = workbook.worksheets[0];
        const periodo = sheet.getRow(5).getCell(1).value;

        expect(periodo).toBe("Enero");
    });

    it("sigue mostrando '-' para duplicados de INVENTARIO (no tienen mes, es una foto única, no algo mensual)", async () => {
        const exporter = new Rule006Exporter();
        const results = [
            {
                product_code: "0000000073",
                product_name: "PRODUCTO X",
                risk_level: "MEDIO",
                metadata: {
                    source: "INVENTARIO_INICIAL",
                    occurrences: 2,
                    rows: [10, 25]
                }
            }
        ];

        const workbook = await exporter.export(results, header);
        const sheet = workbook.worksheets[0];
        const periodo = sheet.getRow(5).getCell(1).value;

        expect(periodo).toBe("-");
    });

    it("trazabilidad de KARDEX (Opción A): identifica cada ocurrencia por fecha/documento/movimientos, ya no sale 'Filas: ' vacío", async () => {
        const exporter = new Rule006Exporter();
        const results = [
            {
                product_code: "0000000001",
                product_name: "PRODUCTO A",
                risk_level: "MEDIO",
                metadata: {
                    source: "KARDEX",
                    month: 1,
                    occurrences: 2,
                    duplicateOccurrences: [
                        { date: "2024-01-01", document: "00 Saldo Inicial", movementCount: 1 },
                        { date: "2024-01-01", document: "00 Saldo Inicial", movementCount: 2 }
                    ]
                }
            }
        ];

        const workbook = await exporter.export(results, header);
        const sheet = workbook.worksheets[0];
        const traceability = String(sheet.getRow(5).getCell(10).value);

        expect(traceability).toContain("Ocurrencia 1: 2024-01-01, Doc. 00 Saldo Inicial, 1 movimiento(s)");
        expect(traceability).toContain("Ocurrencia 2: 2024-01-01, Doc. 00 Saldo Inicial, 2 movimiento(s)");
        expect(traceability).not.toContain("Filas:");
    });

    it("trazabilidad de INVENTARIO sigue mostrando 'Filas: <items>' como antes (esa ruta no se tocó)", async () => {
        const exporter = new Rule006Exporter();
        const results = [
            {
                product_code: "0000000073",
                product_name: "PRODUCTO X",
                risk_level: "MEDIO",
                metadata: {
                    source: "INVENTARIO_INICIAL",
                    occurrences: 2,
                    rows: [10, 25]
                }
            }
        ];

        const workbook = await exporter.export(results, header);
        const sheet = workbook.worksheets[0];
        const traceability = String(sheet.getRow(5).getCell(10).value);

        expect(traceability).toContain("Filas: 10, 25");
    });
});
