import { Rule010Exporter } from "./Rule010Exporter";
import { ReportHeader } from "./base/ReportHeader";

const header: ReportHeader = {
    companyName: "COMERCIAL L&M EIRL",
    ruc: "20451412508",
    year: 2024
};

describe("Rule010Exporter", () => {
    it("muestra la CANTIDAD real del faltante en Valor Encontrado, segun el diagrama oficial (antes mostraba el codigo de operacion '28', y luego -por error- el costo)", async () => {
        const exporter = new Rule010Exporter();
        const results = [
            {
                product_code: "009833",
                product_name: "ZAM-3A CARETA P/SOLDAR ALEMANA NEGRA S/M",
                risk_level: "ALTO",
                metadata: {
                    date: "2023-12-31",
                    month: 12,
                    document: "28 Ajuste Dif Inventario",
                    operation: "28",
                    exitQuantity: 1,
                    exitUnitCost: 2.35,
                    exitTotalCost: 2.35,
                    balanceQuantity: 0,
                    balanceUnitCost: 2.35,
                    balanceTotalCost: 0
                }
            }
        ];

        const workbook = await exporter.export(results, header);
        const sheet = workbook.worksheets[0];
        const row = sheet.getRow(5);

        expect(row.getCell(5).value).toBe(0);   // Valor Esperado, segun el diagrama
        expect(row.getCell(6).value).toBe(1);   // Valor Encontrado = CANTIDAD (1), segun el diagrama -- NO el costo (2.35)
        expect(row.getCell(7).value).toBe(1);   // Diferencia = misma cantidad
    });
});
