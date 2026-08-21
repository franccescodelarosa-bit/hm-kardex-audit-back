import { Rule009Exporter } from "./Rule009Exporter";
import { ReportHeader } from "./base/ReportHeader";

const header: ReportHeader = {
    companyName: "COMERCIAL L&M EIRL",
    ruc: "20451412508",
    year: 2024
};

describe("Rule009Exporter", () => {
    it("muestra la CANTIDAD real del sobrante en Valor Encontrado, segun el mismo criterio del diagrama oficial de Regla 10 (antes mostraba el codigo de operacion '28')", async () => {
        const exporter = new Rule009Exporter();
        const results = [
            {
                product_code: "000087",
                product_name: "ACEITE P/BEBE OIL W/VIT 12ONZ M/PERFECT PURITY BABY",
                risk_level: "ALTO",
                metadata: {
                    date: "2024-12-31",
                    month: 12,
                    document: "28 Ajuste Dif Inventario",
                    operation: "28",
                    entryQuantity: 3,
                    entryUnitCost: 10.75,
                    entryTotalCost: 32.25,
                    balanceQuantity: 87,
                    balanceUnitCost: 10.75,
                    balanceTotalCost: 935.25
                }
            }
        ];

        const workbook = await exporter.export(results, header);
        const sheet = workbook.worksheets[0];
        const row = sheet.getRow(5);

        expect(row.getCell(5).value).toBe(0);   // Valor Esperado
        expect(row.getCell(6).value).toBe(3);   // Valor Encontrado = CANTIDAD (3), NO el costo (32.25)
        expect(row.getCell(7).value).toBe(3);   // Diferencia = misma cantidad
    });
});
