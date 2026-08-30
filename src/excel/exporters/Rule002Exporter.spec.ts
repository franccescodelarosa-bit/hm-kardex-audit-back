
import { Rule002Exporter } from "./Rule002Exporter";
import { ReportHeader } from "./base/ReportHeader";

const header: ReportHeader = {
    companyName: "COMERCIAL L&M EIRL",
    ruc: "20451412508",
    year: 2024
};

describe("Rule002Exporter", () => {
    it("la trazabilidad muestra las cantidades reales y la diferencia, no solo la etiqueta 'Cantidad'", async () => {
        const exporter = new Rule002Exporter();
        const results = [
            {
                product_code: "006749",
                product_name: "PEGAMENTO E/LATA DE 1/32 T/MED.M/AFRICANO",
                risk_level: "ALTO",
                metadata: {
                    fromMonth: 1,
                    toMonth: 2,
                    finalQuantity: 260,
                    initialQuantity: 247,
                    difference: 13
                }
            }
        ];

        const workbook = await exporter.export(results, header);
        const sheet = workbook.worksheets[0];
        const traceability = String(sheet.getRow(5).getCell(10).value);

        // Ya no debe decir solo "Campos con diferencia: Cantidad" sin numeros
        expect(traceability).not.toBe("Mes Cierre: Enero\nMes Inicial: Febrero\nCampos con diferencia: Cantidad");

        // Debe mostrar las cantidades reales, mapeadas con su mes
        expect(traceability).toContain("Cantidad Saldo Final (Enero): 260");
        expect(traceability).toContain("Cantidad Saldo Inicial (Febrero): 247");
        expect(traceability).toContain("Diferencia");
        expect(traceability).toContain("13");
    });

    it("si el producto no existe en el mes siguiente (initialQuantity null), no inventa una diferencia falsa", async () => {
        const exporter = new Rule002Exporter();
        const results = [
            {
                product_code: "000144",
                product_name: "ALCOHOL YODADO D/30ML M/D LEOS",
                risk_level: "ALTO",
                metadata: {
                    fromMonth: 3,
                    toMonth: 4,
                    finalQuantity: 50,
                    initialQuantity: null,
                    difference: null
                }
            }
        ];

        const workbook = await exporter.export(results, header);
        const sheet = workbook.worksheets[0];
        const traceability = String(sheet.getRow(5).getCell(10).value);
        const foundValue = sheet.getRow(5).getCell(6).value;

        expect(traceability).not.toContain("null");
        expect(traceability).toContain("no tiene Kardex registrado");
        expect(foundValue).toBe(0);
    });
});
