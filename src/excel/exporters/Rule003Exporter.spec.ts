
import { Rule003Exporter } from "./Rule003Exporter";
import { ReportHeader } from "./base/ReportHeader";

const header: ReportHeader = {
    companyName: "COMERCIAL L&M EIRL",
    ruc: "20451412508",
    year: 2024
};

function baseMetadata(overrides: Partial<any> = {}) {
    return {
        fromIndex: 11,
        toIndex: 12,
        finalBalance: { quantity: 100, unitCost: 8, totalCost: 904 },
        initialBalance: { quantity: 100, unitCost: 7, totalCost: 791 },
        differences: ["Costo Unitario", "Costo Total"],
        ...overrides
    };
}

describe("Rule003Exporter", () => {
    it("siempre muestra las 2 filas (Costo Unitario y Costo Total), aunque UNA de ellas tenga 0% de diferencia -- asi lo pidio el equipo", async () => {
        const exporter = new Rule003Exporter();
        const results = [
            {
                product_code: "006749",
                product_name: "PEGAMENTO E/LATA DE 1/32 T/MED.M/AFRICANO",
                risk_level: "ALTO",
                metadata: baseMetadata({
                    // Caso real: Costo Unitario identico (4.04 = 4.04), solo Costo Total difirio
                    finalBalance: { quantity: 247, unitCost: 4.04, totalCost: 997.88 },
                    initialBalance: { quantity: 251, unitCost: 4.04, totalCost: 1014.04 },
                    differences: ["Costo Total"]
                })
            }
        ];

        const workbook = await exporter.export(results, header);
        const sheet = workbook.worksheets[0];

        expect(sheet.rowCount).toBe(6); // 4 de header + 2 de datos, siempre

        // Fila 1: Costo Unitario, con 0 de diferencia real
        expect(sheet.getRow(5).getCell(4).value).toBe("Continuidad de Costo Unitario");
        expect(sheet.getRow(5).getCell(7).value).toBe(0); // diferencia
        const unitRowTrace = String(sheet.getRow(5).getCell(10).value);
        expect(unitRowTrace).toContain("Campos con diferencia: Costo Unitario");
        expect(unitRowTrace).not.toContain("Costo Total");

        // Fila 2: Costo Total, con la diferencia real
        expect(sheet.getRow(6).getCell(4).value).toBe("Continuidad de Costo Total");
        const totalRowTrace = String(sheet.getRow(6).getCell(10).value);
        expect(totalRowTrace).toContain("Campos con diferencia: Costo Total");
        expect(totalRowTrace).not.toContain("Costo Unitario");
    });

    it("si las dos difieren, muestra las 2 filas, cada trazabilidad habla SOLO de su propio campo (sin mezclar)", async () => {
        const exporter = new Rule003Exporter();
        const results = [
            {
                product_code: "019706",
                product_name: "SHORT B/VARON TELA WING JASPEADO C/BOLS",
                risk_level: "ALTO",
                metadata: baseMetadata()
            }
        ];

        const workbook = await exporter.export(results, header);
        const sheet = workbook.worksheets[0];

        expect(sheet.rowCount).toBe(6);

        const unitRowTrace = String(sheet.getRow(5).getCell(10).value);
        expect(unitRowTrace).toContain("Campos con diferencia: Costo Unitario");
        expect(unitRowTrace).not.toContain("Costo Total");

        const totalRowTrace = String(sheet.getRow(6).getCell(10).value);
        expect(totalRowTrace).toContain("Campos con diferencia: Costo Total");
        expect(totalRowTrace).not.toContain("Costo Unitario");
    });

    it("si el producto no existe en el mes siguiente (initialBalance null), no revienta y muestra una fila clara", async () => {
        const exporter = new Rule003Exporter();
        const results = [
            {
                product_code: "000144",
                product_name: "ALCOHOL YODADO D/30ML M/D LEOS",
                risk_level: "ALTO",
                metadata: {
                    fromIndex: 3,
                    toIndex: 4,
                    finalBalance: { quantity: 50, unitCost: 5, totalCost: 250 },
                    initialBalance: null,
                    differences: undefined
                }
            }
        ];

        const workbook = await exporter.export(results, header);
        const sheet = workbook.worksheets[0];

        expect(sheet.rowCount).toBe(5); // 4 de header + 1 sola fila
        const trace = String(sheet.getRow(5).getCell(10).value);
        expect(trace).not.toContain("null");
        expect(trace).toContain("no tiene Kardex registrado");
    });
});
