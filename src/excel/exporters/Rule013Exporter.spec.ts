
import { Rule013Exporter } from "./Rule013Exporter";
import { ReportHeader } from "./base/ReportHeader";

const header: ReportHeader = {
    companyName: "COMERCIAL L&M EIRL",
    ruc: "20451412508",
    year: 2024
};

function baseMetadata(overrides: Partial<any> = {}) {
    return {
        month: 1,
        normalizedCode: "129",
        initialBalance: { quantity: 12, totalCost: 73.92 },
        totals: {
            entry: { quantity: 120, totalCost: 700.8 },
            exit: { quantity: 37, totalCost: 219.92, totalCostArchivo: 222.42 }
        },
        expectedFinalBalance: { quantity: 95, unitCost: 5.84, totalCost: 554.8 },
        costTolerance: { percentage: 0, lowerLimit: 554.8, upperLimit: 554.8 },
        actualFinalBalance: { quantity: 95, unitCost: 5.94, totalCost: 564.3 },
        difference: { quantity: 0, unitCost: -0.1, totalCost: -9.5 },
        movementCount: 22,
        differences: ["Costo Total de Salidas", "Costo Total de Saldo Final"],
        ...overrides
    };
}

describe("Rule013Exporter", () => {
    it("reproduce el ejemplo real (producto 000129): 2 diferencias reales -> exactamente 2 filas, SIN la fila fantasma de Cantidad", async () => {
        const exporter = new Rule013Exporter();
        const results = [
            {
                product_code: "000129",
                product_name: "AGUJA PLATEADA E/DISCO 24-1 ROSADA M/NEEDLES",
                risk_level: "CRITICO",
                metadata: baseMetadata()
            }
        ];

        const workbook = await exporter.export(results, header);
        const sheet = workbook.worksheets[0];

        expect(sheet.rowCount).toBe(6);
        expect(sheet.getRow(5).getCell(4).value).toBe("Costo Total de Salidas");
        expect(sheet.getRow(6).getCell(4).value).toBe("Costo Total de Saldo Final");

        for (let r = 5; r <= sheet.rowCount; r++) {
            expect(sheet.getRow(r).getCell(4).value).not.toBe("Sumatoria Mensual - Cantidad");
        }
    });

    it("'Valor esperado' = lo que dice el ARCHIVO, 'Valor encontrado' = lo que la regla calculó con el CPP -- según el diagrama oficial y el Anexo 02", async () => {
        const exporter = new Rule013Exporter();
        const results = [
            {
                product_code: "000129",
                product_name: "AGUJA PLATEADA E/DISCO 24-1 ROSADA M/NEEDLES",
                risk_level: "CRITICO",
                metadata: baseMetadata()
            }
        ];

        const workbook = await exporter.export(results, header);
        const sheet = workbook.worksheets[0];

        // Fila "Costo Total de Salidas" (5)
        expect(sheet.getRow(5).getCell(5).value).toBe(222.42); // esperado = archivo
        expect(sheet.getRow(5).getCell(6).value).toBe(219.92); // encontrado = mi cálculo (CPP)

        // Fila "Costo Valorizado Mensual" (6)
        expect(sheet.getRow(6).getCell(5).value).toBe(564.3); // esperado = archivo
        expect(sheet.getRow(6).getCell(6).value).toBe(554.8); // encontrado = mi cálculo (CPP)
    });

    it("fila 'Costo Unitario de Saldo Final': tambien esperado=archivo, encontrado=CPP", async () => {
        const exporter = new Rule013Exporter();
        const results = [
            {
                product_code: "000129",
                product_name: "AGUJA PLATEADA E/DISCO 24-1 ROSADA M/NEEDLES",
                risk_level: "CRITICO",
                metadata: baseMetadata({ differences: ["Costo Unitario de Saldo Final"] })
            }
        ];

        const workbook = await exporter.export(results, header);
        const sheet = workbook.worksheets[0];

        expect(sheet.getRow(5).getCell(4).value).toBe("Costo Unitario de Saldo Final");
        expect(sheet.getRow(5).getCell(5).value).toBe(5.94); // esperado = archivo
        expect(sheet.getRow(5).getCell(6).value).toBe(5.84); // encontrado = mi cálculo (CPP)
    });

    it("si las 3 validaciones fallan, muestra 3 filas", async () => {
        const exporter = new Rule013Exporter();
        const results = [
            {
                product_code: "000144",
                product_name: "ALCOHOL YODADO D/30ML M/D LEOS",
                risk_level: "CRITICO",
                metadata: baseMetadata({
                    differences: ["Costo Total de Salidas", "Costo Total de Saldo Final", "Costo Unitario de Saldo Final"]
                })
            }
        ];

        const workbook = await exporter.export(results, header);
        const sheet = workbook.worksheets[0];

        expect(sheet.rowCount).toBe(7); // 4 filas de header + 3 de datos
        expect(sheet.getRow(5).getCell(4).value).toBe("Costo Total de Salidas");
        expect(sheet.getRow(6).getCell(4).value).toBe("Costo Total de Saldo Final");
        expect(sheet.getRow(7).getCell(4).value).toBe("Costo Unitario de Saldo Final");
    });

    it("trazabilidad autocontenida por fila: nombres correctos (Esperado/Encontrado) y 'Campos con diferencia' SOLO con la diferencia de esa fila puntual, no todas juntas", async () => {
        const exporter = new Rule013Exporter();
        const results = [
            {
                product_code: "000144",
                product_name: "ALCOHOL YODADO D/30ML M/D LEOS",
                risk_level: "CRITICO",
                metadata: baseMetadata({
                    differences: ["Costo Total de Salidas", "Costo Total de Saldo Final", "Costo Unitario de Saldo Final"]
                })
            }
        ];

        const workbook = await exporter.export(results, header);
        const sheet = workbook.worksheets[0];

        // Fila 1: Costo Total de Salidas
        const trace1 = String(sheet.getRow(5).getCell(10).value);
        expect(trace1).toContain("Costo de Salidas Esperado (Archivo): 222.42");
        expect(trace1).toContain("Costo de Salidas Encontrado (CPP): 219.92");
        expect(trace1).toContain("Campos con diferencia: Costo Total de Salidas");
        expect(trace1).not.toContain("Costo Total de Saldo Final");
        expect(trace1).not.toContain("Costo Unitario de Saldo Final");

        // Fila 2: Costo Valorizado Mensual (Costo Total de Saldo Final)
        const trace2 = String(sheet.getRow(6).getCell(10).value);
        expect(trace2).toContain("Campos con diferencia: Costo Total de Saldo Final");
        expect(trace2).not.toContain("Costo Total de Salidas");
        expect(trace2).not.toContain("Costo Unitario de Saldo Final");

        // Fila 3: Costo Unitario de Saldo Final
        const trace3 = String(sheet.getRow(7).getCell(10).value);
        expect(trace3).toContain("Campos con diferencia: Costo Unitario de Saldo Final");
        expect(trace3).not.toContain("Costo Total de Salidas");
        expect(trace3).not.toContain("Costo Total de Saldo Final");
    });

    it("si solo 1 validacion falla (ej. solo Costo Unitario), muestra 1 sola fila", async () => {
        const exporter = new Rule013Exporter();
        const results = [
            {
                product_code: "000200",
                product_name: "PRODUCTO EJEMPLO",
                risk_level: "CRITICO",
                metadata: baseMetadata({
                    differences: ["Costo Unitario de Saldo Final"]
                })
            }
        ];

        const workbook = await exporter.export(results, header);
        const sheet = workbook.worksheets[0];

        expect(sheet.rowCount).toBe(5); // 4 filas de header + 1 sola fila de datos
        expect(sheet.getRow(5).getCell(4).value).toBe("Costo Unitario de Saldo Final");
    });
});
