
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
            exit: { quantity: 37, totalCost: 564.3, totalCostArchivo: 554.8 }
        },
        expectedFinalBalance: { quantity: 95, unitCost: 5.94, totalCost: 564.3 },
        costTolerance: { percentage: 0, lowerLimit: 554.8, upperLimit: 554.8 },
        actualFinalBalance: { quantity: 95, unitCost: 5.84, totalCost: 554.8 },
        difference: { quantity: 0, unitCost: 0.1, totalCost: 9.5 },
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

        // Solo 2 filas de datos (filas 5 y 6), nada mas -- antes de este
        // fix, ademas de estas 2 habria una tercera fila fantasma de
        // "Sumatoria Mensual - Cantidad" con Diferencia 0.
        expect(sheet.rowCount).toBe(6);
        expect(sheet.getRow(5).getCell(4).value).toBe("Costo Total de Salidas");
        expect(sheet.getRow(6).getCell(4).value).toBe("Costo Valorizado Mensual");

        // Ninguna fila debe decir "Sumatoria Mensual - Cantidad"
        for (let r = 5; r <= sheet.rowCount; r++) {
            expect(sheet.getRow(r).getCell(4).value).not.toBe("Sumatoria Mensual - Cantidad");
        }
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
        expect(sheet.getRow(6).getCell(4).value).toBe("Costo Valorizado Mensual");
        expect(sheet.getRow(7).getCell(4).value).toBe("Costo Unitario de Saldo Final");
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
