
import { Rule014Exporter } from "./Rule014Exporter";
import { ReportHeader } from "./base/ReportHeader";

const header: ReportHeader = {
    companyName: "COMERCIAL L&M EIRL",
    ruc: "20451412508",
    year: 2024
};

function baseMetadata(overrides: Partial<any> = {}) {
    return {
        month: 1,
        initialBalance: { quantity: 217963, totalCost: 945647.12 },
        totals: {
            entry: { quantity: 88017, totalCost: 310438.92 },
            exit: { quantity: 74947, totalCost: 274389.11 }
        },
        // Resultado de la FORMULA (Inicio + Entrada - Salida) -- el diagrama lo llama "valor encontrado"
        expectedFinalBalance: { quantity: 231033, totalCost: 981696.93 },
        costTolerance: { percentage: 0, lowerLimit: 981696.93, upperLimit: 981696.93 },
        // Inventario Valorizado de CIERRE real -- el diagrama lo llama "valor esperado"
        actualFinalBalance: { quantity: 231100, totalCost: 990339.13 },
        difference: { quantity: -67, totalCost: -8642.2 },
        productCount: 40985,
        movementCount: 52422,
        differences: ["Cantidad", "Costo valorizado fuera del rango permitido"],
        ...overrides
    };
}

describe("Rule014Exporter", () => {
    it("'Valor esperado' muestra el Inventario de Cierre y 'Valor encontrado' muestra la fórmula, segun el diagrama oficial", async () => {
        const exporter = new Rule014Exporter();
        const results = [{ risk_level: "CRITICO", metadata: baseMetadata() }];

        const workbook = await exporter.export(results, header);
        const sheet = workbook.worksheets[0];

        // Fila de Cantidad (5)
        expect(sheet.getRow(5).getCell(5).value).toBe(231100); // Valor esperado = Cierre real
        expect(sheet.getRow(5).getCell(6).value).toBe(231033); // Valor encontrado = Formula

        // Fila de Costo (6)
        expect(sheet.getRow(6).getCell(5).value).toBe(990339.13); // Valor esperado = Cierre real
        expect(sheet.getRow(6).getCell(6).value).toBe(981696.93); // Valor encontrado = Formula
    });

    it("si SOLO el Costo esta mal (Cantidad cierra bien), muestra 1 sola fila -- no la fantasma de Cantidad", async () => {
        const exporter = new Rule014Exporter();
        const results = [{
            risk_level: "CRITICO",
            metadata: baseMetadata({ differences: ["Costo valorizado fuera del rango permitido"] })
        }];

        const workbook = await exporter.export(results, header);
        const sheet = workbook.worksheets[0];

        expect(sheet.rowCount).toBe(5); // 4 de header + 1 de datos
        expect(sheet.getRow(5).getCell(4).value).toBe("Costo Valorizado Consolidado");
    });

    it("si SOLO la Cantidad esta mal (Costo cierra bien), muestra 1 sola fila -- no la fantasma de Costo", async () => {
        const exporter = new Rule014Exporter();
        const results = [{
            risk_level: "CRITICO",
            metadata: baseMetadata({ differences: ["Cantidad"] })
        }];

        const workbook = await exporter.export(results, header);
        const sheet = workbook.worksheets[0];

        expect(sheet.rowCount).toBe(5);
        expect(sheet.getRow(5).getCell(4).value).toBe("Sumatoria Consolidada - Cantidad");
    });

    it("si las dos fallan, muestra las 2 filas (comportamiento real de todos los hallazgos actuales)", async () => {
        const exporter = new Rule014Exporter();
        const results = [{ risk_level: "CRITICO", metadata: baseMetadata() }];

        const workbook = await exporter.export(results, header);
        const sheet = workbook.worksheets[0];

        expect(sheet.rowCount).toBe(6);
    });
});
