
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
        differences: ["Costo valorizado fuera del rango permitido"],
        ...overrides
    };
}

describe("Rule014Exporter", () => {
    it("'Valor esperado' muestra el Inventario de Cierre y 'Valor encontrado' muestra la fórmula, segun el diagrama oficial", async () => {
        const exporter = new Rule014Exporter();
        const results = [{ risk_level: "CRITICO", metadata: baseMetadata() }];

        const workbook = await exporter.export(results, header);
        const sheet = workbook.worksheets[0];

        expect(sheet.getRow(5).getCell(5).value).toBe(990339.13); // Valor esperado = Cierre real
        expect(sheet.getRow(5).getCell(6).value).toBe(981696.93); // Valor encontrado = Formula
    });

    it("YA NO existe la fila de Cantidad -- RULE_014 solo valida costo (confirmado contra el diagrama y el Anexo 03)", async () => {
        const exporter = new Rule014Exporter();
        const results = [{ risk_level: "CRITICO", metadata: baseMetadata() }];

        const workbook = await exporter.export(results, header);
        const sheet = workbook.worksheets[0];

        expect(sheet.rowCount).toBe(5); // 4 de header + 1 sola fila (costo)
        expect(sheet.getRow(5).getCell(4).value).not.toBe("Sumatoria Consolidada - Cantidad");
    });

    it("'Tipo de inconsistencia' coincide con 'Campos con diferencia' -- mismo nombre en los dos lados", async () => {
        const exporter = new Rule014Exporter();
        const results = [{ risk_level: "CRITICO", metadata: baseMetadata() }];

        const workbook = await exporter.export(results, header);
        const sheet = workbook.worksheets[0];

        const tipo = sheet.getRow(5).getCell(4).value;
        const trace = String(sheet.getRow(5).getCell(10).value);

        expect(tipo).toBe("Costo valorizado fuera del rango permitido");
        expect(trace).toContain(`Campos con diferencia: ${tipo}`);
    });

    it("'Rango Permitido' ya no muestra el mismo numero repetido como si fuera un rango -- ahora es un texto claro (0% tolerancia = sin tolerancia)", async () => {
        const exporter = new Rule014Exporter();
        const results = [{ risk_level: "CRITICO", metadata: baseMetadata() }];

        const workbook = await exporter.export(results, header);
        const sheet = workbook.worksheets[0];
        const trace = String(sheet.getRow(5).getCell(10).value);

        expect(trace).not.toContain("981696.93 - 981696.93");
        expect(trace).toContain("Sin tolerancia");
    });
});
