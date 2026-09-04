
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

    it("la trazabilidad se lee como una cadena de principio a fin: Saldo Inicial -> Entradas/Salidas -> Encontrado -> Esperado -> Diferencia (ya NO repite 'Campos con diferencia', eso ya lo dice 'Tipo de inconsistencia')", async () => {
        const exporter = new Rule014Exporter();
        const results = [{ risk_level: "CRITICO", metadata: baseMetadata() }];

        const workbook = await exporter.export(results, header);
        const sheet = workbook.worksheets[0];

        const tipo = sheet.getRow(5).getCell(4).value;
        const trace = String(sheet.getRow(5).getCell(10).value);

        expect(tipo).toBe("Costo valorizado fuera del rango permitido");
        expect(trace).not.toContain("Campos con diferencia");

        // El orden importa: es una cadena de principio (Saldo Inicial) a fin (Diferencia).
        const puntoDePartida = trace.indexOf("Punto de partida - Saldo Inicial consolidado del mes: 945647.12");
        const entradas = trace.indexOf("+ Total Entradas (Compras, Op. 02): 310438.92");
        const salidas = trace.indexOf("− Total Salidas (Ventas, Op. 01): 274389.11");
        const encontrado = trace.indexOf("= Resultado de la fórmula (Encontrado): 981696.93");
        const esperado = trace.indexOf("Cierre real que trae el Kardex (Esperado): 990339.13");
        const diferencia = trace.indexOf("Diferencia: 8642.2");

        expect(puntoDePartida).toBeGreaterThanOrEqual(0);
        expect(entradas).toBeGreaterThan(puntoDePartida);
        expect(salidas).toBeGreaterThan(entradas);
        expect(encontrado).toBeGreaterThan(salidas);
        expect(esperado).toBeGreaterThan(encontrado);
        expect(diferencia).toBeGreaterThan(esperado);
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

    it("la trazabilidad incluye el mensaje literal exigido por el Anexo 03: 'ERROR DE CONSOLIDACIÓN DEL KARDEX Y LA DIFERENCIA'", async () => {
        const exporter = new Rule014Exporter();
        const results = [{ risk_level: "CRITICO", metadata: baseMetadata() }];

        const workbook = await exporter.export(results, header);
        const sheet = workbook.worksheets[0];
        const trace = String(sheet.getRow(5).getCell(10).value);

        expect(trace).toContain("ERROR DE CONSOLIDACIÓN DEL KARDEX Y LA DIFERENCIA");
    });
});
