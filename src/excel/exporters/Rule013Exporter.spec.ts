
import { Rule013Exporter } from "./Rule013Exporter";
import { ReportHeader } from "./base/ReportHeader";

const header: ReportHeader = {
    companyName: "COMERCIAL L&M EIRL",
    ruc: "20451412508",
    year: 2024
};

/**
 * Metadata basada en el caso real 000129 (Enero 2024), con el motor
 * re-anclado POR ENTRADA (confirmado con el cliente vía capturas del
 * Excel de referencia, fórmula por fórmula): las DOS entradas del mes
 * (18/01 y 19/01) tienen su propio Costo Total desalineado de la fórmula
 * re-anclada -- eso es "Costo Unitario de Saldo Final", UNA fila por
 * cada una. Salidas coincide (222.42), no genera fila. Y el cierre real
 * del mes (95 unidades, después de la salida) tampoco coincide -- eso es
 * "Costo Total de Saldo Final", UNA sola fila.
 */
function baseMetadata(overrides: Partial<any> = {}) {
    return {
        month: 1,
        normalizedCode: "129",
        totals: {
            exit: { quantity: 37, totalCost: 222.42, totalCostArchivo: 222.42 }
        },
        unitCostMismatches: [
            { date: null, document: "F001-00004970", quantity: 60, expectedTotalCost: 358.20, foundTotalCost: 346.80 },
            { date: null, document: "F001-00004994", quantity: 120, expectedTotalCost: 712.80, foundTotalCost: 712.20 }
        ],
        // Ni el Costo Unitario ni el Costo Total calculados vienen
        // redondeados a centavos -- así se ve la diferencia real
        // (563.825 vs 564.30), no "564.30 vs 564.30".
        expectedFinalBalance: { quantity: 95, unitCost: 5.935, totalCost: 563.825 },
        actualFinalBalance: { quantity: 95, unitCost: 5.94, totalCost: 564.3 },
        movementCount: 22,
        differences: ["Costo Unitario de Saldo Final", "Costo Total de Saldo Final"],
        ...overrides
    };
}

describe("Rule013Exporter", () => {
    it("reproduce el caso real (producto 000129): Salidas coincide, las 2 entradas del mes tienen su propio Costo Unitario mal, y el cierre del mes (Total de Saldo Final) también falla -> 3 filas", async () => {
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

        expect(sheet.rowCount).toBe(7); // 4 de header + 2 (Unitario, una por entrada) + 1 (Total)

        // Orden confirmado con el cliente: Salidas -> Costo Unitario de
        // Saldo Final (una fila por entrada) -> Costo Total de Saldo Final.
        expect(sheet.getRow(5).getCell(4).value).toBe("Costo Unitario de Saldo Final");
        expect(sheet.getRow(5).getCell(5).value).toBe(358.2); // esperado = archivo, entrada 18/01
        expect(sheet.getRow(5).getCell(6).value).toBe(346.8); // encontrado = fórmula re-anclada

        expect(sheet.getRow(6).getCell(4).value).toBe("Costo Unitario de Saldo Final");
        expect(sheet.getRow(6).getCell(5).value).toBe(712.8); // esperado = archivo, entrada 19/01
        expect(sheet.getRow(6).getCell(6).value).toBe(712.2); // encontrado = fórmula re-anclada

        expect(sheet.getRow(7).getCell(4).value).toBe("Costo Total de Saldo Final");
        expect(sheet.getRow(7).getCell(5).value).toBe(564.3);
        expect(sheet.getRow(7).getCell(6).value).toBe(563.825);
    });

    it("'Costo Total de Salidas': esperado = lo que dice el archivo, encontrado = lo que la regla calculó con el CPP", async () => {
        const exporter = new Rule013Exporter();
        const results = [
            {
                product_code: "000144",
                product_name: "ALCOHOL YODADO D/30ML M/D LEOS",
                risk_level: "CRITICO",
                metadata: baseMetadata({
                    totals: { exit: { quantity: 8, totalCost: 46.57, totalCostArchivo: 46.80 } },
                    unitCostMismatches: [],
                    differences: ["Costo Total de Salidas"]
                })
            }
        ];

        const workbook = await exporter.export(results, header);
        const sheet = workbook.worksheets[0];

        expect(sheet.getRow(5).getCell(4).value).toBe("Costo Total de Salidas");
        expect(sheet.getRow(5).getCell(5).value).toBe(46.80); // esperado = archivo
        expect(sheet.getRow(5).getCell(6).value).toBe(46.57); // encontrado = CPP calculado
    });

    it("'Costo Unitario de Saldo Final': una fila por entrada, esperado = archivo (esa fila), encontrado = fórmula re-anclada (esa misma fila)", async () => {
        const exporter = new Rule013Exporter();
        const results = [
            {
                product_code: "000129",
                product_name: "AGUJA PLATEADA E/DISCO 24-1 ROSADA M/NEEDLES",
                risk_level: "CRITICO",
                metadata: baseMetadata({
                    unitCostMismatches: [
                        { date: null, document: "F001-00004970", quantity: 60, expectedTotalCost: 358.20, foundTotalCost: 346.80 }
                    ],
                    differences: ["Costo Unitario de Saldo Final"]
                })
            }
        ];

        const workbook = await exporter.export(results, header);
        const sheet = workbook.worksheets[0];

        expect(sheet.rowCount).toBe(5); // 4 de header + 1 sola entrada con error
        expect(sheet.getRow(5).getCell(4).value).toBe("Costo Unitario de Saldo Final");
        expect(sheet.getRow(5).getCell(5).value).toBe(358.2); // esperado = archivo
        expect(sheet.getRow(5).getCell(6).value).toBe(346.8); // encontrado = fórmula re-anclada

        const trace = String(sheet.getRow(5).getCell(10).value);
        expect(trace).toContain("Documento: F001-00004970");
        expect(trace).toContain("Cantidad de esa Entrada: 60");
    });

    it("'Costo Total de Saldo Final': esperado = archivo (última fila del mes), encontrado = CPP de la última entrada x cantidad final real", async () => {
        const exporter = new Rule013Exporter();
        const results = [
            {
                product_code: "000129",
                product_name: "AGUJA PLATEADA E/DISCO 24-1 ROSADA M/NEEDLES",
                risk_level: "CRITICO",
                metadata: baseMetadata({
                    unitCostMismatches: [],
                    differences: ["Costo Total de Saldo Final"]
                })
            }
        ];

        const workbook = await exporter.export(results, header);
        const sheet = workbook.worksheets[0];

        expect(sheet.getRow(5).getCell(4).value).toBe("Costo Total de Saldo Final");
        expect(sheet.getRow(5).getCell(5).value).toBe(564.3); // esperado = archivo
        expect(sheet.getRow(5).getCell(6).value).toBe(563.825); // encontrado = CPP x cantidad final, sin redondeo visual

        const trace = String(sheet.getRow(5).getCell(10).value);
        expect(trace).toContain("Cantidad Final del Mes: 95");
    });

    it("si las 3 validaciones fallan, muestra 3 filas en el orden confirmado (Salidas -> Unitario -> Total)", async () => {
        const exporter = new Rule013Exporter();
        const results = [
            {
                product_code: "000144",
                product_name: "ALCOHOL YODADO D/30ML M/D LEOS",
                risk_level: "CRITICO",
                metadata: baseMetadata({
                    unitCostMismatches: [
                        { date: null, document: "F001-00004970", quantity: 60, expectedTotalCost: 358.20, foundTotalCost: 346.80 }
                    ],
                    differences: ["Costo Total de Salidas", "Costo Unitario de Saldo Final", "Costo Total de Saldo Final"]
                })
            }
        ];

        const workbook = await exporter.export(results, header);
        const sheet = workbook.worksheets[0];

        expect(sheet.rowCount).toBe(7); // 4 de header + 3 de datos
        expect(sheet.getRow(5).getCell(4).value).toBe("Costo Total de Salidas");
        expect(sheet.getRow(6).getCell(4).value).toBe("Costo Unitario de Saldo Final");
        expect(sheet.getRow(7).getCell(4).value).toBe("Costo Total de Saldo Final");
    });

    it("trazabilidad autocontenida por fila: 'Campos con diferencia' SOLO con la diferencia de esa fila puntual, no todas juntas", async () => {
        const exporter = new Rule013Exporter();
        const results = [
            {
                product_code: "000144",
                product_name: "ALCOHOL YODADO D/30ML M/D LEOS",
                risk_level: "CRITICO",
                metadata: baseMetadata({
                    unitCostMismatches: [
                        { date: null, document: "F001-00004970", quantity: 60, expectedTotalCost: 358.20, foundTotalCost: 346.80 }
                    ],
                    differences: ["Costo Total de Salidas", "Costo Unitario de Saldo Final", "Costo Total de Saldo Final"]
                })
            }
        ];

        const workbook = await exporter.export(results, header);
        const sheet = workbook.worksheets[0];

        const trace1 = String(sheet.getRow(5).getCell(10).value);
        expect(trace1).toContain("Campos con diferencia: Costo Total de Salidas");
        expect(trace1).not.toContain("Costo Unitario de Saldo Final");
        expect(trace1).not.toContain("Costo Total de Saldo Final");

        const trace2 = String(sheet.getRow(6).getCell(10).value);
        expect(trace2).toContain("Campos con diferencia: Costo Unitario de Saldo Final");
        expect(trace2).not.toContain("Costo Total de Salidas");
        expect(trace2).not.toContain("Costo Total de Saldo Final");

        const trace3 = String(sheet.getRow(7).getCell(10).value);
        expect(trace3).toContain("Campos con diferencia: Costo Total de Saldo Final");
        expect(trace3).not.toContain("Costo Total de Salidas");
        expect(trace3).not.toContain("Costo Unitario de Saldo Final");
    });

    it("si solo 1 validacion falla (ej. solo 1 entrada con Costo Unitario mal), muestra 1 sola fila", async () => {
        const exporter = new Rule013Exporter();
        const results = [
            {
                product_code: "000200",
                product_name: "PRODUCTO EJEMPLO",
                risk_level: "CRITICO",
                metadata: baseMetadata({
                    unitCostMismatches: [
                        { date: null, document: "F001-00000001", quantity: 10, expectedTotalCost: 100, foundTotalCost: 95 }
                    ],
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
