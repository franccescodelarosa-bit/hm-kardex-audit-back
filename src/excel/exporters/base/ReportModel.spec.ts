
import ExcelJS from "exceljs";

import { ReportHeader } from "./ReportHeader";
import { Rule001Exporter } from "../Rule001Exporter";
import { Rule002Exporter } from "../Rule002Exporter";
import { Rule003Exporter } from "../Rule003Exporter";
import { Rule004Exporter } from "../Rule004Exporter";
import { Rule005Exporter } from "../Rule005Exporter";
import { Rule006Exporter } from "../Rule006Exporter";
import { Rule007Exporter } from "../Rule007Exporter";
import { Rule008Exporter } from "../Rule008Exporter";
import { Rule009Exporter } from "../Rule009Exporter";
import { Rule010Exporter } from "../Rule010Exporter";
import { Rule011Exporter } from "../Rule011Exporter";
import { Rule012Exporter } from "../Rule012Exporter";
import { Rule013Exporter } from "../Rule013Exporter";
import { Rule014Exporter } from "../Rule014Exporter";

/**
 * MODELO DE REPORTE aprobado por el cliente: los reportes RULE_001..014
 * que el equipo retocó a mano en Excel (auditoría del 29.09.2026) y que
 * ahora tienen que salir ASÍ directamente del sistema:
 *
 *  - Fila 1: título combinado, en negrita.
 *  - Fila 2: Empresa / RUC / Periodo en negrita + cartel amarillo con
 *    letra roja "AUDITORIA dd.mm.aaaa" en la ÚLTIMA columna.
 *  - Fila 3: encabezado (negrita, centrado, con bordes).
 *  - Datos desde la fila 4 -- SIN fila vacía en el medio.
 *  - Panel congelado en la fila 3 y sin autofiltro.
 *  - Todas las celdas en negrita, con bordes finos y texto ajustado.
 *  - Montos con formato #,##0.00.
 *  - Anchos de columna y tamaños de letra propios de cada regla.
 *
 * Excepción: RULE_001 tiene un SEGUNDO renglón de encabezado (fila 4) con
 * "Valor esperado del inventario final" / "Valor encontrado del Kardex",
 * datos desde la fila 5, congelado en la 4 y autofiltro A4:J4.
 */

// 29.09.2026 22:00 en Lima = 30.09.2026 03:00 UTC: el cartel tiene que
// mostrar el día de LIMA, no el de UTC.
const AUDIT_DATE = new Date("2026-09-30T03:00:00.000Z");

const header: ReportHeader = {
    companyName: "EXCLUSIVIDADES KIMBERLY E.I.R.L",
    ruc: "20493342062",
    year: 2025,
    auditDate: AUDIT_DATE
};

const RED = "FFFF0000";
const YELLOW = "FFFFFF00";

async function sheetOf(exporter: { export(r: any[], h: ReportHeader): Promise<ExcelJS.Workbook> }, results: any[]) {
    const workbook = await exporter.export(results, header);
    return workbook.worksheets[0];
}

function expectThinBorders(cell: ExcelJS.Cell) {
    for (const side of ["top", "left", "bottom", "right"] as const) {
        expect(cell.border?.[side]?.style).toBe("thin");
    }
}

/** Resultado mínimo y válido de cada regla (una sola fila de datos). */
const SAMPLES: Record<string, { exporter: any; results: any[]; lastColumn: string; columns: number }> = {
    RULE_002: {
        exporter: new Rule002Exporter(), lastColumn: "J", columns: 10,
        results: [{ product_code: "46045", product_name: "PANTALON", risk_level: "ALTO", metadata: { fromMonth: 1, toMonth: 2, finalQuantity: 3, initialQuantity: 2 } }]
    },
    RULE_003: {
        exporter: new Rule003Exporter(), lastColumn: "J", columns: 10,
        results: [{
            product_code: "83068", product_name: "BERMUDA", risk_level: "ALTO", error_type: "MONTHLY_COST_CONTINUITY_ERROR",
            metadata: {
                fromIndex: 1, toIndex: 2,
                finalBalance: { quantity: 1, unitCost: 229.73, totalCost: 229.73 },
                initialBalance: { quantity: 1, unitCost: 229.73, totalCost: 229.74 },
                differences: ["Costo Total"]
            }
        }]
    },
    RULE_005: {
        exporter: new Rule005Exporter(), lastColumn: "J", columns: 10,
        results: [{
            product_code: "49466", product_name: "MORRAL", risk_level: "CRITICO",
            metadata: { source: "KARDEX", month: 2, document: "D", operation: "11", date: "2025-02-24", balanceQuantity: -2, negatives: ["Cantidad de Saldo"] }
        }]
    },
    RULE_006: {
        exporter: new Rule006Exporter(), lastColumn: "J", columns: 10,
        results: [{ product_code: "11538", product_name: "PANTALON", risk_level: "MEDIO", metadata: { source: "KARDEX", month: 10, occurrences: 3, rows: [1, 2, 3] } }]
    },
    RULE_007: {
        exporter: new Rule007Exporter(), lastColumn: "J", columns: 10,
        results: [{
            product_code: "1", product_name: "X", risk_level: "ALTO",
            metadata: {
                month: 3, movement: 4, operation: "01", document: "D",
                previousBalance: { quantity: 10, totalCost: 100 },
                movimientos: { entryQuantity: 0, exitQuantity: 1, entryTotalCost: 0, exitTotalCost: 10 },
                expectedBalance: { quantity: 9, totalCost: 90 },
                actualBalance: { quantity: 8, totalCost: 80 },
                differences: ["Cantidad"]
            }
        }]
    },
    RULE_008: {
        exporter: new Rule008Exporter(), lastColumn: "J", columns: 10,
        results: [{ product_code: "29461", product_name: "PANTALON NORTON", risk_level: "ALTO", metadata: { source: "INVENTARIO_INICIAL", month: 1, stock: 0 } }]
    },
    RULE_009: {
        exporter: new Rule009Exporter(), lastColumn: "J", columns: 10,
        results: [{
            product_code: "1", product_name: "X", risk_level: "ALTO",
            metadata: { date: "2025-03-01", month: 3, document: "D", operation: "28", entryQuantity: 5, entryUnitCost: 2, entryTotalCost: 10, balanceQuantity: 5, balanceUnitCost: 2, balanceTotalCost: 10 }
        }]
    },
    RULE_010: {
        exporter: new Rule010Exporter(), lastColumn: "J", columns: 10,
        results: [{
            product_code: "1", product_name: "X", risk_level: "ALTO",
            metadata: { date: "2025-03-01", month: 3, document: "D", operation: "28", exitQuantity: 5, exitUnitCost: 2, exitTotalCost: 10, balanceQuantity: 0, balanceUnitCost: 2, balanceTotalCost: 0 }
        }]
    },
    RULE_011: {
        exporter: new Rule011Exporter(), lastColumn: "J", columns: 10,
        results: [{ product_code: "84604", product_name: "MOCHILA", risk_level: "MEDIO", metadata: { date: "2025-02-01", month: 2, document: "00 SALDO INICIAL", operation: "16", previousCost: 112.58, currentCost: 0, variationPercent: 100 } }]
    },
    RULE_012: {
        exporter: new Rule012Exporter(), lastColumn: "J", columns: 10,
        results: [{
            product_code: "FAC-F100-19436", product_name: "", risk_level: "MEDIO",
            metadata: {
                issueDate: "2024-12-26", warehouseDate: "2025-01-07", supplierRuc: "20600892470", supplier: "S",
                document: "FAC-F100-19436", normalizedDocument: "F10000019436",
                expectedCost: 7439.29, kardexCost: 0, difference: 7439.29, differencePercent: 100,
                thresholdPercent: 5, isIncident: true, movements: 0, month: 1, transitItem: "2024-12-26"
            }
        }]
    },
    RULE_013: {
        exporter: new Rule013Exporter(), lastColumn: "J", columns: 10,
        results: [{
            product_code: "54123", product_name: "ZAPATILLA", risk_level: "CRITICO",
            metadata: {
                month: 1, normalizedCode: "54123",
                initialBalance: { quantity: 0, totalCost: 0 },
                totals: { entry: { quantity: 1, totalCost: 257 }, exit: { quantity: 0, totalCost: 0, totalCostArchivo: 0 } },
                unitCostMismatches: [{ date: null, document: "F001", quantity: 1, expectedTotalCost: 261.8, foundTotalCost: 257, expectedUnitCost: 261.8, foundUnitCost: 257 }],
                expectedFinalBalance: { quantity: 1, unitCost: 257, totalCost: 257 },
                actualFinalBalance: { quantity: 1, unitCost: 261.8, totalCost: 261.8 },
                movementCount: 3,
                differences: ["Costo Unitario de Saldo Final"]
            }
        }]
    },
    RULE_014: {
        exporter: new Rule014Exporter(), lastColumn: "J", columns: 10,
        results: [{
            risk_level: "CRITICO",
            metadata: {
                month: 1,
                initialBalance: { quantity: 1, totalCost: 100 },
                totals: { entry: { quantity: 1, totalCost: 50 }, exit: { quantity: 1, totalCost: 30 } },
                expectedFinalBalance: { quantity: 1, totalCost: 120 },
                costTolerance: { percentage: 0, lowerLimit: 120, upperLimit: 120 },
                actualFinalBalance: { quantity: 1, totalCost: 125 },
                difference: { quantity: 0, totalCost: -5 },
                productCount: 4, movementCount: 9,
                differences: ["Costo valorizado fuera del rango permitido"]
            }
        }]
    }
};

describe("Modelo de reporte (reglas con encabezado estándar de 10 columnas)", () => {
    for (const [code, sample] of Object.entries(SAMPLES)) {
        describe(code, () => {
            let sheet: ExcelJS.Worksheet;
            beforeAll(async () => {
                sheet = await sheetOf(sample.exporter, sample.results);
            });

            it("título combinado en negrita en la fila 1", () => {
                expect(sheet.model.merges).toContain(`A1:${sample.lastColumn}1`);
                expect(sheet.getCell("A1").font?.bold).toBe(true);
                expect(String(sheet.getCell("A1").value)).toContain(code);
                expect(sheet.getCell("A1").alignment?.vertical).toBe("middle");
            });

            it("fila 2: Empresa / RUC / Periodo en negrita", () => {
                expect(sheet.getCell("A2").value).toBe("Empresa:");
                expect(sheet.getCell("B2").value).toBe("EXCLUSIVIDADES KIMBERLY E.I.R.L");
                expect(sheet.getCell("C2").value).toBe("RUC:");
                expect(sheet.getCell("D2").value).toBe("20493342062");
                expect(sheet.getCell("E2").value).toBe("Periodo:");
                expect(sheet.getCell("F2").value).toBe(2025);
                for (const address of ["A2", "B2", "C2", "D2", "E2", "F2"]) {
                    expect(sheet.getCell(address).font?.bold).toBe(true);
                }
            });

            it("cartel amarillo con letra roja 'AUDITORIA dd.mm.aaaa' en la última columna (día de Lima, no de UTC)", () => {
                const stamp = sheet.getCell(`${sample.lastColumn}2`);
                expect(stamp.value).toBe("AUDITORIA 29.09.2026");
                expect(stamp.font?.bold).toBe(true);
                expect(stamp.font?.size).toBe(14);
                expect(stamp.font?.color?.argb).toBe(RED);
                expect((stamp.fill as ExcelJS.FillPattern).fgColor?.argb).toBe(YELLOW);
                expect(stamp.alignment?.horizontal).toBe("center");
            });

            it("encabezado en la fila 3: negrita, centrado, texto ajustado y bordes finos", () => {
                const labels: ExcelJS.CellValue[] = [];
                for (let c = 1; c <= sample.columns; c++) {
                    const cell = sheet.getRow(3).getCell(c);
                    labels.push(cell.value);
                    expect(cell.font?.bold).toBe(true);
                    expect(cell.alignment?.horizontal).toBe("center");
                    expect(cell.alignment?.vertical).toBe("middle");
                    expect(cell.alignment?.wrapText).toBe(true);
                    expectThinBorders(cell);
                }
                expect(labels.slice(0, 3)).toEqual(["Periodo", "Código del producto", "Descripción del producto"]);
                expect(labels[sample.columns - 1]).toBe("Trazabilidad");
                expect(sheet.getRow(3).height).toBeGreaterThanOrEqual(30);
            });

            it("los datos empiezan en la fila 4 (sin fila vacía en el medio)", () => {
                expect(sheet.getRow(4).getCell(1).value).toBeTruthy();
                expect(sheet.getRow(4).getCell(1).value).not.toBe("Periodo");
            });

            it("celdas de datos: negrita, bordes finos, texto ajustado y alineación vertical al medio", () => {
                for (let c = 1; c <= sample.columns; c++) {
                    const cell = sheet.getRow(4).getCell(c);
                    expect(cell.font?.bold).toBe(true);
                    expect(cell.alignment?.wrapText).toBe(true);
                    expect(cell.alignment?.vertical).toBe("middle");
                    expectThinBorders(cell);
                }
            });

            it("panel congelado en la fila 3 y SIN autofiltro", () => {
                expect(sheet.views[0]?.state).toBe("frozen");
                expect((sheet.views[0] as any).ySplit).toBe(3);
                expect(sheet.autoFilter).toBeFalsy();
            });
        });
    }
});

describe("Modelo de reporte - detalles propios de cada regla", () => {
    it("RULE_008: título 20, información 14, anchos del modelo y columna de código centrada", async () => {
        const sheet = await sheetOf(SAMPLES.RULE_008.exporter, SAMPLES.RULE_008.results);
        expect(sheet.getCell("A1").font?.size).toBe(20);
        expect(sheet.getRow(1).height).toBe(26.25);
        expect(sheet.getCell("A2").font?.size).toBe(14);
        expect(sheet.getRow(3).height).toBe(37.5);
        expect(sheet.getCell("A4").font?.size).toBe(14);
        expect(sheet.getCell("J4").font?.size).toBe(11);
        expect(sheet.getColumn(3).width).toBeCloseTo(34.29, 1);
        expect(sheet.getColumn(4).width).toBeCloseTo(27, 1);
        expect(sheet.getColumn(10).width).toBeCloseTo(35.14, 1);
        expect(sheet.getCell("B4").alignment?.horizontal).toBe("center");
        expect(sheet.getCell("A4").alignment?.horizontal).toBeUndefined();
    });

    it("RULE_002: título 16 y letra 11 en los datos", async () => {
        const sheet = await sheetOf(SAMPLES.RULE_002.exporter, SAMPLES.RULE_002.results);
        expect(sheet.getCell("A1").font?.size).toBe(16);
        expect(sheet.getCell("A4").font?.size).toBe(11);
        expect(sheet.getRow(3).height).toBe(45);
    });

    it("RULE_013: montos (esperado / encontrado / diferencia) con formato #,##0.00", async () => {
        const sheet = await sheetOf(SAMPLES.RULE_013.exporter, SAMPLES.RULE_013.results);
        expect(sheet.getCell("E4").numFmt).toBe("#,##0.00");
        expect(sheet.getCell("F4").numFmt).toBe("#,##0.00");
        expect(sheet.getCell("G4").numFmt).toBe("#,##0.00");
        expect(sheet.getCell("A4").numFmt).not.toBe("#,##0.00");
        expect(sheet.getCell("A4").font?.size).toBe(14);
        expect(sheet.getCell("J4").font?.size).toBe(12);
    });

    it("RULE_014: el nombre de la empresa va en letra 12 (el resto de la fila 2, en 14), como en el modelo", async () => {
        const sheet = await sheetOf(SAMPLES.RULE_014.exporter, SAMPLES.RULE_014.results);
        expect(sheet.getCell("A2").font?.size).toBe(14);
        expect(sheet.getCell("B2").font?.size).toBe(12);
    });

    it("RULE_014: la trazabilidad en colores conserva el tamaño 12 en cada tramo de texto", async () => {
        const sheet = await sheetOf(SAMPLES.RULE_014.exporter, SAMPLES.RULE_014.results);
        const value = sheet.getCell("J4").value as ExcelJS.CellRichTextValue;
        expect(value.richText.length).toBeGreaterThan(5);
        for (const run of value.richText) {
            expect(run.font?.size).toBe(12);
            expect(run.font?.bold).toBe(true);
        }
    });

    it("RULE_003: solo la columna Valor esperado lleva formato de monto", async () => {
        const sheet = await sheetOf(SAMPLES.RULE_003.exporter, SAMPLES.RULE_003.results);
        expect(sheet.getCell("E4").numFmt).toBe("#,##0.00");
        expect(sheet.getCell("F4").numFmt).not.toBe("#,##0.00");
    });
});

describe("Modelo de reporte - RULE_001 (doble encabezado)", () => {
    const results = [{
        product_code: "78700", product_name: "ZAPATO CALIMOD", risk_level: "CRITICO", error_type: "INVENTORY_MISMATCH",
        metadata: {
            inventoryCode: "78700", normalizedCode: "78700",
            inventoryStock: 1, kardexStock: 1,
            inventoryUnitCost: 194.8, kardexUnitCost: 0,
            inventoryTotalCost: 194.8, kardexTotalCost: 194.8,
            kardexMovements: 2
        }
    }];

    it("fila 3 con el encabezado estándar y fila 4 con la descripción de Valor esperado / encontrado", async () => {
        const sheet = await sheetOf(new Rule001Exporter(), results);
        expect(sheet.getCell("E3").value).toBe("Valor esperado");
        expect(sheet.getCell("F3").value).toBe("Valor encontrado");
        expect(sheet.getCell("E4").value).toBe("Valor esperado del inventario final");
        expect(sheet.getCell("F4").value).toBe("Valor encontrado del Kardex");
        expect(sheet.getRow(4).height).toBe(45);
        for (let c = 1; c <= 10; c++) {
            const cell = sheet.getRow(4).getCell(c);
            expect(cell.font?.bold).toBe(true);
            expect(cell.alignment?.horizontal).toBe("center");
            expectThinBorders(cell);
        }
    });

    it("los datos empiezan en la fila 5; congelado en la 4 y autofiltro A4:J4", async () => {
        const sheet = await sheetOf(new Rule001Exporter(), results);
        expect(sheet.getCell("A5").value).toBe("Enero");
        expect(sheet.getCell("B5").value).toBe("78700");
        expect((sheet.views[0] as any).ySplit).toBe(4);
        expect(sheet.autoFilter).toBe("A4:J4");
    });

    it("anchos del modelo (Valor esperado / encontrado = 15) y título 18", async () => {
        const sheet = await sheetOf(new Rule001Exporter(), results);
        expect(sheet.getColumn(5).width).toBeCloseTo(15, 1);
        expect(sheet.getColumn(6).width).toBeCloseTo(15, 1);
        expect(sheet.getColumn(10).width).toBeCloseTo(41.14, 1);
        expect(sheet.getCell("A1").font?.size).toBe(18);
        expect(sheet.getCell("A5").font?.bold).toBe(true);
        expectThinBorders(sheet.getCell("A5"));
    });
});

describe("Modelo de reporte - RULE_004 (16 columnas)", () => {
    const results = [{
        risk_level: "BAJO", error_type: "ACCEPTED",
        metadata: {
            transitItem: "2024-12-06", issueDate: "2024-12-06", warehouseDate: "2025-01-04",
            supplierRuc: "20532145415", supplier: "KS DEPOR S.A.", document: "FAC-F016-123786", normalizedDocument: "F01600123786",
            month: "1", expectedCost: 28102.89, foundCost: 28640.64, differencePercent: 1.91,
            isIncident: false, thresholdPercent: 5, usedFallback: false,
            evaluatedProducts: [{ code: "92222", description: "ZAPATILLA CATERPILLAR", cost: 10 }]
        }
    }];

    it("título combinado A1:P1, cartel en P2, encabezado de 16 columnas en la fila 3 y datos desde la fila 4", async () => {
        const sheet = await sheetOf(new Rule004Exporter(), results);
        expect(sheet.model.merges).toContain("A1:P1");
        expect(sheet.getCell("P2").value).toBe("AUDITORIA 29.09.2026");
        expect(sheet.getCell("A3").value).toBe("Periodo");
        expect(sheet.getCell("J3").value).toBe("Tipo de inconsistencia");
        expect(sheet.getCell("P3").value).toBe("Trazabilidad");
        expect(sheet.getRow(3).height).toBe(47.25);
        expect(sheet.getCell("A4").value).toBe("Enero");
        expect(sheet.getCell("F4").value).toBe("FAC-F016-123786");
        expect((sheet.views[0] as any).ySplit).toBe(3);
        expect(sheet.autoFilter).toBeFalsy();
    });

    it("anchos del modelo, descripción y trazabilidad en letra 10, montos con #,##0.00", async () => {
        const sheet = await sheetOf(new Rule004Exporter(), results);
        expect(sheet.getColumn(9).width).toBeCloseTo(59, 1);
        expect(sheet.getColumn(16).width).toBeCloseTo(76, 1);
        expect(sheet.getCell("I4").font?.size).toBe(10);
        expect(sheet.getCell("P4").font?.size).toBe(10);
        expect(sheet.getCell("A4").font?.size).toBe(12);
        expect(sheet.getCell("K4").numFmt).toBe("#,##0.00");
        expect(sheet.getCell("L4").numFmt).toBe("#,##0.00");
        expect(sheet.getCell("M4").numFmt).toBe("#,##0.00");
        for (let c = 1; c <= 16; c++) {
            const cell = sheet.getRow(4).getCell(c);
            expect(cell.font?.bold).toBe(true);
            expect(cell.alignment?.wrapText).toBe(true);
            expectThinBorders(cell);
        }
    });
});
