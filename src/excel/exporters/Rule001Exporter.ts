import ExcelJS from "exceljs";

import { BaseExcelExporter } from "./base/BaseExcelExporter";
import { ReportHeader } from "./base/ReportHeader";
import { AuditFindingRow } from "./base/AuditFindingRow";
import { DateUtils } from "../helpers/dateutils";

export interface Rule001Metadata {
    month?: number;
    inventoryCode?: string;
    normalizedCode?: string;
    inventoryStock?: number;
    kardexStock?: number;
    inventoryUnitCost?: number;
    kardexUnitCost?: number;
    inventoryTotalCost?: number;
    kardexTotalCost?: number;
    kardexMovements?: number;
}

export class Rule001Exporter extends BaseExcelExporter {

    private static equals(a: number, b: number): boolean {
        return Math.abs(a - b) < 0.01;
    }

    private static puntoDePartida(auditYear?: number): string {
        const previousYear = (auditYear ?? new Date().getFullYear()) - 1;
        return `Punto de partida: Inventario final del cierre del ejercicio anterior (dic ${previousYear})`;
    }

    async export(
        results: any[],
        header: ReportHeader
    ) {

        const workbook = new ExcelJS.Workbook();

        workbook.creator = "HM Kardex Audit";
        workbook.created = new Date();

        const worksheet =
            workbook.addWorksheet("RULE_001");

        this.writeHeader(
            worksheet,
            "RULE_001 - Validación y comparacion de cantidades y costos del inventario al cierre del año",
            header,
            "J"
        );

        this.writeTableHeader(worksheet);

        const findings =
            this.expand(results, header.year);

        this.writeRows(
            worksheet,
            findings
        );

        worksheet.views = [
            {
                state: "frozen",
                ySplit: 4
            }
        ];

        worksheet.autoFilter = {
            from: "A4",
            to: "J4"
        };

        return workbook;
    }

    protected writeTableHeader(worksheet: ExcelJS.Worksheet) {

        super.writeTableHeader(worksheet);

        const headerRow = worksheet.getRow(4);
        headerRow.getCell(5).value = "Valor esperado del inventario final";
        headerRow.getCell(6).value = "Valor encontrado del Kardex";

        worksheet.getColumn(5).width = 30;
        worksheet.getColumn(6).width = 30;
    }

    private expand(
        results: any[],
        auditYear?: number
    ): AuditFindingRow[] {

        const rows: AuditFindingRow[] = [];

        const puntoDePartida =
            Rule001Exporter.puntoDePartida(auditYear);

        for (const result of results) {

            const metadata =
                result.metadata as Rule001Metadata;

            const errorType =
                result.error_type;

            const month =
                result.month && result.month > 0
                    ? DateUtils.monthName(result.month)
                    : "Sin período";

            /*
             * ========================================================
             * PRODUCTO NO ENCONTRADO
             * ========================================================
             */

            if (errorType === "PRODUCT_NOT_FOUND") {

                rows.push({
                    period: month,

                    productCode:
                        result.product_code ?? "",

                    productDescription:
                        result.product_name ?? "",

                    inconsistencyType:
                        "Producto no encontrado en Kardex",

                    expectedValue:
                        "Producto registrado en Inventario",

                    foundValue:
                        "Producto no encontrado",

                    difference:
                        0,

                    differencePercent:
                        0,

                    riskLevel:
                        result.risk_level,

                    traceability: [
                        puntoDePartida,
                        `Descripción: ${result.description}`,
                        `Recomendación: ${result.recommendation}`,
                        `Código Inventario: ${result.product_code ?? ""}`
                    ].join("\n")
                });

                continue;
            }

            /*
             * ========================================================
             * SIN OPERACIONES CON CANTIDAD
             * ========================================================
             *
             * Esto puede existir en auditorías antiguas.
             *
             * Ya NO debería generarse desde Rule001,
             * pero si existe en DB lo mostramos sin intentar
             * hacer cálculos numéricos.
             */

            if (errorType === "WITHOUT_MOVEMENTS") {

                rows.push({
                    period: "Sin período",

                    productCode:
                        result.product_code ?? "",

                    productDescription:
                        result.product_name ?? "",

                    inconsistencyType:
                        "Sin operaciones con cantidad",

                    expectedValue:
                        "No aplica",

                    foundValue:
                        "No aplica",

                    difference:
                        0,

                    differencePercent:
                        0,

                    riskLevel:
                        result.risk_level,

                    traceability: [
                        puntoDePartida,
                        `Descripción: ${result.description}`,
                        `Recomendación: ${result.recommendation}`,
                        `Código Inventario: ${result.product_code ?? ""}`,
                        "Este registro no participa en la comparación porque no existen operaciones con cantidad."
                    ].join("\n")
                });

                continue;
            }

            /*
             * ========================================================
             * INVENTORY MISMATCH
             * ========================================================
             */

            if (errorType === "INVENTORY_MISMATCH") {

                const inventoryStock = metadata.inventoryStock ?? 0;
                const kardexStock = metadata.kardexStock ?? 0;
                const inventoryUnitCost = metadata.inventoryUnitCost ?? 0;
                const kardexUnitCost = metadata.kardexUnitCost ?? 0;
                const inventoryTotalCost = metadata.inventoryTotalCost ?? 0;
                const kardexTotalCost = metadata.kardexTotalCost ?? 0;

                const stockDiffiere =
                    !Rule001Exporter.equals(inventoryStock, kardexStock);

                const costoUnitarioDiffiere =
                    !Rule001Exporter.equals(inventoryUnitCost, kardexUnitCost);

                const costoTotalDiffiere =
                    !Rule001Exporter.equals(inventoryTotalCost, kardexTotalCost);

                /*
                 * STOCK
                 */

                if (stockDiffiere) {

                    rows.push({
                        period: month,

                        productCode:
                            result.product_code,

                        productDescription:
                            result.product_name,

                        inconsistencyType:
                            "Cantidad",

                        expectedValue:
                            inventoryStock,

                        foundValue:
                            kardexStock,

                        difference:
                            inventoryStock - kardexStock,

                        differencePercent:
                            inventoryStock === 0
                                ? 0
                                : Math.abs(
                                    (inventoryStock - kardexStock) /
                                    inventoryStock
                                ) * 100,

                        riskLevel:
                            result.risk_level,

                        traceability: [
                            puntoDePartida,
                            "Descripción: El inventario final no coincide con el Saldo Inicial del Kardex (Cantidad).",
                            `Código Inventario: ${metadata.inventoryCode ?? ""}`,
                            `Código Normalizado: ${metadata.normalizedCode ?? ""}`,
                            `Stock Inventario: ${inventoryStock}`,
                            `Stock Kardex: ${kardexStock}`,
                            `Movimientos Kardex: ${metadata.kardexMovements ?? 0}`
                        ].join("\n")
                    });
                }

                /*
                 * COSTO UNITARIO
                 */

                if (costoUnitarioDiffiere) {

                    rows.push({
                        period: month,

                        productCode:
                            result.product_code,

                        productDescription:
                            result.product_name,

                        inconsistencyType:
                            "Costo Unitario",

                        expectedValue:
                            inventoryUnitCost,

                        foundValue:
                            kardexUnitCost,

                        difference:
                            inventoryUnitCost - kardexUnitCost,

                        differencePercent:
                            inventoryUnitCost === 0
                                ? 0
                                : Math.abs(
                                    (inventoryUnitCost - kardexUnitCost) /
                                    inventoryUnitCost
                                ) * 100,

                        riskLevel:
                            result.risk_level,

                        traceability: [
                            puntoDePartida,
                            "Descripción: El inventario final no coincide con el Saldo Inicial del Kardex (Costo Unitario).",
                            `Costo Inventario: ${inventoryUnitCost}`,
                            `Costo Kardex: ${kardexUnitCost}`,
                            `Movimientos Kardex: ${metadata.kardexMovements ?? 0}`
                        ].join("\n")
                    });
                }

                /*
                 * COSTO TOTAL
                 */

                if (costoTotalDiffiere) {

                    rows.push({
                        period: month,

                        productCode:
                            result.product_code,

                        productDescription:
                            result.product_name,

                        inconsistencyType:
                            "Costo Total",

                        expectedValue:
                            inventoryTotalCost,

                        foundValue:
                            kardexTotalCost,

                        difference:
                            inventoryTotalCost - kardexTotalCost,

                        differencePercent:
                            inventoryTotalCost === 0
                                ? 0
                                : Math.abs(
                                    (inventoryTotalCost - kardexTotalCost) /
                                    inventoryTotalCost
                                ) * 100,

                        riskLevel:
                            result.risk_level,

                        traceability: [
                            puntoDePartida,
                            "Descripción: El inventario final no coincide con el Saldo Inicial del Kardex (Costo Total).",
                            `Total Inventario: ${inventoryTotalCost}`,
                            `Total Kardex: ${kardexTotalCost}`,
                            `Movimientos Kardex: ${metadata.kardexMovements ?? 0}`
                        ].join("\n")
                    });
                }
            }
        }

        return rows;
    }
}
