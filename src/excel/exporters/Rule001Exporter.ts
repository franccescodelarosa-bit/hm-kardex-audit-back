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
            this.expand(results);

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

    private expand(
        results: any[]
    ): AuditFindingRow[] {

        const rows: AuditFindingRow[] = [];

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

                /*
                 * STOCK
                 */

                rows.push({
                    period: month,

                    productCode:
                        result.product_code,

                    productDescription:
                        result.product_name,

                    inconsistencyType:
                        "Cantidad",

                    expectedValue:
                        metadata.inventoryStock ?? 0,

                    foundValue:
                        metadata.kardexStock ?? 0,

                    difference:
                        (metadata.inventoryStock ?? 0) -
                        (metadata.kardexStock ?? 0),

                    differencePercent:
                        metadata.inventoryStock === 0
                            ? 0
                            : Math.abs(
                                (
                                    (metadata.inventoryStock ?? 0) -
                                    (metadata.kardexStock ?? 0)
                                ) /
                                metadata.inventoryStock!
                            ) * 100,

                    riskLevel:
                        result.risk_level,

                    traceability: [
                        `Descripción: ${result.description}`,
                        `Código Inventario: ${metadata.inventoryCode ?? ""}`,
                        `Código Normalizado: ${metadata.normalizedCode ?? ""}`,
                        `Stock Inventario: ${metadata.inventoryStock ?? 0}`,
                        `Stock Kardex: ${metadata.kardexStock ?? 0}`,
                        `Movimientos Kardex: ${metadata.kardexMovements ?? 0}`
                    ].join("\n")
                });

                /*
                 * COSTO UNITARIO
                 */

                rows.push({
                    period: month,

                    productCode:
                        result.product_code,

                    productDescription:
                        result.product_name,

                    inconsistencyType:
                        "Costo Unitario",

                    expectedValue:
                        metadata.inventoryUnitCost ?? 0,

                    foundValue:
                        metadata.kardexUnitCost ?? 0,

                    difference:
                        (metadata.inventoryUnitCost ?? 0) -
                        (metadata.kardexUnitCost ?? 0),

                    differencePercent:
                        metadata.inventoryUnitCost === 0
                            ? 0
                            : Math.abs(
                                (
                                    (metadata.inventoryUnitCost ?? 0) -
                                    (metadata.kardexUnitCost ?? 0)
                                ) /
                                metadata.inventoryUnitCost!
                            ) * 100,

                    riskLevel:
                        result.risk_level,

                    traceability: [
                        `Descripción: ${result.description}`,
                        `Costo Inventario: ${metadata.inventoryUnitCost ?? 0}`,
                        `Costo Kardex: ${metadata.kardexUnitCost ?? 0}`,
                        `Movimientos Kardex: ${metadata.kardexMovements ?? 0}`
                    ].join("\n")
                });

                /*
                 * COSTO TOTAL
                 */

                rows.push({
                    period: month,

                    productCode:
                        result.product_code,

                    productDescription:
                        result.product_name,

                    inconsistencyType:
                        "Costo Total",

                    expectedValue:
                        metadata.inventoryTotalCost ?? 0,

                    foundValue:
                        metadata.kardexTotalCost ?? 0,

                    difference:
                        (metadata.inventoryTotalCost ?? 0) -
                        (metadata.kardexTotalCost ?? 0),

                    differencePercent:
                        metadata.inventoryTotalCost === 0
                            ? 0
                            : Math.abs(
                                (
                                    (metadata.inventoryTotalCost ?? 0) -
                                    (metadata.kardexTotalCost ?? 0)
                                ) /
                                metadata.inventoryTotalCost!
                            ) * 100,

                    riskLevel:
                        result.risk_level,

                    traceability: [
                        `Descripción: ${result.description}`,
                        `Total Inventario: ${metadata.inventoryTotalCost ?? 0}`,
                        `Total Kardex: ${metadata.kardexTotalCost ?? 0}`,
                        `Movimientos Kardex: ${metadata.kardexMovements ?? 0}`
                    ].join("\n")
                });
            }
        }

        return rows;
    }
}