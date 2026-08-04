import ExcelJS from "exceljs";

import { BaseExcelExporter } from "./base/BaseExcelExporter";
import { ReportHeader } from "./base/ReportHeader";
import { AuditFindingRow } from "./base/AuditFindingRow";
import { DateUtils } from "../helpers/dateutils";

export interface Rule013Metadata {
    month: number;
    normalizedCode: string;
    initialBalance: {
        quantity: number;
        totalCost: number;
    };
    totals: {
        entry: {
            quantity: number;
            totalCost: number;
        };
        exit: {
            quantity: number;
            totalCost: number;
        };
    };
    expectedFinalBalance: {
        quantity: number;
        totalCost: number;
    };
    costTolerance: {
        percentage: number;
        lowerLimit: number;
        upperLimit: number;
    };
    actualFinalBalance: {
        quantity: number;
        totalCost: number;
    };
    difference: {
        quantity: number;
        totalCost: number;
    };
    movementCount: number;
    differences: string[];
}

export class Rule013Exporter extends BaseExcelExporter {
    async export(
        results: any[],
        header: ReportHeader
    ) {
        const workbook = new ExcelJS.Workbook();
        workbook.creator = "HM Kardex Audit";
        workbook.created = new Date();
        const worksheet = workbook.addWorksheet("RULE_013");
        this.writeHeader(
            worksheet,
            "RULE_013 - Validación de Sumatorias Mensuales",
            header,
            "J"
        );
        this.writeTableHeader(worksheet);
        const findings =
            this.buildFindings(results);
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

    private buildFindings( results: any[] ): AuditFindingRow[] {
        const rows: AuditFindingRow[] = [];
        for (const result of results) {
            const metadata =
                result.metadata as Rule013Metadata;
            rows.push({
                period: DateUtils.monthName(metadata.month),
                productCode: result.product_code,
                productDescription: result.product_name,
                inconsistencyType: "Sumatoria Mensual - Cantidad",
                expectedValue: metadata.expectedFinalBalance.quantity,
                foundValue: metadata.actualFinalBalance.quantity,
                difference: metadata.difference.quantity,
                differencePercent:
                    metadata.expectedFinalBalance.quantity === 0
                        ? 0
                        : Math.abs(
                            metadata.difference.quantity /
                            metadata.expectedFinalBalance.quantity
                        ) * 100,
                riskLevel: result.risk_level,
                traceability: [
                    `Código Normalizado: ${metadata.normalizedCode}`,
                    `Saldo Inicial: ${metadata.initialBalance.quantity}`,
                    `Entradas: ${metadata.totals.entry.quantity}`,
                    `Salidas: ${metadata.totals.exit.quantity}`,
                    `Saldo Esperado: ${metadata.expectedFinalBalance.quantity}`,
                    `Saldo Real: ${metadata.actualFinalBalance.quantity}`,
                    `Movimientos Analizados: ${metadata.movementCount}`,
                    `Campos con diferencia: ${metadata.differences.join(", ")}`
                ].join("\n")
            });

            /*
             * =====================================================
             * COSTO VALORIZADO
             * =====================================================
             */

            rows.push({
                period: DateUtils.monthName(metadata.month),
                productCode: result.product_code,
                productDescription: result.product_name,
                inconsistencyType: "Costo Valorizado Mensual",
                expectedValue: metadata.expectedFinalBalance.totalCost,
                foundValue: metadata.actualFinalBalance.totalCost,
                difference: metadata.difference.totalCost,
                differencePercent:
                    metadata.expectedFinalBalance.totalCost === 0
                        ? 0
                        : Math.abs(
                            metadata.difference.totalCost /
                            metadata.expectedFinalBalance.totalCost
                        ) * 100,
                riskLevel: result.risk_level,
                traceability: [
                    `Código Normalizado: ${metadata.normalizedCode}`,
                    `Saldo Inicial: ${metadata.initialBalance.totalCost}`,
                    `Entradas: ${metadata.totals.entry.totalCost}`,
                    `Salidas: ${metadata.totals.exit.totalCost}`,
                    `Costo Esperado: ${metadata.expectedFinalBalance.totalCost}`,
                    `Costo Real: ${metadata.actualFinalBalance.totalCost}`,
                    `Rango Permitido (${metadata.costTolerance.percentage}%): ${metadata.costTolerance.lowerLimit} - ${metadata.costTolerance.upperLimit}`,
                    `Movimientos Analizados: ${metadata.movementCount}`,
                    `Campos con diferencia: ${metadata.differences.join(", ")}`
                ].join("\n")
            });
        }
        return rows;
    }
}