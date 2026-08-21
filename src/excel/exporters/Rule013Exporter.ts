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
            totalCostArchivo: number;
        };
    };
    expectedFinalBalance: {
        quantity: number;
        unitCost: number;
        totalCost: number;
    };
    costTolerance: {
        percentage: number;
        lowerLimit: number;
        upperLimit: number;
    };
    actualFinalBalance: {
        quantity: number;
        unitCost: number;
        totalCost: number;
    };
    difference: {
        quantity: number;
        unitCost: number;
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

    private static percent(expected: number, difference: number): number {
        return expected === 0
            ? 0
            : Math.abs(difference / expected) * 100;
    }

    private buildFindings( results: any[] ): AuditFindingRow[] {
        const rows: AuditFindingRow[] = [];
        for (const result of results) {
            const metadata =
                result.metadata as Rule013Metadata;
            const period = DateUtils.monthName(metadata.month);
            const codigoNormalizado = `Código Normalizado: ${metadata.normalizedCode}`;
            const camposConDiferencia = `Campos con diferencia: ${metadata.differences.join(", ")}`;

            if (metadata.differences.includes("Costo Total de Salidas")) {
                rows.push({
                    period,
                    productCode: result.product_code,
                    productDescription: result.product_name,
                    inconsistencyType: "Costo Total de Salidas",
                    expectedValue: metadata.totals.exit.totalCost,
                    foundValue: metadata.totals.exit.totalCostArchivo,
                    difference: Rule013Exporter.round(
                        metadata.totals.exit.totalCost - metadata.totals.exit.totalCostArchivo
                    ),
                    differencePercent: Rule013Exporter.percent(
                        metadata.totals.exit.totalCost,
                        metadata.totals.exit.totalCost - metadata.totals.exit.totalCostArchivo
                    ),
                    riskLevel: result.risk_level,
                    traceability: [
                        codigoNormalizado,
                        `Cantidad de Salidas: ${metadata.totals.exit.quantity}`,
                        `Costo de Salidas Recalculado (CPP): ${metadata.totals.exit.totalCost}`,
                        `Costo de Salidas del Archivo: ${metadata.totals.exit.totalCostArchivo}`,
                        `Movimientos Analizados: ${metadata.movementCount}`,
                        camposConDiferencia
                    ].join("\n")
                });
            }

            if (metadata.differences.includes("Costo Total de Saldo Final")) {
                rows.push({
                    period,
                    productCode: result.product_code,
                    productDescription: result.product_name,
                    inconsistencyType: "Costo Valorizado Mensual",
                    expectedValue: metadata.expectedFinalBalance.totalCost,
                    foundValue: metadata.actualFinalBalance.totalCost,
                    difference: metadata.difference.totalCost,
                    differencePercent: Rule013Exporter.percent(
                        metadata.expectedFinalBalance.totalCost,
                        metadata.difference.totalCost
                    ),
                    riskLevel: result.risk_level,
                    traceability: [
                        codigoNormalizado,
                        `Saldo Inicial: ${metadata.initialBalance.totalCost}`,
                        `Entradas: ${metadata.totals.entry.totalCost}`,
                        `Salidas: ${metadata.totals.exit.totalCost}`,
                        `Costo Esperado: ${metadata.expectedFinalBalance.totalCost}`,
                        `Costo Real: ${metadata.actualFinalBalance.totalCost}`,
                        `Rango Permitido (${metadata.costTolerance.percentage}%): ${metadata.costTolerance.lowerLimit} - ${metadata.costTolerance.upperLimit}`,
                        `Movimientos Analizados: ${metadata.movementCount}`,
                        camposConDiferencia
                    ].join("\n")
                });
            }

            if (metadata.differences.includes("Costo Unitario de Saldo Final")) {
                rows.push({
                    period,
                    productCode: result.product_code,
                    productDescription: result.product_name,
                    inconsistencyType: "Costo Unitario de Saldo Final",
                    expectedValue: metadata.expectedFinalBalance.unitCost,
                    foundValue: metadata.actualFinalBalance.unitCost,
                    difference: metadata.difference.unitCost,
                    differencePercent: Rule013Exporter.percent(
                        metadata.expectedFinalBalance.unitCost,
                        metadata.difference.unitCost
                    ),
                    riskLevel: result.risk_level,
                    traceability: [
                        codigoNormalizado,
                        `CPP Recalculado: ${metadata.expectedFinalBalance.unitCost}`,
                        `Costo Unitario del Archivo: ${metadata.actualFinalBalance.unitCost}`,
                        `Movimientos Analizados: ${metadata.movementCount}`,
                        camposConDiferencia
                    ].join("\n")
                });
            }
        }
        return rows;
    }

    private static round(value: number): number {
        return Math.round(value * 100) / 100;
    }
}
