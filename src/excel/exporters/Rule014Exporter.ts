import ExcelJS from "exceljs";

import { BaseExcelExporter } from "./base/BaseExcelExporter";
import { ReportHeader } from "./base/ReportHeader";
import { AuditFindingRow } from "./base/AuditFindingRow";
import { DateUtils } from "../helpers/dateutils";
export interface Rule014Metadata {
    month: number;
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
    /** Inventario Valorizado de CIERRE real (suma de saldos finales). El diagrama oficial de Regla 14 lo llama "valor esperado". */
    actualFinalBalance: {
        quantity: number;
        totalCost: number;
    };
    difference: {
        quantity: number;
        totalCost: number;
    };
    productCount: number;
    movementCount: number;
    differences: string[];
}

export class Rule014Exporter extends BaseExcelExporter {
    async export(
        results: any[],
        header: ReportHeader
    ) {
        const workbook = new ExcelJS.Workbook();
        workbook.creator = "HM Kardex Audit";
        workbook.created = new Date();
        const worksheet =
            workbook.addWorksheet("RULE_014");
        this.writeHeader(
            worksheet,
            "RULE_014 - Validación Consolidada de Sumatorias Mensuales",
            header,
            "J"
        );
        this.writeTableHeader(worksheet);
        const findings = this.buildFindings(results);
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

    private static round(value: number): number {
        return Math.round(value * 100) / 100;
    }

    private static percent(expected: number, difference: number): number {
        return expected === 0
            ? 0
            : Math.abs(difference / expected) * 100;
    }

    private buildFindings( results: any[] ): AuditFindingRow[] {
        const rows: AuditFindingRow[] = [];
        for (const result of results) {
            const metadata = result.metadata as Rule014Metadata;
            const period = DateUtils.monthName(metadata.month);
            const productosYMovimientos = [
                `Productos Consolidados: ${metadata.productCount}`,
                `Movimientos Consolidados: ${metadata.movementCount}`,
                `Campos con diferencia: ${metadata.differences.join(", ")}`
            ];

            if (metadata.differences.includes("Cantidad")) {
                // Esperado (diagrama) = Cierre real. Encontrado (diagrama) = Fórmula.
                const diferencia = Rule014Exporter.round(
                    metadata.actualFinalBalance.quantity - metadata.expectedFinalBalance.quantity
                );
                rows.push({
                    period,
                    productCode: "CONSOLIDADO",
                    productDescription: "Consolidado Mensual",
                    inconsistencyType: "Sumatoria Consolidada - Cantidad",
                    expectedValue: metadata.actualFinalBalance.quantity,
                    foundValue: metadata.expectedFinalBalance.quantity,
                    difference: diferencia,
                    differencePercent: Rule014Exporter.percent(
                        metadata.actualFinalBalance.quantity,
                        diferencia
                    ),
                    riskLevel: result.risk_level,
                    traceability: [
                        `Saldo Inicial: ${metadata.initialBalance.quantity}`,
                        `Entradas: ${metadata.totals.entry.quantity}`,
                        `Salidas: ${metadata.totals.exit.quantity}`,
                        `Inventario Valorizado de Cierre (Esperado): ${metadata.actualFinalBalance.quantity}`,
                        `Resultado Fórmula Inicio+Entrada-Salida (Encontrado): ${metadata.expectedFinalBalance.quantity}`,
                        ...productosYMovimientos
                    ].join("\n")
                });
            }

            if (metadata.differences.includes("Costo valorizado fuera del rango permitido")) {
                const diferencia = Rule014Exporter.round(
                    metadata.actualFinalBalance.totalCost - metadata.expectedFinalBalance.totalCost
                );
                rows.push({
                    period,
                    productCode: "CONSOLIDADO",
                    productDescription: "Consolidado Mensual",
                    inconsistencyType: "Costo Valorizado Consolidado",
                    expectedValue: metadata.actualFinalBalance.totalCost,
                    foundValue: metadata.expectedFinalBalance.totalCost,
                    difference: diferencia,
                    differencePercent: Rule014Exporter.percent(
                        metadata.actualFinalBalance.totalCost,
                        diferencia
                    ),
                    riskLevel: result.risk_level,
                    traceability: [
                        `Saldo Inicial: ${metadata.initialBalance.totalCost}`,
                        `Entradas: ${metadata.totals.entry.totalCost}`,
                        `Salidas: ${metadata.totals.exit.totalCost}`,
                        `Inventario Valorizado de Cierre (Esperado): ${metadata.actualFinalBalance.totalCost}`,
                        `Resultado Fórmula Inicio+Entrada-Salida (Encontrado): ${metadata.expectedFinalBalance.totalCost}`,
                        `Rango Permitido (${metadata.costTolerance.percentage}%): ${metadata.costTolerance.lowerLimit} - ${metadata.costTolerance.upperLimit}`,
                        ...productosYMovimientos
                    ].join("\n")
                });
            }
        }
        return rows;
    }
}
