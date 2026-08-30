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
            "RULE_013 - Validación de Sumatorias MensualesValidacion del costo promedio ponderado - CPP por codigo",
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

            if (metadata.differences.includes("Costo Total de Salidas")) {
                const diferencia = Rule013Exporter.round(
                    metadata.totals.exit.totalCostArchivo - metadata.totals.exit.totalCost
                );
                rows.push({
                    period,
                    productCode: result.product_code,
                    productDescription: result.product_name,
                    inconsistencyType: "Costo Total de Salidas",
                    // Según el diagrama oficial ("valor esperado: Costo total")
                    // y el Anexo 02: esperado = lo que dice el ARCHIVO,
                    // encontrado = lo que la regla calculó con el CPP.
                    expectedValue: metadata.totals.exit.totalCostArchivo,
                    foundValue: metadata.totals.exit.totalCost,
                    difference: diferencia,
                    differencePercent: Rule013Exporter.percent(
                        metadata.totals.exit.totalCostArchivo,
                        diferencia
                    ),
                    riskLevel: result.risk_level,
                    traceability: [
                        codigoNormalizado,
                        `Cantidad de Salidas: ${metadata.totals.exit.quantity}`,
                        `Costo de Salidas Esperado (Archivo): ${metadata.totals.exit.totalCostArchivo}`,
                        `Costo de Salidas Encontrado (CPP): ${metadata.totals.exit.totalCost}`,
                        `Movimientos Analizados: ${metadata.movementCount}`,
                        // Solo la diferencia de ESTA fila -- no la lista
                        // completa del producto (evita mezclar con las otras
                        // filas, mismo criterio que RULE_003).
                        `Campos con diferencia: Costo Total de Salidas`
                    ].join("\n")
                });
            }

            if (metadata.differences.includes("Costo Total de Saldo Final")) {
                const diferencia = Rule013Exporter.round(-metadata.difference.totalCost);
                rows.push({
                    period,
                    productCode: result.product_code,
                    productDescription: result.product_name,
                    // Antes decía "Costo Valorizado Mensual" -- nombre
                    // distinto al de la diferencia real ("Costo Total de
                    // Saldo Final"), lo que no coincidía con "Campos con
                    // diferencia". Unificado al mismo nombre.
                    inconsistencyType: "Costo Total de Saldo Final",
                    expectedValue: metadata.actualFinalBalance.totalCost,
                    foundValue: metadata.expectedFinalBalance.totalCost,
                    difference: diferencia,
                    differencePercent: Rule013Exporter.percent(
                        metadata.actualFinalBalance.totalCost,
                        diferencia
                    ),
                    riskLevel: result.risk_level,
                    traceability: [
                        codigoNormalizado,
                        `Saldo Inicial: ${metadata.initialBalance.totalCost}`,
                        `Entradas: ${metadata.totals.entry.totalCost}`,
                        `Salidas: ${metadata.totals.exit.totalCost}`,
                        `Costo Esperado (Archivo): ${metadata.actualFinalBalance.totalCost}`,
                        `Costo Encontrado (CPP recalculado): ${metadata.expectedFinalBalance.totalCost}`,
                        `Rango Permitido (${metadata.costTolerance.percentage}%): ${metadata.costTolerance.lowerLimit} - ${metadata.costTolerance.upperLimit}`,
                        `Movimientos Analizados: ${metadata.movementCount}`,
                        `Campos con diferencia: Costo Total de Saldo Final`
                    ].join("\n")
                });
            }

            if (metadata.differences.includes("Costo Unitario de Saldo Final")) {
                const diferencia = Rule013Exporter.round(-metadata.difference.unitCost);
                rows.push({
                    period,
                    productCode: result.product_code,
                    productDescription: result.product_name,
                    inconsistencyType: "Costo Unitario de Saldo Final",
                    expectedValue: metadata.actualFinalBalance.unitCost,
                    foundValue: metadata.expectedFinalBalance.unitCost,
                    difference: diferencia,
                    differencePercent: Rule013Exporter.percent(
                        metadata.actualFinalBalance.unitCost,
                        diferencia
                    ),
                    riskLevel: result.risk_level,
                    traceability: [
                        codigoNormalizado,
                        `Costo Unitario del Archivo: ${metadata.actualFinalBalance.unitCost}`,
                        `CPP Recalculado: ${metadata.expectedFinalBalance.unitCost}`,
                        `Movimientos Analizados: ${metadata.movementCount}`,
                        `Campos con diferencia: Costo Unitario de Saldo Final`
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
