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

    unitCostMismatches: {
        date: string | Date | null;
        document: string;
        quantity: number;
        expectedTotalCost: number;
        foundTotalCost: number;
        expectedUnitCost?: number;
        foundUnitCost?: number;
    }[];

    expectedFinalBalance: {
        quantity: number;
        unitCost: number;
        totalCost: number;
    };
    actualFinalBalance: {
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
                    // esperado = lo que dice el ARCHIVO, encontrado = lo
                    // que la regla calculó con el CPP.
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
                        `Costo Total de Salidas Esperado (Archivo): ${metadata.totals.exit.totalCostArchivo}`,
                        `Costo Total de Salidas Encontrado (CPP): ${metadata.totals.exit.totalCost}`,
                        `Movimientos Analizados: ${metadata.movementCount}`,
                        // Solo la diferencia de ESTA fila -- no la lista
                        // completa del producto (evita mezclar con las otras
                        // filas, mismo criterio que RULE_003).
                        `Campos con diferencia: Costo Total de Salidas`
                    ].join("\n")
                });
            }


            for (const mismatch of metadata.unitCostMismatches) {
                const diferencia = Rule013Exporter.round(
                    mismatch.expectedTotalCost - mismatch.foundTotalCost
                );
                rows.push({
                    period,
                    productCode: result.product_code,
                    productDescription: result.product_name,
                    inconsistencyType: "Costo Unitario de Saldo Final",
                    expectedValue: mismatch.expectedTotalCost,
                    foundValue: mismatch.foundTotalCost,
                    difference: diferencia,
                    differencePercent: Rule013Exporter.percent(mismatch.expectedTotalCost, diferencia),
                    riskLevel: result.risk_level,
                    traceability: Rule013Exporter.unitCostTrace(codigoNormalizado, mismatch, metadata.movementCount)
                });
            }

            if (metadata.differences.includes("Costo Total de Saldo Final")) {
                const esperado = metadata.actualFinalBalance.totalCost;
                const encontrado = metadata.expectedFinalBalance.totalCost;
                const diferencia = Rule013Exporter.round(esperado - encontrado);
                rows.push({
                    period,
                    productCode: result.product_code,
                    productDescription: result.product_name,
                    inconsistencyType: "Costo Total de Saldo Final",
                    expectedValue: esperado,
                    foundValue: encontrado,
                    difference: diferencia,
                    differencePercent: Rule013Exporter.percent(esperado, diferencia),
                    riskLevel: result.risk_level,
                    traceability: [
                        codigoNormalizado,
                        `Saldo Inicial del Costo Total: ${metadata.initialBalance.totalCost}`,
                        `Costo Total por Entrada: ${metadata.totals.entry.totalCost}`,
                        `Costo Total de Salidas Encontrado: ${metadata.totals.exit.totalCost}`,
                        `Costo Total Esperado (Archivo): ${esperado}`,
                        `Costo Total Encontrado (CPP recalculado): ${encontrado}`,
                        `Rango Permitido (0%): ${encontrado} - ${encontrado}`,
                        `Movimientos Analizados: ${metadata.movementCount}`,
                        `Campos con diferencia: Costo Total de Saldo Final`
                    ].join("\n")
                });
            }
        }
        return rows;
    }

    private static unitCostTrace(
        codigoNormalizado: string,
        mismatch: Rule013Metadata["unitCostMismatches"][number],
        movementCount: number
    ): string {
        const lines = [
            codigoNormalizado,
            // `Documento: ${mismatch.document}`,
            // `Cantidad de esa Entrada: ${mismatch.quantity}`
        ];

        if (mismatch.expectedUnitCost !== undefined && mismatch.foundUnitCost !== undefined) {
            lines.push(
                `Costo Promedio Ponderado Unitario Esperado: ${mismatch.expectedUnitCost}`,
                `CPP encontrado: ${mismatch.foundUnitCost}`
            );
        }

        lines.push(
            `Movimientos Analizados: ${movementCount}`,
            `Campos con diferencia: Costo Unitario de Saldo Final`
        );

        return lines.join("\n");
    }

    private static round(value: number): number {
        return Math.round(value * 100) / 100;
    }
}
