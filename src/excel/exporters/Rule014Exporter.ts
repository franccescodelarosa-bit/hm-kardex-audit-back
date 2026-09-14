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
            "RULE_014 - Validacion consolidada de sumatorias mensuales con la ecuacion de conciliacion global del kardex valorizado",
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

            if (metadata.differences.includes("Costo valorizado fuera del rango permitido")) {
                const diferencia = Rule014Exporter.round(
                    metadata.actualFinalBalance.totalCost - metadata.expectedFinalBalance.totalCost
                );
                rows.push({
                    period,
                    productCode: "CONSOLIDADO",
                    productDescription: "Consolidado Mensual",
                    inconsistencyType: "Costo valorizado fuera del rango permitido",
                    expectedValue: metadata.actualFinalBalance.totalCost,
                    foundValue: metadata.expectedFinalBalance.totalCost,
                    difference: diferencia,
                    differencePercent: Rule014Exporter.percent(
                        metadata.actualFinalBalance.totalCost,
                        diferencia
                    ),
                    riskLevel: result.risk_level,
                    traceability: [
                        "Mensaje del sistema (Anexo 03): ERROR DE CONSOLIDACIÓN DEL KARDEX Y LA DIFERENCIA",
                        `Punto de partida - Saldo Inicial consolidado del mes: ${metadata.initialBalance.totalCost}`,
                        `+ Total Entradas (Compras, Op. 02): ${metadata.totals.entry.totalCost}`,
                        `− Total Salidas (Ventas, Op. 01): ${metadata.totals.exit.totalCost}`,
                        `= Resultado de la fórmula (Encontrado): ${metadata.expectedFinalBalance.totalCost}`,
                        `Cierre real que trae el Kardex (Esperado): ${metadata.actualFinalBalance.totalCost}`,
                        `Diferencia: ${diferencia}`,
                        `Tolerancia: Sin tolerancia (debe coincidir exacto)`,
                        `Productos consolidados: ${metadata.productCount}`,
                        `Movimientos consolidados: ${metadata.movementCount}`
                    ].join("\n")
                });
            }
        }
        return rows;
    }
}
