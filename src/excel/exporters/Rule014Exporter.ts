import ExcelJS, { CellRichTextValue, RichText } from "exceljs";

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
                    traceability: Rule014Exporter.buildTraceability(metadata, diferencia)
                });
            }
        }
        return rows;
    }

    /*
     * Trazabilidad con el formato pedido por el cliente: verde = valor
     * esperado, azul = valor encontrado, rojo = operadores de la fórmula.
     */
    private static buildTraceability(
        metadata: Rule014Metadata,
        diferencia: number
    ): CellRichTextValue {
        const text = (value: string, color: string, underline = false): RichText => ({
            text: value,
            font: { bold: true, underline, color: { argb: color } }
        });
        const green = (value: string, underline = false) => text(value, GREEN, underline);
        const blue = (value: string, underline = false) => text(value, BLUE, underline);
        const lightBlue = (value: string) => text(value, LIGHT_BLUE);
        const red = (value: string) => text(value, RED);
        const black = (value: string) => text(value, BLACK);
        const money = Rule014Exporter.formatMoney;

        return {
            richText: [
                green("VALOR ESPERADO = SUMATORIA DE LOS SALDOS FINALES DE CADA ITEM DEL MES SEGUN EL KARDEX : "),
                green(money(metadata.actualFinalBalance.totalCost), true),
                blue("\nVALOR ENCONTRADO = SUMATORIA DE LOS SALDOS INICIALES DE CADA ITEM DEL MES"),
                red(" + "),
                blue("TOTAL ENTRADAS (Todos los ingresos)"),
                red(" - "),
                blue("TOTAL SALIDAS (todos los egresos)"),
                red("\n= "),
                blue("RESULTADO DE LA FORMULA (VALOR ENCONTRADO):\n"),
                blue(money(metadata.initialBalance.totalCost)),
                red(" + "),
                lightBlue(money(metadata.totals.entry.totalCost)),
                red(" - "),
                blue(money(metadata.totals.exit.totalCost)),
                red("\n= "),
                blue(money(metadata.expectedFinalBalance.totalCost), true),
                black("\n\nDiferencia = "),
                green(money(metadata.actualFinalBalance.totalCost)),
                red(" - "),
                lightBlue(money(metadata.expectedFinalBalance.totalCost)),
                lightBlue(" = "),
                black(money(diferencia)),
                black(
                    "\nTolerancia: Sin tolerancia (debe coincidir exacto)" +
                    `\nProductos consolidados: ${Rule014Exporter.formatNumber(metadata.productCount)}` +
                    `\nMovimientos consolidados: ${Rule014Exporter.formatNumber(metadata.movementCount)}`
                )
            ]
        };
    }

    private static formatMoney(value: number): string {
        return `S/ ${Rule014Exporter.formatNumber(value)}`;
    }

    private static formatNumber(value: number): string {
        const [integerPart, decimals] = Math.abs(value).toFixed(2).split(".");
        const groups = integerPart.replace(/\B(?=(\d{3})+(?!\d))/g, " ").split(" ");
        const thousands = groups.pop()!;
        const integer = groups.length > 0
            ? `${groups.join("´")},${thousands}`
            : thousands;
        return `${value < 0 ? "-" : ""}${integer}.${decimals}`;
    }
}

const GREEN = "FF00B050";
const BLUE = "FF0070C0";
const LIGHT_BLUE = "FF00B0F0";
const RED = "FFFF0000";
const BLACK = "FF000000";
