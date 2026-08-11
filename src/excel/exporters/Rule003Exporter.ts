import ExcelJS from "exceljs";

import { BaseExcelExporter } from "./base/BaseExcelExporter";
import { ReportHeader } from "./base/ReportHeader";
import { AuditFindingRow } from "./base/AuditFindingRow";
import { DateUtils } from "../helpers/dateutils";

export interface Rule003Metadata {
    fromIndex: number;
    toIndex: number;
    finalBalance: {
        quantity: number;
        unitCost: number;
        totalCost: number;
    };
    initialBalance: {
        quantity: number;
        unitCost: number;
        totalCost: number;
    };
    differences: string[];
}
export class Rule003Exporter extends BaseExcelExporter {
    async export(
        results: any[],
        header: ReportHeader
    ) {
        const workbook = new ExcelJS.Workbook();
        workbook.creator = "HM Kardex Audit";
        workbook.created = new Date();
        const worksheet = workbook.addWorksheet("RULE_003");
        this.writeHeader(
            worksheet,
            "RULE_003 - Validación de Continuidad de Costos",
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

    private buildFindings(results: any[]): AuditFindingRow[] {
        const rows: AuditFindingRow[] = [];
        for (const result of results) {
            const metadata = result.metadata as Rule003Metadata;
            // COSTO UNITARIO
            rows.push({
                period:
                    `${DateUtils.monthName(metadata.fromIndex)} → ${DateUtils.monthName(metadata.toIndex)}`,
                productCode: result.product_code,
                productDescription: result.product_name,
                inconsistencyType: "Continuidad de Costo Unitario",
                expectedValue: metadata.finalBalance.unitCost,
                foundValue: metadata.initialBalance.unitCost,
                difference:
                    metadata.finalBalance.unitCost -
                    metadata.initialBalance.unitCost,
                differencePercent:
                    metadata.finalBalance.unitCost === 0
                        ? 0
                        : Math.abs(
                            (
                                metadata.finalBalance.unitCost -
                                metadata.initialBalance.unitCost
                            ) /
                            metadata.finalBalance.unitCost
                        ) * 100,
                riskLevel: result.risk_level,
                traceability: [
                    `Mes Cierre: ${DateUtils.monthName(metadata.fromIndex)}`,
                    `Mes Inicio: ${DateUtils.monthName(metadata.toIndex)}`,
                    `Costo Unitario Final: ${metadata.finalBalance.unitCost}`,
                    `Costo Unitario Inicial: ${metadata.initialBalance.unitCost}`,
                    `Campos con diferencia: ${metadata.differences.join(", ")}`
                ].join("\n")
            });

            // COSTO TOTAL
            rows.push({
                period:
                    `${DateUtils.monthName(metadata.fromIndex)} → ${DateUtils.monthName(metadata.toIndex)}`,
                productCode: result.product_code,
                productDescription: result.product_name,
                inconsistencyType: "Continuidad de Costo Total",
                expectedValue: metadata.finalBalance.totalCost,
                foundValue: metadata.initialBalance.totalCost,
                difference:
                    metadata.finalBalance.totalCost -
                    metadata.initialBalance.totalCost,
                differencePercent:
                    metadata.finalBalance.totalCost === 0
                        ? 0
                        : Math.abs(
                            (
                                metadata.finalBalance.totalCost -
                                metadata.initialBalance.totalCost
                            ) /
                            metadata.finalBalance.totalCost
                        ) * 100,
                riskLevel: result.risk_level,
                traceability: [
                    `Mes Cierre: ${DateUtils.monthName(metadata.fromIndex)}`,
                    `Mes Inicio: ${DateUtils.monthName(metadata.toIndex)}`,
                    `Costo Total Final: ${metadata.finalBalance.totalCost}`,
                    `Costo Total Inicial: ${metadata.initialBalance.totalCost}`,
                    `Campos con diferencia: ${metadata.differences.join(", ")}`
                ].join("\n")
            });
        }
        return rows;
    }
}