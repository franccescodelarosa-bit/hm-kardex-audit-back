import ExcelJS from "exceljs";

import { BaseExcelExporter } from "./base/BaseExcelExporter";
import { ReportHeader } from "./base/ReportHeader";
import { AuditFindingRow } from "./base/AuditFindingRow";
import { DateUtils } from "../helpers/dateutils";

export interface Rule002Metadata {
    fromMonth: number;
    toMonth: number;
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

export class Rule002Exporter extends BaseExcelExporter {
    async export(
        results: any[],
        header: ReportHeader
    ) {
        const workbook = new ExcelJS.Workbook();
        workbook.creator = "HM Kardex Audit";
        workbook.created = new Date();
        const worksheet = workbook.addWorksheet("RULE_002");
        this.writeHeader(
            worksheet,
            "RULE_002 - Validación de Continuidad Mensual",
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
            const metadata = result.metadata as Rule002Metadata;
            // CONTINUIDAD CANTIDAD
            rows.push({
                period: `${DateUtils.monthName(metadata.fromMonth)} → ${DateUtils.monthName(metadata.toMonth)}`,
                productCode: result.product_code,
                productDescription: result.product_name,
                inconsistencyType: "Continuidad de Cantidad",
                expectedValue: metadata.finalBalance.quantity,
                foundValue: metadata.initialBalance.quantity,
                difference: metadata.finalBalance.quantity - metadata.initialBalance.quantity,
                differencePercent:
                    metadata.finalBalance.quantity === 0
                        ? 0
                        : Math.abs(
                            (
                                metadata.finalBalance.quantity -
                                metadata.initialBalance.quantity
                            ) /
                            metadata.finalBalance.quantity
                        ) * 100,
                riskLevel: result.risk_level,
                traceability: [
                    `Mes Final: ${DateUtils.monthName(metadata.fromMonth)}`,
                    `Mes Inicial: ${DateUtils.monthName(metadata.toMonth)}`,
                    `Campos con diferencia: ${metadata.differences.join(", ")}`
                ].join("\n")
            });

            // CONTINUIDAD COSTO UNITARIO
            rows.push({
                period: `${DateUtils.monthName(metadata.fromMonth)} → ${DateUtils.monthName(metadata.toMonth)}`,
                productCode: result.product_code,
                productDescription: result.product_name,
                inconsistencyType: "Continuidad de Costo Unitario",
                expectedValue: metadata.finalBalance.unitCost,
                foundValue: metadata.initialBalance.unitCost,
                difference: metadata.finalBalance.unitCost - metadata.initialBalance.unitCost,
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
                    `Mes Final: ${DateUtils.monthName(metadata.fromMonth)}`,
                    `Mes Inicial: ${DateUtils.monthName(metadata.toMonth)}`,
                    `Campos con diferencia: ${metadata.differences.join(", ")}`
                ].join("\n")
            });

            // CONTINUIDAD COSTO TOTAL
            rows.push({
                period: `${DateUtils.monthName(metadata.fromMonth)} → ${DateUtils.monthName(metadata.toMonth)}`,
                productCode: result.product_code,
                productDescription: result.product_name,
                inconsistencyType: "Continuidad de Costo Total",
                expectedValue: metadata.finalBalance.totalCost,
                foundValue: metadata.initialBalance.totalCost,
                difference: metadata.finalBalance.totalCost - metadata.initialBalance.totalCost,
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
                    `Mes Final: ${DateUtils.monthName(metadata.fromMonth)}`,
                    `Mes Inicial: ${DateUtils.monthName(metadata.toMonth)}`,
                    `Campos con diferencia: ${metadata.differences.join(", ")}`
                ].join("\n")
            });
        }
        return rows;
    }
}