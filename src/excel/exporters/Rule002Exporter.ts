import ExcelJS from "exceljs";

import { BaseExcelExporter } from "./base/BaseExcelExporter";
import { ReportHeader } from "./base/ReportHeader";
import { AuditFindingRow } from "./base/AuditFindingRow";
import { DateUtils } from "../helpers/dateutils";

export interface Rule002Metadata {
    fromMonth: number;
    toMonth: number;
    finalQuantity: number;
    initialQuantity: number;
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
            "RULE_002 - Validación de continuidad mensual de los saldos final e inicial en cantidades",
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

        rows.push({
            period: `${DateUtils.monthName(metadata.fromMonth)} → ${DateUtils.monthName(metadata.toMonth)}`,
            productCode: result.product_code,
            productDescription: result.product_name,
            inconsistencyType: "Continuidad de Cantidad",

            expectedValue: metadata.finalQuantity,
            foundValue: metadata.initialQuantity,

            difference:
                metadata.finalQuantity -
                metadata.initialQuantity,

            differencePercent:
                metadata.finalQuantity === 0
                    ? 0
                    : Math.abs(
                        (
                            metadata.finalQuantity -
                            metadata.initialQuantity
                        ) / metadata.finalQuantity
                    ) * 100,

            riskLevel: result.risk_level,

            traceability: [
                `Mes Final: ${DateUtils.monthName(metadata.fromMonth)}`,
                `Mes Inicial: ${DateUtils.monthName(metadata.toMonth)}`,
                `Campos con diferencia: Cantidad`
            ].join("\n")
        });
    }

    return rows;
}
}