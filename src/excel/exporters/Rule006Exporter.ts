import ExcelJS from "exceljs";

import { BaseExcelExporter } from "./base/BaseExcelExporter";
import { ReportHeader } from "./base/ReportHeader";
import { AuditFindingRow } from "./base/AuditFindingRow";

export interface Rule006Metadata {
    source: string;
    occurrences: number;
    rows: number[];
}

export class Rule006Exporter extends BaseExcelExporter {
    async export(
        results: any[],
        header: ReportHeader
    ) {
        const workbook = new ExcelJS.Workbook();
        workbook.creator = "HM Kardex Audit";
        workbook.created = new Date();
        const worksheet = workbook.addWorksheet("RULE_006");
        this.writeHeader(
            worksheet,
            "RULE_006 - Detección de Productos Duplicados",
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
            const metadata = result.metadata as Rule006Metadata;
            rows.push({
                period: "-",
                productCode: result.product_code,
                productDescription: result.product_name,
                inconsistencyType: "Producto Duplicado",
                expectedValue: "1 registro",
                foundValue: `${metadata.occurrences} registros`,
                difference: metadata.occurrences - 1,
                differencePercent: undefined,
                riskLevel: result.risk_level,
                traceability: [
                    `Origen: ${metadata.source}`,
                    `Ocurrencias: ${metadata.occurrences}`,
                    `Filas: ${metadata.rows.join(", ")}`
                ].join("\n")
            });
        }
        return rows;
    }
}