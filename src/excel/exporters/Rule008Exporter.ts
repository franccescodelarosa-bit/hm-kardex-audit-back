import ExcelJS from "exceljs";

import { BaseExcelExporter } from "./base/BaseExcelExporter";
import { ReportHeader } from "./base/ReportHeader";
import { AuditFindingRow } from "./base/AuditFindingRow";
import { DateUtils } from "../helpers/dateutils";

export interface Rule008Metadata {
    source: string;
    month: number;
}
export class Rule008Exporter extends BaseExcelExporter {
    async export(
        results: any[],
        header: ReportHeader
    ) {
        const workbook = new ExcelJS.Workbook();
        workbook.creator = "HM Kardex Audit";
        workbook.created = new Date();
        const worksheet = workbook.addWorksheet("RULE_008");
        this.writeHeader(
            worksheet,
            "RULE_008 - Productos No Encontrados en el Maestro",
            header,
            "J"
        );
        this.writeTableHeader(worksheet);
        const findings = this.buildFindings(results);
        this.writeRows(
            worksheet,
            findings
        );
        worksheet.views = [{
            state: "frozen",
            ySplit: 4
        }];
        worksheet.autoFilter = {
            from: "A4",
            to: "J4"
        };
        return workbook;
    }

    private buildFindings(results: any[]): AuditFindingRow[] {
        const rows: AuditFindingRow[] = [];
        for (const result of results) {
            const metadata = result.metadata as Rule008Metadata;
            rows.push({
                period: DateUtils.monthName(metadata.month),
                productCode: result.product_code,
                productDescription: result.product_name,
                inconsistencyType: "Producto no encontrado en el maestro",
                expectedValue: "Producto registrado",
                foundValue: "Producto inexistente",
                difference: result.product_code,
                differencePercent: undefined,
                riskLevel: result.risk_level,
                traceability: [
                    `Origen: ${metadata.source}`,
                    `Código: ${result.product_code}`,
                    `Producto: ${result.product_name}`
                ].join("\n")
            });
        }
        return rows;
    }
}