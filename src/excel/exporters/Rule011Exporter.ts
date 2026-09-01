import ExcelJS from "exceljs";

import { BaseExcelExporter } from "./base/BaseExcelExporter";
import { ReportHeader } from "./base/ReportHeader";
import { AuditFindingRow } from "./base/AuditFindingRow";
import { DateUtils } from "../helpers/dateutils";

export interface Rule011Metadata {
    date: string;
    month: number;
    document: string;
    operation: string;
    previousCost: number;
    currentCost: number;
    variationPercent: number;
}

export class Rule011Exporter extends BaseExcelExporter {
    async export(
        results: any[],
        header: ReportHeader
    ) {
        const workbook = new ExcelJS.Workbook();
        workbook.creator = "HM Kardex Audit";
        workbook.created = new Date();
        const worksheet = workbook.addWorksheet("RULE_011");
        this.writeHeader(
            worksheet,
            "RULE_011 - Deteccion de Variación inusual de costos unitarios",
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
            const metadata = result.metadata as Rule011Metadata;
            rows.push({
                period: DateUtils.monthName(metadata.month),
                productCode: result.product_code,
                productDescription: result.product_name,
                inconsistencyType: "Variación Inusual del Costo Unitario",
                expectedValue: metadata.previousCost,
                foundValue: metadata.currentCost,
                difference:
                    metadata.currentCost - metadata.previousCost,
                differencePercent:
                    metadata.variationPercent,
                riskLevel: result.risk_level,
                traceability: [
                    `Fecha: ${metadata.date}`,
                    `Documento: ${metadata.document}`,
                    `Operación: ${metadata.operation}`,
                    `Costo Anterior: ${metadata.previousCost}`,
                    `Costo Actual: ${metadata.currentCost}`,
                    `Variación: ${metadata.variationPercent.toFixed(2)} %`
                ].join("\n")
            });
        }
        return rows;
    }
}