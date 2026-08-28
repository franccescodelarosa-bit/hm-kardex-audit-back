import ExcelJS from "exceljs";

import { BaseExcelExporter } from "./base/BaseExcelExporter";
import { ReportHeader } from "./base/ReportHeader";
import { AuditFindingRow } from "./base/AuditFindingRow";
import { DateUtils } from "../helpers/dateutils";

export interface Rule006DuplicateOccurrence {
    date: string | null;
    document: string;
    movementCount: number;
}

export interface Rule006Metadata {
    source: string;
    month?: number;
    occurrences: number;
    rows?: number[];
    duplicateOccurrences?: Rule006DuplicateOccurrence[];
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
            "RULE_006 - Detección de duplicidad de códigos",
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
                period: metadata.month
                    ? DateUtils.monthName(Number(metadata.month))
                    : "-",
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
                    ...this.buildFilasLines(metadata)
                ].join("\n")
            });
        }
        return rows;
    }

    private buildFilasLines(metadata: Rule006Metadata): string[] {
        if (metadata.duplicateOccurrences?.length) {
            return metadata.duplicateOccurrences.map((occ, i) =>
                `Ocurrencia ${i + 1}: ${occ.date ?? "Sin fecha"}, Doc. ${occ.document || "Sin documento"}, ${occ.movementCount} movimiento(s)`
            );
        }
        return [`Filas: ${(metadata.rows ?? []).join(", ")}`];
    }
}