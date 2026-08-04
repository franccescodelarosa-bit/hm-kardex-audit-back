import ExcelJS from "exceljs";

import { BaseExcelExporter } from "./base/BaseExcelExporter";
import { ReportHeader } from "./base/ReportHeader";
import { AuditFindingRow } from "./base/AuditFindingRow";
import { DateUtils } from "../helpers/dateutils";

export interface Rule005Metadata {
    source: string;
    stock?: number;
    unitCost?: number;
    totalCost?: number;
    date?: string;
    month?: number;
    document?: string;
    operation?: string;
    balanceQuantity?: number;
    balanceUnitCost?: number;
    balanceTotalCost?: number;
    entryQuantity?: number;
    entryUnitCost?: number;
    entryTotalCost?: number;
    exitQuantity?: number;
    exitUnitCost?: number;
    exitTotalCost?: number;
    negatives?: string[];
}

export class Rule005Exporter extends BaseExcelExporter {
    async export(
        results: any[],
        header: ReportHeader
    ) {
        const workbook = new ExcelJS.Workbook();
        workbook.creator = "HM Kardex Audit";
        workbook.created = new Date();
        const worksheet = workbook.addWorksheet("RULE_005");
        this.writeHeader(
            worksheet,
            "RULE_005 - Validación de Saldos Negativos",
            header,
            "J"
        );
        this.writeTableHeader(worksheet);
        const findings = this.buildFindings(results);
        this.writeRows(worksheet, findings);
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
            const metadata = result.metadata as Rule005Metadata;
            for (const negative of metadata.negatives ?? []) {
                rows.push({
                    period: DateUtils.monthName(Number(metadata.month)),
                    productCode: result.product_code,
                    productDescription: result.product_name,
                    inconsistencyType: `Saldo Negativo (${negative})`,
                    expectedValue: ">= 0",
                    foundValue: this.getNegativeValue(metadata, negative),
                    difference: this.getNegativeValue(metadata, negative),
                    differencePercent: undefined,
                    riskLevel: result.risk_level,
                    traceability: [
                        `Origen: ${metadata.source}`,
                        `Documento: ${metadata.document}`,
                        `Operación: ${metadata.operation}`,
                        `Fecha: ${metadata.date}`,
                        `Campos Negativos: ${metadata.negatives?.join(", ")}`
                    ].join("\n")
                });
            }
        }
        return rows;
    }

    private getNegativeValue(
        metadata: Rule005Metadata,
        field: string
    ): number | string {
        switch (field) {
            case "stock":
                return metadata.stock ?? "";
            case "unitCost":
                return metadata.unitCost ?? "";
            case "totalCost":
                return metadata.totalCost ?? "";
            case "balanceQuantity":
                return metadata.balanceQuantity ?? "";
            case "balanceUnitCost":
                return metadata.balanceUnitCost ?? "";
            case "balanceTotalCost":
                return metadata.balanceTotalCost ?? "";
            case "entryQuantity":
                return metadata.entryQuantity ?? "";
            case "entryUnitCost":
                return metadata.entryUnitCost ?? "";
            case "entryTotalCost":
                return metadata.entryTotalCost ?? "";
            case "exitQuantity":
                return metadata.exitQuantity ?? "";
            case "exitUnitCost":
                return metadata.exitUnitCost ?? "";
            case "exitTotalCost":
                return metadata.exitTotalCost ?? "";
            default:
                return "";
        }
    }
}