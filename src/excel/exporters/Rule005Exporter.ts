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
            "RULE_005 - Detección de saldos negativos en cantidad, costo unitario y costo total",
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
                    inconsistencyType: negative,
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
                        `Campos Negativos: ${negative}`
                    ].join("\n")
                });
            }
        }
        return rows;
    }

    private static readonly FIELD_BY_LABEL: Record<string, keyof Rule005Metadata> = {
        "Cantidad de Entrada": "entryQuantity",
        "Costo Unitario de Entrada": "entryUnitCost",
        "Costo Total de Entrada": "entryTotalCost",
        "Cantidad de Salida": "exitQuantity",
        "Costo Unitario de Salida": "exitUnitCost",
        "Costo Total de Salida": "exitTotalCost",
        "Cantidad de Saldo": "balanceQuantity",
        "Costo Unitario de Saldo": "balanceUnitCost",
        "Costo Total de Saldo": "balanceTotalCost",

        // Etiquetas viejas (compatibilidad hacia atrás)
        "Cantidad Entrada": "entryQuantity",
        "Costo Entrada": "entryUnitCost",
        "Total Entrada": "entryTotalCost",
        "Cantidad Salida": "exitQuantity",
        "Costo Salida": "exitUnitCost",
        "Total Salida": "exitTotalCost",
        "Cantidad Saldo": "balanceQuantity",
        "Costo Saldo": "balanceUnitCost",
        "Total Saldo": "balanceTotalCost"
    };

    private getNegativeValue(
        metadata: Rule005Metadata,
        field: string
    ): number | string {
        const key = Rule005Exporter.FIELD_BY_LABEL[field];
        if (!key) {
            return "";
        }
        return (metadata[key] as number | undefined) ?? "";
    }
}