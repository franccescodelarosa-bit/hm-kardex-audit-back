import ExcelJS from "exceljs";

import { BaseExcelExporter } from "./base/BaseExcelExporter";
import { ReportHeader } from "./base/ReportHeader";
import { AuditFindingRow } from "./base/AuditFindingRow";
import { DateUtils } from "../helpers/dateutils";

export interface Rule012Metadata {
    issueDate: string;
    warehouseDate: string;
    supplierRuc: string;
    supplier: string;
    document: string;
    normalizedDocument: string;
    transitAmount: number;
    kardexAmount: number;
    difference: number;
    movements: number;
    month?: number;
    transitItem: string;
}

export class Rule012Exporter extends BaseExcelExporter {
    async export(
        results: any[],
        header: ReportHeader
    ) {
        const workbook = new ExcelJS.Workbook();
        workbook.creator = "HM Kardex Audit";
        workbook.created = new Date();
        const worksheet = workbook.addWorksheet("RULE_012");
        this.writeHeader(
            worksheet,
            "RULE_012 - Validación del Monto de Mercadería en Tránsito vs Kardex",
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
            const metadata = result.metadata as Rule012Metadata;
            rows.push({
                period: DateUtils.monthName(Number(metadata.month)),
                productCode: metadata.transitItem,
                productDescription: result.product_name,
                inconsistencyType: "Monto de Mercadería en Tránsito",
                expectedValue: metadata.transitAmount,
                foundValue: metadata.kardexAmount,
                difference: metadata.difference,
                differencePercent:
                    metadata.transitAmount === 0
                        ? 0
                        : Math.abs(metadata.difference / metadata.transitAmount) * 100,

                riskLevel: result.risk_level,
                traceability: [
                    `Documento: ${metadata.document}`,
                    `Documento Normalizado: ${metadata.normalizedDocument}`,
                    `Proveedor: ${metadata.supplier}`,
                    `RUC: ${metadata.supplierRuc}`,
                    `Fecha Emisión: ${metadata.issueDate}`,
                    `Fecha Ingreso Almacén: ${metadata.warehouseDate}`,
                    `Movimientos Kardex: ${metadata.movements}`,
                    `Ítem Tránsito: ${metadata.transitItem}`
                ].join("\n")
            });
        }
        return rows;
    }
}