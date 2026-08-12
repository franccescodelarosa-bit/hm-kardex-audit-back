import ExcelJS from "exceljs";

import { BaseExcelExporter } from "./base/BaseExcelExporter";
import { ReportHeader } from "./base/ReportHeader";
import { AuditFindingRow } from "./base/AuditFindingRow";
import { DateUtils } from "../helpers/dateutils";

export interface Rule004Metadata {
    transitItem: string;
    issueDate: string;
    warehouseDate: string;
    supplierRuc: string;
    supplier: string;
    document: string;
    normalizedDocument: string;
    month: string | null;
    duplicatedItems: number;
    products: string;
    expectedCost: number;
    foundCost: number;
}

export class Rule004Exporter extends BaseExcelExporter {
    async export(
        results: any[],
        header: ReportHeader
    ) {
        const workbook = new ExcelJS.Workbook();
        workbook.creator = "HM Kardex Audit";
        workbook.created = new Date();
        const worksheet = workbook.addWorksheet("RULE_004");
        this.writeHeader(
            worksheet,
            "RULE_004 - Validación de facturas de mercadería en transito al cierre del año",
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
            const metadata = result.metadata as Rule004Metadata;
            rows.push({
                period: metadata.month ? DateUtils.monthName(Number(metadata.month)) : "Sin período",
                productCode: metadata.transitItem,
                productDescription: metadata.products,
                inconsistencyType: "Mercadería en tránsito no registrada",
                expectedValue: metadata.expectedCost,
                foundValue: metadata.foundCost,
                difference: metadata.expectedCost - metadata.foundCost,
                differencePercent:
                    metadata.expectedCost === 0
                        ? 0
                        : Math.abs(
                            (
                                (metadata.expectedCost - metadata.foundCost) /
                                metadata.expectedCost
                            ) * 100
                        ),
                riskLevel: result.risk_level,
                traceability: [
                    `Documento: ${metadata.document}`,
                    `Documento Normalizado: ${metadata.normalizedDocument}`,
                    `Proveedor: ${metadata.supplier}`,
                    `RUC: ${metadata.supplierRuc}`,
                    `Fecha Emisión: ${metadata.issueDate}`,
                    `Fecha Almacén: ${metadata.warehouseDate}`,
                    `Items Encontrados: ${metadata.duplicatedItems}`,
                    `Item Tránsito: ${metadata.transitItem}`
                ].join("\n")
            });
        }
        return rows;
    }
}