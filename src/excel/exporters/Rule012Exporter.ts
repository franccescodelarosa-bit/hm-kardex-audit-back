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

    expectedCost: number;
    kardexCost: number;
    difference: number;
    differencePercent: number;

    thresholdPercent: number;
    isIncident: boolean;

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
                period: metadata.month
                    ? DateUtils.monthName(Number(metadata.month))
                    : "Sin período",
                productCode: result.product_code,

                productDescription: result.product_name,

                inconsistencyType:
                    metadata.isIncident
                        ? "INCIDENCIA"
                        : "ACEPTADA",

                expectedValue: metadata.expectedCost,

                foundValue: metadata.kardexCost,

                difference: metadata.difference,

                differencePercent: metadata.differencePercent,

                riskLevel: result.risk_level,

                traceability: [
                    `Documento: ${metadata.document}`,
                    `Documento Normalizado: ${metadata.normalizedDocument}`,
                    `Proveedor: ${metadata.supplier}`,
                    `RUC: ${metadata.supplierRuc}`,
                    `Fecha Emisión: ${metadata.issueDate}`,
                    `Fecha Ingreso Almacén: ${metadata.warehouseDate}`,
                    `Valor Documento: ${metadata.expectedCost.toFixed(2)}`,
                    `Valor Kardex: ${metadata.kardexCost.toFixed(2)}`,
                    `Diferencia: ${metadata.difference.toFixed(2)}`,
                    `% Diferencia: ${metadata.differencePercent.toFixed(2)}%`,
                    `Umbral permitido: ${metadata.thresholdPercent.toFixed(2)}%`,
                    `Resultado: ${metadata.isIncident ? "INCIDENCIA" : "ACEPTADA"}`,
                    `Movimientos Kardex: ${metadata.movements}`,
                    `Ítem Tránsito: ${metadata.transitItem}`
                ].join("\n")
            });
        }
        return rows;
    }
}