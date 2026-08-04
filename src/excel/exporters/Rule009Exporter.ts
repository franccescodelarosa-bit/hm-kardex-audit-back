import ExcelJS from "exceljs";

import { BaseExcelExporter } from "./base/BaseExcelExporter";
import { ReportHeader } from "./base/ReportHeader";
import { AuditFindingRow } from "./base/AuditFindingRow";
import { DateUtils } from "../helpers/dateutils";

export interface Rule009Metadata {
    date: string;
    month: number;
    document: string;
    operation: string;
    entryQuantity: number;
    entryUnitCost: number;
    entryTotalCost: number;
    balanceQuantity: number;
    balanceUnitCost: number;
    balanceTotalCost: number;
}
export class Rule009Exporter extends BaseExcelExporter {
    async export(
        results: any[],
        header: ReportHeader
    ) {
        const workbook = new ExcelJS.Workbook();
        workbook.creator = "HM Kardex Audit";
        workbook.created = new Date();
        const worksheet = workbook.addWorksheet("RULE_009");
        this.writeHeader(
            worksheet,
            "RULE_009 - Ingresos por Ajuste de Inventario",
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
            const metadata = result.metadata as Rule009Metadata;
            rows.push({
                period: DateUtils.monthName(metadata.month),
                productCode: result.product_code,
                productDescription: result.product_name,
                inconsistencyType: "Ingreso por Ajuste de Inventario",
                expectedValue: "Ingreso sustentado",
                foundValue: metadata.operation,
                difference: metadata.entryQuantity,
                differencePercent: undefined,
                riskLevel: result.risk_level,
                traceability: [
                    `Fecha: ${metadata.date}`,
                    `Documento: ${metadata.document}`,
                    `Operación: ${metadata.operation}`,
                    `Cantidad Ingreso: ${metadata.entryQuantity}`,
                    `Costo Unitario: ${metadata.entryUnitCost}`,
                    `Costo Total: ${metadata.entryTotalCost}`,
                    `Saldo Cantidad: ${metadata.balanceQuantity}`,
                    `Saldo Costo Unitario: ${metadata.balanceUnitCost}`,
                    `Saldo Costo Total: ${metadata.balanceTotalCost}`
                ].join("\n")
            });
        }
        return rows;
    }
}