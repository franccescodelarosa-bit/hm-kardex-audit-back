import ExcelJS from "exceljs";

import { BaseExcelExporter } from "./base/BaseExcelExporter";
import { ReportHeader } from "./base/ReportHeader";
import { AuditFindingRow } from "./base/AuditFindingRow";
import { DateUtils } from "../helpers/dateutils";
export interface Rule010Metadata {
    date: string;
    month: number;
    document: string;
    operation: string;
    exitQuantity: number;
    exitUnitCost: number;
    exitTotalCost: number;
    balanceQuantity: number;
    balanceUnitCost: number;
    balanceTotalCost: number;
}

export class Rule010Exporter extends BaseExcelExporter {
    async export(
        results: any[],
        header: ReportHeader
    ) {
        const workbook = new ExcelJS.Workbook();
        workbook.creator = "HM Kardex Audit";
        workbook.created = new Date();
        const worksheet = workbook.addWorksheet("RULE_010");
        this.writeHeader(
            worksheet,
            "RULE_010 - Salidas por Ajuste de Inventario",
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
            const metadata = result.metadata as Rule010Metadata;
            rows.push({
                period: DateUtils.monthName(metadata.month),
                productCode: result.product_code,
                productDescription: result.product_name,
                inconsistencyType: "Salida por Ajuste de Inventario",
                expectedValue: 0,
                foundValue: metadata.exitQuantity,
                difference: metadata.exitQuantity,
                differencePercent: undefined,
                riskLevel: result.risk_level,
                traceability: [
                    `Fecha: ${metadata.date}`,
                    `Documento: ${metadata.document}`,
                    `Operación: ${metadata.operation}`,
                    `Cantidad Salida: ${metadata.exitQuantity}`,
                    `Costo Unitario: ${metadata.exitUnitCost}`,
                    `Costo Total: ${metadata.exitTotalCost}`,
                    `Saldo Cantidad: ${metadata.balanceQuantity}`,
                    `Saldo Costo Unitario: ${metadata.balanceUnitCost}`,
                    `Saldo Costo Total: ${metadata.balanceTotalCost}`
                ].join("\n")
            });
        }
        return rows;
    }
}