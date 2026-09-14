import ExcelJS from "exceljs";

import { BaseExcelExporter } from "./base/BaseExcelExporter";
import { ReportHeader } from "./base/ReportHeader";
import { AuditFindingRow } from "./base/AuditFindingRow";
import { DateUtils } from "../helpers/dateutils";

export interface Rule007Metadata {
    month: number;
    movement: number;
    operation: string;
    document: string;
    previousBalance: {
        quantity: number;
        totalCost: number;
    };
    movimientos: {
        entryQuantity: number;
        exitQuantity: number;
        entryTotalCost: number;
        exitTotalCost: number;
    };
    expectedBalance: {
        quantity: number;
        totalCost: number;
    };
    actualBalance: {
        quantity: number;
        totalCost: number;
    };
    differences: string[];
}

export class Rule007Exporter extends BaseExcelExporter {
    async export(
        results: any[],
        header: ReportHeader
    ) {
        const workbook = new ExcelJS.Workbook();
        workbook.creator = "HM Kardex Audit";
        workbook.created = new Date();
        const worksheet = workbook.addWorksheet("RULE_007");
        this.writeHeader(
            worksheet,
            "RULE_007 - Validación de sumatorias en cantidades y costos",
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
            const metadata = result.metadata as Rule007Metadata;
            // VALIDACIÓN CANTIDAD
            rows.push({
                period: DateUtils.monthName(metadata.month),
                productCode: result.product_code,
                productDescription: result.product_name,
                inconsistencyType: "Fórmula Kardex - Cantidad",
                expectedValue: metadata.expectedBalance.quantity,
                foundValue: metadata.actualBalance.quantity,
                difference:
                    metadata.expectedBalance.quantity -
                    metadata.actualBalance.quantity,
                differencePercent:
                    metadata.expectedBalance.quantity === 0
                        ? 0
                        : Math.abs(
                            (
                                metadata.expectedBalance.quantity -
                                metadata.actualBalance.quantity
                            ) /
                            metadata.expectedBalance.quantity
                        ) * 100,
                riskLevel: result.risk_level,
                traceability: [
                    `Movimiento: ${metadata.movement}`,
                    `Operación: ${metadata.operation}`,
                    `Documento: ${metadata.document}`,
                    `Saldo Anterior: ${metadata.previousBalance.quantity}`,
                    `Entrada: ${metadata.movimientos.entryQuantity}`,
                    `Salida: ${metadata.movimientos.exitQuantity}`,
                    `Saldo Esperado: ${metadata.expectedBalance.quantity}`,
                    `Saldo Real: ${metadata.actualBalance.quantity}`
                ].join("\n")
            });

            // VALIDACIÓN COSTO TOTAL
            rows.push({
                period: DateUtils.monthName(metadata.month),
                productCode: result.product_code,
                productDescription: result.product_name,
                inconsistencyType: "Fórmula Kardex - Costo Total",
                expectedValue: metadata.expectedBalance.totalCost,
                foundValue: metadata.actualBalance.totalCost,
                difference:
                    metadata.expectedBalance.totalCost -
                    metadata.actualBalance.totalCost,
                differencePercent:
                    metadata.expectedBalance.totalCost === 0
                        ? 0
                        : Math.abs(
                            (
                                metadata.expectedBalance.totalCost -
                                metadata.actualBalance.totalCost
                            ) /
                            metadata.expectedBalance.totalCost
                        ) * 100,
                riskLevel: result.risk_level,
                traceability: [
                    `Movimiento: ${metadata.movement}`,
                    `Operación: ${metadata.operation}`,
                    `Documento: ${metadata.document}`,
                    `Saldo Anterior: ${metadata.previousBalance.totalCost}`,
                    `Entrada: ${metadata.movimientos.entryTotalCost}`,
                    `Salida: ${metadata.movimientos.exitTotalCost}`,
                    `Saldo Esperado: ${metadata.expectedBalance.totalCost}`,
                    `Saldo Real: ${metadata.actualBalance.totalCost}`
                ].join("\n")
            });
        }
        return rows;
    }
}