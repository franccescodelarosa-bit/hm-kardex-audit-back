import ExcelJS from "exceljs";

import { BaseExcelExporter } from "./base/BaseExcelExporter";
import { ReportHeader } from "./base/ReportHeader";
import { AuditFindingRow } from "./base/AuditFindingRow";
import { DateUtils } from "../helpers/dateutils";

export interface Rule001Metadata {
    inventoryCode: string;
    normalizedCode: string;
    inventoryStock: number;
    kardexStock: number;
    inventoryUnitCost: number;
    kardexUnitCost: number;
    inventoryTotalCost: number;
    kardexTotalCost: number;
    kardexMovements: number;
}

export class Rule001Exporter extends BaseExcelExporter {
    async export(
        results: any[],
        header: ReportHeader
    ) {
        const workbook = new ExcelJS.Workbook();
        workbook.creator = "HM Kardex Audit";
        workbook.created = new Date();
        const worksheet = workbook.addWorksheet("RULE_001");
        this.writeHeader(
            worksheet,
            "RULE_001 - Validación Inventario Final vs Kardex",
            header,
            "J"
        );
        this.writeTableHeader(worksheet);
        const findings = this.expand(results);
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

    private expand(results: any[]): AuditFindingRow[] {
        const rows: AuditFindingRow[] = [];
        for (const result of results) {
            const metadata = result.metadata as Rule001Metadata;
            // STOCK
            rows.push({
                period: DateUtils.monthName(result.month),
                productCode: result.product_code,
                productDescription: result.product_name,
                inconsistencyType: "Stock Final",
                expectedValue: metadata.inventoryStock,
                foundValue: metadata.kardexStock,
                difference: metadata.inventoryStock - metadata.kardexStock,
                differencePercent : metadata.inventoryStock === 0 ? 0 : Math.abs( (metadata.inventoryStock - metadata.kardexStock) / metadata.inventoryStock ) * 100,
                riskLevel: result.risk_level,
                traceability: [
                    `Código Inventario: ${metadata.inventoryCode}`,
                    `Código Normalizado: ${metadata.normalizedCode}`,
                    `Movimientos Kardex: ${metadata.kardexMovements}`
                ].join("\n")
            });

            // COSTO UNITARIO
            rows.push({
                period: DateUtils.monthName(result.month),
                productCode: result.product_code,
                productDescription: result.product_name,
                inconsistencyType: "Costo Unitario",
                expectedValue: metadata.inventoryUnitCost,
                foundValue: metadata.kardexUnitCost,
                difference: metadata.inventoryUnitCost - metadata.kardexUnitCost,
                differencePercent : metadata.inventoryUnitCost === 0 ? 0 : Math.abs( (metadata.inventoryUnitCost - metadata.kardexUnitCost) / metadata.inventoryUnitCost ) * 100,
                riskLevel: result.risk_level,
                traceability: [
                    `Costo Inventario: ${metadata.inventoryUnitCost}`,
                    `Costo Kardex: ${metadata.kardexUnitCost}`,
                    `Movimientos Kardex: ${metadata.kardexMovements}`
                ].join("\n")                
            });

            // COSTO TOTAL
            rows.push({
                period: DateUtils.monthName(result.month),
                productCode: result.product_code,
                productDescription: result.product_name,
                inconsistencyType: "Costo Total",
                expectedValue: metadata.inventoryTotalCost,
                foundValue: metadata.kardexTotalCost,
                difference: metadata.inventoryTotalCost - metadata.kardexTotalCost,
                differencePercent : metadata.inventoryTotalCost === 0 ? 0 : Math.abs( (metadata.inventoryTotalCost - metadata.kardexTotalCost) / metadata.inventoryTotalCost ) * 100,
                riskLevel: result.risk_level,                
                traceability: [
                    `Total Inventario: ${metadata.inventoryTotalCost}`,
                    `Total Kardex: ${metadata.kardexTotalCost}`,
                    `Movimientos Kardex: ${metadata.kardexMovements}`
                ].join("\n")    
            });
        }
        return rows;
    }
}