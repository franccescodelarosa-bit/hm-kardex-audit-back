import ExcelJS from "exceljs";

import { BaseExcelExporter } from "./base/BaseExcelExporter";
import { ReportHeader } from "./base/ReportHeader";
import { AuditFindingRow } from "./base/AuditFindingRow";
import { DateUtils } from "../helpers/dateutils";

export interface Rule003Metadata {
    fromIndex: number;
    toIndex: number;
    finalBalance: {
        quantity: number;
        unitCost: number;
        totalCost: number;
    } | null;
    initialBalance: {
        quantity: number;
        unitCost: number;
        totalCost: number;
    } | null;
    differences?: string[];
}
export class Rule003Exporter extends BaseExcelExporter {
    async export(
        results: any[],
        header: ReportHeader
    ) {
        const workbook = new ExcelJS.Workbook();
        workbook.creator = "HM Kardex Audit";
        workbook.created = new Date();
        const worksheet = workbook.addWorksheet("RULE_003");
        this.writeHeader(
            worksheet,
            "RULE_003 - Validación de continuidad mensual de los saldos final e inicial en costos",
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
            const metadata = result.metadata as Rule003Metadata;
            const mesCierre = DateUtils.monthName(metadata.fromIndex);
            const mesInicio = DateUtils.monthName(metadata.toIndex);

            if (!metadata.initialBalance) {
                rows.push({
                    period: `${mesCierre} → ${mesInicio}`,
                    productCode: result.product_code,
                    productDescription: result.product_name,
                    inconsistencyType: "Producto no encontrado en el mes siguiente",
                    expectedValue: metadata.finalBalance?.totalCost ?? 0,
                    foundValue: 0,
                    difference: metadata.finalBalance?.totalCost ?? 0,
                    differencePercent: 0,
                    riskLevel: result.risk_level,
                    traceability: [
                        `Mes de Cierre: ${mesCierre}`,
                        `Costo Total Final (${mesCierre}): ${metadata.finalBalance?.totalCost ?? 0}`,
                        `Mes Siguiente: ${mesInicio}`,
                        `El producto no tiene Kardex registrado en ${mesInicio} — no se puede validar la continuidad.`
                    ].join("\n")
                });
                continue;
            }

            const finalBalance = metadata.finalBalance!;
            const initialBalance = metadata.initialBalance;

            rows.push({
                period: `${mesCierre} → ${mesInicio}`,
                productCode: result.product_code,
                productDescription: result.product_name,
                inconsistencyType: "Continuidad de Costo Unitario",
                expectedValue: finalBalance.unitCost,
                foundValue: initialBalance.unitCost,
                difference:
                    finalBalance.unitCost -
                    initialBalance.unitCost,
                differencePercent:
                    finalBalance.unitCost === 0
                        ? 0
                        : Math.abs(
                            (
                                finalBalance.unitCost -
                                initialBalance.unitCost
                            ) /
                            finalBalance.unitCost
                        ) * 100,
                riskLevel: result.risk_level,
                traceability: [
                    `Mes Cierre: ${mesCierre}`,
                    `Mes Inicio: ${mesInicio}`,
                    `Costo Unitario Final: ${finalBalance.unitCost}`,
                    `Costo Unitario Inicial: ${initialBalance.unitCost}`,
                    `Campos con diferencia: Costo Unitario`
                ].join("\n")
            });

            // COSTO TOTAL -- también siempre visible, misma decisión que arriba.
            rows.push({
                period: `${mesCierre} → ${mesInicio}`,
                productCode: result.product_code,
                productDescription: result.product_name,
                inconsistencyType: "Continuidad de Costo Total",
                expectedValue: finalBalance.totalCost,
                foundValue: initialBalance.totalCost,
                difference:
                    finalBalance.totalCost -
                    initialBalance.totalCost,
                differencePercent:
                    finalBalance.totalCost === 0
                        ? 0
                        : Math.abs(
                            (
                                finalBalance.totalCost -
                                initialBalance.totalCost
                            ) /
                            finalBalance.totalCost
                        ) * 100,
                riskLevel: result.risk_level,
                traceability: [
                    `Mes Cierre: ${mesCierre}`,
                    `Mes Inicio: ${mesInicio}`,
                    `Costo Total Final: ${finalBalance.totalCost}`,
                    `Costo Total Inicial: ${initialBalance.totalCost}`,
                    `Campos con diferencia: Costo Total`
                ].join("\n")
            });
        }
        return rows;
    }
}