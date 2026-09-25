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

    private static equals(a: number, b: number): boolean {
        return Math.abs(a - b) < 0.01;
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
                    foundValue: "Producto no encontrado",
                    difference: "No aplicable",
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

            
            const fields: { name: string; expected: number; found: number }[] = [
                { name: "Costo Unitario", expected: finalBalance.unitCost, found: initialBalance.unitCost },
                { name: "Costo Total", expected: finalBalance.totalCost, found: initialBalance.totalCost }
            ];

            for (const field of fields) {
                const hasDifference = metadata.differences
                    ? metadata.differences.includes(field.name)
                    : !Rule003Exporter.equals(field.expected, field.found);

                if (!hasDifference) {
                    continue;
                }

                rows.push({
                    period: `${mesCierre} → ${mesInicio}`,
                    productCode: result.product_code,
                    productDescription: result.product_name,
                    inconsistencyType: `Continuidad de ${field.name}`,
                    expectedValue: field.expected,
                    foundValue: field.found,
                    difference: field.expected - field.found,
                    differencePercent:
                        field.expected === 0
                            ? 0
                            : Math.abs((field.expected - field.found) / field.expected) * 100,
                    riskLevel: result.risk_level,
                    traceability: [
                        `Mes Cierre: ${mesCierre}`,
                        `Mes Inicio: ${mesInicio}`,
                        `${field.name} Final: ${field.expected}`,
                        `${field.name} Inicial: ${field.found}`,
                        `Campos con diferencia: ${field.name}`
                    ].join("\n")
                });
            }
        }
        return rows;
    }
}
