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
    // campos del cierre distintos de 0 cuando el producto no pasa al mes siguiente
    missingFields?: string[];
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

            if (result.error_type === "INITIAL_BALANCE_NOT_FOUND_NEXT_MONTH") {
                rows.push(...this.missingRows(
                    result,
                    metadata,
                    "Sin Saldo Inicial en el mes siguiente",
                    "Sin Saldo Inicial (TipoOp 16)",
                    `El producto existe en ${mesInicio}, pero no tiene Saldo Inicial (TipoOp 16) en ${mesInicio} — no se puede validar la continuidad.`
                ));
                continue;
            }

            if (!metadata.initialBalance) {
                rows.push(...this.missingRows(
                    result,
                    metadata,
                    "Producto no encontrado en el mes siguiente",
                    "Producto no encontrado",
                    `El producto no tiene Kardex registrado en ${mesInicio} — no se puede validar la continuidad.`
                ));
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
    /*
     * Producto que no pasa al mes siguiente (no aparece o no tiene op 16):
     * una fila por cada campo del cierre distinto de 0 (decisión de la
     * usuaria). Registros viejos sin `missingFields`: una sola fila con el
     * Costo Total, como antes.
     */
    private missingRows(
        result: any,
        metadata: Rule003Metadata,
        label: string,
        foundValue: string,
        explanation: string
    ): AuditFindingRow[] {
        const mesCierre = DateUtils.monthName(metadata.fromIndex);
        const mesInicio = DateUtils.monthName(metadata.toIndex);
        const valuesByField: Record<string, number> = {
            "Cantidad": metadata.finalBalance?.quantity ?? 0,
            "Costo Unitario": metadata.finalBalance?.unitCost ?? 0,
            "Costo Total": metadata.finalBalance?.totalCost ?? 0
        };
        const fields = metadata.missingFields?.length
            ? metadata.missingFields
            : [null];

        return fields.map(field => ({
            period: `${mesCierre} → ${mesInicio}`,
            productCode: result.product_code,
            productDescription: result.product_name,
            inconsistencyType: field ? `${field} - ${label}` : label,
            expectedValue: field ? valuesByField[field] : valuesByField["Costo Total"],
            foundValue,
            difference: "No aplicable",
            riskLevel: result.risk_level,
            traceability: [
                `Mes de Cierre: ${mesCierre}`,
                `Cantidad Final (${mesCierre}): ${valuesByField["Cantidad"]}`,
                `Costo Unitario Final (${mesCierre}): ${valuesByField["Costo Unitario"]}`,
                `Costo Total Final (${mesCierre}): ${valuesByField["Costo Total"]}`,
                `Mes Siguiente: ${mesInicio}`,
                explanation
            ].join("\n")
        }));
    }
}
