import ExcelJS from "exceljs";

import { BaseExcelExporter } from "./base/BaseExcelExporter";
import { ReportHeader } from "./base/ReportHeader";
import { AuditFindingRow } from "./base/AuditFindingRow";
import { DateUtils } from "../helpers/dateutils";

export interface Rule002Metadata {
    fromMonth: number;
    toMonth: number;
    finalQuantity: number;
    initialQuantity: number | null;
}

export class Rule002Exporter extends BaseExcelExporter {
    async export(
        results: any[],
        header: ReportHeader
    ) {
        const workbook = new ExcelJS.Workbook();
        workbook.creator = "HM Kardex Audit";
        workbook.created = new Date();
        const worksheet = workbook.addWorksheet("RULE_002");
        this.writeHeader(
            worksheet,
            "RULE_002 - Validación de continuidad mensual de los saldos final e inicial en cantidades",
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
            const metadata = result.metadata as Rule002Metadata;

            const mesCierre = DateUtils.monthName(metadata.fromMonth);
            const mesSiguiente = DateUtils.monthName(metadata.toMonth);

            const sinKardexSiguiente =
                metadata.initialQuantity === null ||
                metadata.initialQuantity === undefined;

            const initialQuantity = sinKardexSiguiente ? 0 : metadata.initialQuantity;

            const difference = Rule002Exporter.round(
                metadata.finalQuantity - initialQuantity!
            );

            rows.push({
                period: `${mesCierre} → ${mesSiguiente}`,
                productCode: result.product_code,
                productDescription: result.product_name,
                inconsistencyType: sinKardexSiguiente
                    ? "Producto no encontrado en el mes siguiente"
                    : "Continuidad de Cantidad",

                expectedValue: metadata.finalQuantity,
                foundValue: +initialQuantity!,

                difference,

                differencePercent:
                    metadata.finalQuantity === 0
                        ? 0
                        : Math.abs(difference / metadata.finalQuantity) * 100,

                riskLevel: result.risk_level,

                traceability: sinKardexSiguiente
                    ? [
                        `Mes de Cierre: ${mesCierre}`,
                        `Cantidad Saldo Final (${mesCierre}): ${metadata.finalQuantity}`,
                        `Mes Siguiente: ${mesSiguiente}`,
                        `El producto no tiene Kardex registrado en ${mesSiguiente} - no se puede validar la continuidad.`
                    ].join("\n")
                    : [
                        `Mes de Cierre: ${mesCierre}`,
                        `Cantidad Saldo Final (${mesCierre}): ${metadata.finalQuantity}`,
                        `Mes Siguiente: ${mesSiguiente}`,
                        `Cantidad Saldo Inicial (${mesSiguiente}): ${initialQuantity}`,
                        `Diferencia (Cantidad): ${difference}`
                    ].join("\n")
            });
        }

        return rows;
    }

    private static round(value: number): number {
        return Math.round(value * 100) / 100;
    }
}