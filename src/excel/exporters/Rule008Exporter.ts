import ExcelJS from "exceljs";

import { BaseExcelExporter } from "./base/BaseExcelExporter";
import { ReportHeader } from "./base/ReportHeader";
import { AuditFindingRow } from "./base/AuditFindingRow";
import { DateUtils } from "../helpers/dateutils";

export interface Rule008Metadata {
    source: string;
    month: number;
    expectedValue?: string;
    foundValue?: string;
    stock?: number;
    date?: string;
    document?: string;
}
export class Rule008Exporter extends BaseExcelExporter {
    async export(
        results: any[],
        header: ReportHeader
    ) {
        const workbook = new ExcelJS.Workbook();
        workbook.creator = "HM Kardex Audit";
        workbook.created = new Date();
        const worksheet = workbook.addWorksheet("RULE_008");
        this.writeHeader(
            worksheet,
            "RULE_008 - Validación de códigos inexistentes al cierre del año",
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
            const metadata = result.metadata as Rule008Metadata;
            const fromInventory = metadata.source?.startsWith("INVENTARIO");
            rows.push({
                period: DateUtils.monthName(metadata.month),
                productCode: result.product_code,
                productDescription: result.product_name,
                inconsistencyType: fromInventory
                    ? "Código del inventario de cierre no existe en el Kardex de enero"
                    : "Código del Kardex de enero no existe en el inventario de cierre",
                expectedValue: metadata.expectedValue ?? result.product_code,
                foundValue: metadata.foundValue ?? (fromInventory
                    ? "Producto inexistente en el Kardex"
                    : "Producto inexistente en el inventario"),
                difference: "No aplica",
                differencePercent: undefined,
                riskLevel: result.risk_level,
                traceability: [
                    `Origen: ${fromInventory ? "Inventario al cierre del ejercicio anterior" : "Kardex de enero"}`,
                    `Código: ${result.product_code}`,
                    `Producto: ${result.product_name}`,
                    ...(metadata.stock !== undefined ? [`Stock en inventario: ${metadata.stock}`] : []),
                    ...(metadata.document ? [`Documento: ${metadata.document}`] : []),
                    ...(metadata.date ? [`Fecha: ${metadata.date}`] : [])
                ].join("\n")
            });
        }
        return rows;
    }
}