import ExcelJS from "exceljs";

import { BaseExcelExporter } from "./base/BaseExcelExporter";
import { ReportHeader } from "./base/ReportHeader";
import { DateUtils } from "../helpers/dateutils";

export interface Rule004EvaluatedProduct {
    code: string;
    description: string;
    cost: number;
}

export interface Rule004Metadata {
    transitItem: string;
    issueDate: string;
    warehouseDate: string;
    supplierRuc: string;
    supplier: string;
    document: string;
    normalizedDocument: string;
    month: string | null;
    expectedCost: number;
    foundCost: number;
    evaluatedProducts: Rule004EvaluatedProduct[];
    isIncident?: boolean;
    thresholdPercent?: number;
}

interface Rule004Row {
    period: string;
    issueDate: string;
    warehouseDate: string;
    supplierRuc: string;
    supplier: string;
    document: string;
    normalizedDocument: string;
    productCode: string;
    productDescription: string;
    inconsistencyType: string;
    expectedValue: number;
    foundValue: number;
    difference: number;
    differencePercent: number;
    riskLevel: string;
    traceability: string;
}

export class Rule004Exporter extends BaseExcelExporter {
    async export(
        results: any[],
        header: ReportHeader
    ) {
        const workbook = new ExcelJS.Workbook();
        workbook.creator = "HM Kardex Audit";
        workbook.created = new Date();
        const worksheet = workbook.addWorksheet("RULE_004");
        this.writeHeader(
            worksheet,
            "RULE_004 - Validación de facturas de mercadería en transito al cierre del año",
            header,
            "P"
        );
        this.writeRule004TableHeader(worksheet);
        const findings = this.buildFindings(results);
        this.writeRule004Rows(worksheet, findings);
        worksheet.views = [
            {
                state: "frozen",
                ySplit: 4
            }
        ];
        worksheet.autoFilter = {
            from: "A4",
            to: "P4"
        };
        return workbook;
    }

    private writeRule004TableHeader(worksheet: ExcelJS.Worksheet) {
        worksheet.addRow([
            "Periodo",
            "Fecha de Emisión",
            "Fecha de Ingreso a Almacén",
            "RUC Proveedor",
            "Proveedor",
            "Documento",
            "Documento Normalizado",
            "Código del producto",
            "Descripción del producto",
            "Tipo de inconsistencia",
            "Valor esperado",
            "Valor encontrado",
            "Diferencia",
            "% Diferencia",
            "Nivel de riesgo",
            "Trazabilidad"
        ]);
        const headerRow = worksheet.getRow(4);
        headerRow.font = { bold: true };
        headerRow.alignment = { vertical: "middle", horizontal: "center" };
        headerRow.eachCell(cell => {
            cell.fill = {
                type: "pattern",
                pattern: "solid",
                fgColor: { argb: "D9EAD3" }
            };
            cell.border = {
                top: { style: "thin" },
                left: { style: "thin" },
                bottom: { style: "thin" },
                right: { style: "thin" }
            };
        });
        worksheet.columns = [
            { width: 14 }, // Periodo
            { width: 14 }, // Fecha Emisión
            { width: 16 }, // Fecha Almacén
            { width: 16 }, // RUC Proveedor
            { width: 28 }, // Proveedor
            { width: 20 }, // Documento
            { width: 20 }, // Documento Normalizado
            { width: 22 }, // Código del producto
            { width: 40 }, // Descripción
            { width: 30 }, // Tipo
            { width: 18 }, // Esperado
            { width: 18 }, // Encontrado
            { width: 16 }, // Diferencia
            { width: 14 }, // %
            { width: 16 }, // Riesgo
            { width: 50 }  // Trazabilidad
        ];
    }

    private writeRule004Rows(worksheet: ExcelJS.Worksheet, rows: Rule004Row[]) {
        for (const row of rows) {
            const excelRow = worksheet.addRow([
                row.period,
                row.issueDate,
                row.warehouseDate,
                row.supplierRuc,
                row.supplier,
                row.document,
                row.normalizedDocument,
                row.productCode,
                row.productDescription,
                row.inconsistencyType,
                row.expectedValue,
                row.foundValue,
                row.difference,
                `${row.differencePercent.toFixed(2)} %`,
                row.riskLevel,
                row.traceability
            ]);
            excelRow.eachCell(cell => {
                cell.alignment = { vertical: "top" };
            });
            excelRow.getCell(16).alignment = {
                wrapText: true,
                vertical: "top"
            };
        }
    }

    private buildFindings(results: any[]): Rule004Row[] {
        const rows: Rule004Row[] = [];
        for (const result of results) {
            const metadata = result.metadata as Rule004Metadata;
            const evaluatedProducts = metadata.evaluatedProducts ?? [];
            const productosEncontrados = this.buildProductosEncontrados(metadata);

            const esValidacionDeCosto = metadata.isIncident !== undefined;

            rows.push({
                period: metadata.month ? DateUtils.monthName(Number(metadata.month)) : "Sin período",
                issueDate: metadata.issueDate,
                warehouseDate: metadata.warehouseDate,
                supplierRuc: metadata.supplierRuc,
                supplier: metadata.supplier,
                document: metadata.document,
                normalizedDocument: metadata.normalizedDocument,
                productCode: evaluatedProducts.map(p => p.code).join(", "),
                productDescription: evaluatedProducts.map(p => p.description).join(", "),
                inconsistencyType: esValidacionDeCosto
                    ? (metadata.isIncident ? "INCIDENCIA" : "ACEPTADA")
                    : "Mercadería en tránsito no registrada",
                expectedValue: metadata.expectedCost,
                foundValue: metadata.foundCost,
                difference: metadata.expectedCost - metadata.foundCost,
                differencePercent:
                    metadata.expectedCost === 0
                        ? 0
                        : Math.abs(
                            (
                                (metadata.expectedCost - metadata.foundCost) /
                                metadata.expectedCost
                            ) * 100
                        ),
                riskLevel: result.risk_level,
                traceability: [
                    `Productos Encontrados: ${productosEncontrados}`,
                    `Monto Encontrado (suma de esos productos): ${metadata.foundCost}`,
                    ...(esValidacionDeCosto
                        ? [
                            `Umbral permitido: ${metadata.thresholdPercent}%`,
                            `Resultado: ${metadata.isIncident ? "INCIDENCIA" : "ACEPTADA"}`
                        ]
                        : [])
                ].join("\n")
            });
        }
        return rows;
    }

    private buildProductosEncontrados(metadata: Rule004Metadata): string {
        return (metadata.evaluatedProducts ?? [])
            .map(product => `${product.code} - ${product.description}`)
            .join(", ");
    }
}
