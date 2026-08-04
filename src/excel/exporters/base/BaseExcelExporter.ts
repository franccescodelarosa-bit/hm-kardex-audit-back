import ExcelJS from "exceljs";
import { ReportHeader } from "./ReportHeader";
import { AuditFindingRow } from "./AuditFindingRow";

export abstract class BaseExcelExporter {

    protected writeHeader(
        worksheet: ExcelJS.Worksheet,
        title: string,
        header: ReportHeader,
        lastColumn: string
    ) {

        worksheet.mergeCells(`A1:${lastColumn}1`);

        worksheet.getCell("A1").value = title;

        worksheet.getCell("A1").font = {
            bold: true,
            size: 16
        };

        worksheet.getCell("A2").value = "Empresa:";
        worksheet.getCell("B2").value = header.companyName;

        worksheet.getCell("C2").value = "RUC:";
        worksheet.getCell("D2").value = header.ruc;

        worksheet.getCell("E2").value = "Periodo:";
        worksheet.getCell("F2").value = header.year;
    }
    protected writeTableHeader(worksheet: ExcelJS.Worksheet) {
        worksheet.addRow([
            "Periodo",
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
        headerRow.font = {
            bold: true
        };
        headerRow.alignment = {
            vertical: "middle",
            horizontal: "center"
        };
        headerRow.eachCell(cell => {
            cell.fill = {
                type: "pattern",
                pattern: "solid",
                fgColor: {
                    argb: "D9EAD3"
                }
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
            { width: 18 }, // Código
            { width: 40 }, // Descripción
            { width: 30 }, // Tipo
            { width: 20 }, // Esperado
            { width: 20 }, // Encontrado
            { width: 18 }, // Diferencia
            { width: 16 }, // %
            { width: 16 }, // Riesgo
            { width: 50 }  // Trazabilidad
        ];
    }
    protected writeRows(
        worksheet: ExcelJS.Worksheet,
        rows: AuditFindingRow[]
    ) {
        
        for (const row of rows) {
            const excelRow = worksheet.addRow([
                row.period,
                row.productCode,
                row.productDescription,
                row.inconsistencyType,
                row.expectedValue,
                row.foundValue,
                row.difference,
                row.differencePercent != null ? `${row.differencePercent.toFixed(2)} %` : "",
                row.riskLevel,
                row.traceability
            ]);
            excelRow.eachCell(cell => {
                cell.alignment = {
                    vertical: "top"
                };
            });
            excelRow.getCell(10).alignment = {
                wrapText: true,
                vertical: "top"
            };
        }

    }
}