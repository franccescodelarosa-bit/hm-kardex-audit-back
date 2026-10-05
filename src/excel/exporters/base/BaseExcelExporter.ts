import ExcelJS from "exceljs";
import { ReportHeader } from "./ReportHeader";
import { AuditFindingRow } from "./AuditFindingRow";
import { ColumnStyle, ReportStyle, styleFor } from "./ReportStyle";
import { DateUtils } from "../../helpers/dateutils";

export type CellContent = string | number | null | undefined | ExcelJS.CellRichTextValue;

const FONT_NAME = "Calibri";

const THIN = { style: "thin" } as const;
const BORDER: Partial<ExcelJS.Borders> = {
    top: THIN,
    left: THIN,
    bottom: THIN,
    right: THIN
};

const STAMP_COLOR = "FFFF0000";
const STAMP_FILL = "FFFFFF00";

const HEADER_ROW = 3;

const STANDARD_LABELS = [
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
];

export abstract class BaseExcelExporter {

    protected writeHeader(
        worksheet: ExcelJS.Worksheet,
        title: string,
        header: ReportHeader,
        lastColumn: string
    ) {
        const style = styleFor(worksheet.name);

        worksheet.mergeCells(`A1:${lastColumn}1`);

        const titleCell = worksheet.getCell("A1");
        titleCell.value = title;
        titleCell.font = { name: FONT_NAME, size: style.titleSize, bold: true };
        titleCell.alignment = { vertical: "middle", wrapText: true };
        worksheet.getRow(1).height = style.titleHeight;

        this.writeInfoCell(worksheet, "A2", "Empresa:", style);
        this.writeInfoCell(worksheet, "B2", header.companyName, style, style.companyAlign === "center" ? "center" : undefined, style.companySize);
        this.writeInfoCell(worksheet, "C2", "RUC:", style);
        this.writeInfoCell(worksheet, "D2", header.ruc, style);
        this.writeInfoCell(worksheet, "E2", "Periodo:", style);
        this.writeInfoCell(worksheet, "F2", header.year, style, "center");
        worksheet.getRow(2).height = 18.75;

        const stampCell = worksheet.getCell(`${lastColumn}2`);
        stampCell.value = `AUDITORIA ${DateUtils.dottedDate(header.auditDate ?? new Date())}`;
        stampCell.font = { name: FONT_NAME, size: 14, bold: true, color: { argb: STAMP_COLOR } };
        stampCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: STAMP_FILL } };
        stampCell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    }

    private writeInfoCell(
        worksheet: ExcelJS.Worksheet,
        address: string,
        value: string | number,
        style: ReportStyle,
        horizontal?: "center",
        size: number = style.infoSize
    ) {
        const cell = worksheet.getCell(address);
        cell.value = value;
        cell.font = { name: FONT_NAME, size, bold: true };
        cell.alignment = horizontal
            ? { horizontal, vertical: "middle" }
            : { vertical: "middle" };
    }

    protected writeTableHeader(
        worksheet: ExcelJS.Worksheet,
        labels: string[] = STANDARD_LABELS
    ) {
        const style = styleFor(worksheet.name);

        this.writeHeaderRow(worksheet, HEADER_ROW, style, style.headerHeight, column => labels[column - 1]);

        let headerRows = HEADER_ROW;

        if (style.subHeader) {
            const { height, labels: subLabels } = style.subHeader;
            headerRows = HEADER_ROW + 1;
            this.writeHeaderRow(worksheet, headerRows, style, height, column => subLabels[column] ?? null);
            worksheet.autoFilter = `A${headerRows}:${worksheet.getColumn(style.columns.length).letter}${headerRows}`;
        }

        style.columns.forEach((column, index) => {
            worksheet.getColumn(index + 1).width = column.width;
        });

        worksheet.views = [{ state: "frozen", ySplit: headerRows }];
    }

    private writeHeaderRow(
        worksheet: ExcelJS.Worksheet,
        rowNumber: number,
        style: ReportStyle,
        height: number,
        labelOf: (column: number) => string | null | undefined
    ) {
        const row = worksheet.getRow(rowNumber);
        row.height = height;
        style.columns.forEach((column, index) => {
            const cell = row.getCell(index + 1);
            const label = labelOf(index + 1);
            if (label !== null && label !== undefined) {
                cell.value = label;
            }
            cell.font = {
                name: FONT_NAME,
                size: column.headerSize ?? style.headerSize,
                bold: true
            };
            cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
            cell.border = BORDER;
        });
    }

    protected writeRows(
        worksheet: ExcelJS.Worksheet,
        rows: AuditFindingRow[]
    ) {
        this.writeTableRows(
            worksheet,
            rows.map(row => [
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
            ])
        );
    }

    protected writeTableRows(
        worksheet: ExcelJS.Worksheet,
        rows: CellContent[][]
    ) {
        const style = styleFor(worksheet.name);
        const cellStyles = style.columns.map(column => this.cellStyleOf(column));

        for (const values of rows) {
            const excelRow = worksheet.addRow(values);
            cellStyles.forEach((cellStyle, index) => {
                const cell = excelRow.getCell(index + 1);
                cell.font = cellStyle.font;
                cell.alignment = cellStyle.alignment;
                cell.border = BORDER;
                if (cellStyle.numFmt) {
                    cell.numFmt = cellStyle.numFmt;
                }
            });
        }
    }

    private cellStyleOf(column: ColumnStyle) {
        return {
            font: { name: FONT_NAME, size: column.size, bold: true },
            alignment: column.align === "center"
                ? { horizontal: "center", vertical: "middle", wrapText: true } as const
                : { vertical: "middle", wrapText: true } as const,
            numFmt: column.numFmt
        };
    }
}
