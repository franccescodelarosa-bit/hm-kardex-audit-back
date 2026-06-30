import ExcelJS from "exceljs";

export interface Rule006Metadata {
    source: string;
    occurrences: number;
    rows: number[];
}

export class Rule006Exporter {

    async export(results: any[]) {

        const workbook = new ExcelJS.Workbook();

        workbook.creator = "HM Kardex Audit";
        workbook.created = new Date();

        const worksheet = workbook.addWorksheet("RULE_006");

        worksheet.mergeCells("A1:G1");

        worksheet.getCell("A1").value =
            "RULE_006 - Detección de Productos Duplicados";

        worksheet.getCell("A1").font = {
            bold: true,
            size: 16
        };

        // Espacio
        worksheet.addRow([]);

        // Cabeceras
        worksheet.addRow([
            "Origen",
            "Código",
            "Producto",
            "Ocurrencias",
            "Filas",
            "Descripción",
            "Recomendación"
        ]);

        // Anchos
        worksheet.getColumn(1).width = 18;
        worksheet.getColumn(2).width = 18;
        worksheet.getColumn(3).width = 45;
        worksheet.getColumn(4).width = 15;
        worksheet.getColumn(5).width = 30;
        worksheet.getColumn(6).width = 60;
        worksheet.getColumn(7).width = 60;

        // Datos
        for (const result of results) {

            const metadata = result.metadata as Rule006Metadata;

            worksheet.addRow([
                metadata.source,
                result.product_code,
                result.product_name,
                metadata.occurrences,
                metadata.rows.join(", "),
                result.description,
                result.recommendation
            ]);

        }

        worksheet.views = [
            {
                state: "frozen",
                ySplit: 3
            }
        ];

        worksheet.autoFilter = {
            from: "A3",
            to: "G3"
        };

        return workbook;

    }

}