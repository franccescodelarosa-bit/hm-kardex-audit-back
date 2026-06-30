import ExcelJS from "exceljs";

export interface Rule004Metadata {
    transitItem: string;
    issueDate: string;
    warehouseDate: string;
    supplierRuc: string;
    supplier: string;
    document: string;
    normalizedDocument: string;
}

export class Rule004Exporter {

    async export(results: any[]) {

        const workbook = new ExcelJS.Workbook();

        workbook.creator = "HM Kardex Audit";
        workbook.created = new Date();

        const worksheet = workbook.addWorksheet("RULE_004");

        worksheet.mergeCells("A1:I1");

        worksheet.getCell("A1").value =
            "RULE_004 - Mercadería en Tránsito no Registrada en el Kardex";

        worksheet.getCell("A1").font = {
            bold: true,
            size: 16
        };

        // Espacio entre título y tabla
        worksheet.addRow([]);

        // Cabeceras
        worksheet.addRow([
            "Documento",
            "Documento Normalizado",
            "Item",
            "Fecha Emisión",
            "Fecha Almacén",
            "RUC Proveedor",
            "Proveedor",
            "Descripción",
            "Recomendación"
        ]);

        // Anchos
        worksheet.getColumn(1).width = 22;
        worksheet.getColumn(2).width = 24;
        worksheet.getColumn(3).width = 15;
        worksheet.getColumn(4).width = 18;
        worksheet.getColumn(5).width = 18;
        worksheet.getColumn(6).width = 18;
        worksheet.getColumn(7).width = 40;
        worksheet.getColumn(8).width = 55;
        worksheet.getColumn(9).width = 55;

        // Datos
        for (const result of results) {

            const metadata = result.metadata as Rule004Metadata;

            worksheet.addRow([
                metadata.document,
                metadata.normalizedDocument,
                metadata.transitItem,
                metadata.issueDate,
                metadata.warehouseDate,
                metadata.supplierRuc,
                metadata.supplier,
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
            to: "I3"
        };

        return workbook;

    }

}