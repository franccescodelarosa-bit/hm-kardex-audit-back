import ExcelJS from "exceljs";

export interface Rule004Metadata {
    transitItem: string;
    issueDate: string;
    warehouseDate: string;
    supplierRuc: string;
    supplier: string;
    document: string;
    normalizedDocument: string;
    month: string | null;
    duplicatedItems: number;
    products: string;
}

export class Rule004Exporter {

    async export(results: any[]) {

        const workbook = new ExcelJS.Workbook();

        workbook.creator = "HM Kardex Audit";
        workbook.created = new Date();

        const worksheet = workbook.addWorksheet("RULE_004");

        worksheet.mergeCells("A1:L1");

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
            "Periodo Kardex",
            "Items Encontrados",
            "Productos Encontrados",
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
        worksheet.getColumn(3).width = 12;
        worksheet.getColumn(4).width = 18;
        worksheet.getColumn(5).width = 60;
        worksheet.getColumn(6).width = 15;
        worksheet.getColumn(7).width = 18;
        worksheet.getColumn(8).width = 18;
        worksheet.getColumn(9).width = 18;
        worksheet.getColumn(10).width = 40;
        worksheet.getColumn(11).width = 55;
        worksheet.getColumn(12).width = 55;

        // Datos
        for (const result of results) {

            const metadata = result.metadata as Rule004Metadata;

            worksheet.addRow([
                metadata.document,
                metadata.normalizedDocument,
                metadata.month,
                metadata.duplicatedItems,
                metadata.products,
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
            to: "L3"
        };

        return workbook;

    }

}