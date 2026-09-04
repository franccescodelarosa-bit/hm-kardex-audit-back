import { Injectable, BadRequestException, Logger } from '@nestjs/common';
import { S3Client, PutObjectCommand, GetObjectCommand } from '@aws-sdk/client-s3';
import { getSignedUrl } from '@aws-sdk/s3-request-presigner';
import { ZipArchive } from 'archiver';
import { AuditResultsRepository } from './repositories/audit-results.repository';
import { ReportHeader } from '../excel/exporters/base/ReportHeader';
import { FindingsQueryDto } from './dto/findings-query.dto';
import { UpdateAuditFollowUpDto } from './dto/update-audit.dto';
import { Rule001Exporter } from '../excel/exporters/Rule001Exporter';
import { Rule002Exporter } from '../excel/exporters/Rule002Exporter';
import { Rule003Exporter } from '../excel/exporters/Rule003Exporter';
import { Rule004Exporter } from '../excel/exporters/Rule004Exporter';
import { Rule005Exporter } from '../excel/exporters/Rule005Exporter';
import { Rule006Exporter } from '../excel/exporters/Rule006Exporter';
import { Rule007Exporter } from '../excel/exporters/Rule007Exporter';
import { Rule008Exporter } from '../excel/exporters/Rule008Exporter';
import { Rule009Exporter } from '../excel/exporters/Rule009Exporter';
import { Rule010Exporter } from '../excel/exporters/Rule010Exporter';
import { Rule011Exporter } from '../excel/exporters/Rule011Exporter';
import { Rule012Exporter } from '../excel/exporters/Rule012Exporter';
import { Rule013Exporter } from '../excel/exporters/Rule013Exporter';
import { Rule014Exporter } from '../excel/exporters/Rule014Exporter';

const RULE_CODES = [
    "RULE_001", "RULE_002", "RULE_003", "RULE_004", "RULE_005",
    "RULE_006", "RULE_007", "RULE_008", "RULE_009", "RULE_010",
    "RULE_011", "RULE_012", "RULE_013", "RULE_014"
] as const;

@Injectable()
export class AuditResultsService {

    private readonly logger = new Logger(AuditResultsService.name);

    private readonly s3 = new S3Client({
        region: process.env.AWS_REGION
    });

    constructor(
        private readonly repository: AuditResultsRepository
    ){}
    async getAudits(){
        return this.repository.getAudits();
    }
    async getDashboard(auditJobId: string) {
        return this.repository.getDashboard(auditJobId);
    }
    async getRules(auditJobId: string) {
        return this.repository.getRules(auditJobId);
    }
    async getFindings(
        auditJobId: string,
        dto: FindingsQueryDto
    ) {
        return this.repository.getFindings(auditJobId, dto);

    }
    async getFinding(id: string) {
        return this.repository.getFinding(id);
    }
    async putAuditResult(id: string, dto: UpdateAuditFollowUpDto) {
        return this.repository.putAuditResult(id, dto);
    }

    //todo: armar el libro
    private async buildWorkbookForRule(auditJobId: string, ruleId: string) {
        const { header, rows } = await this.repository.getExcel(auditJobId, ruleId);
        if (!rows.length) {
            return null;
        }
        switch (rows[0].audit_rules!.code) {
            case "RULE_001":
                return new Rule001Exporter().export(rows, header!);
            case "RULE_002":
                return new Rule002Exporter().export(rows, header!);
            case "RULE_003":
                return new Rule003Exporter().export(rows, header!);
            case "RULE_004":
                return new Rule004Exporter().export(rows, header!);
            case "RULE_005":
                return new Rule005Exporter().export(rows, header!);
            case "RULE_006":
                return new Rule006Exporter().export(rows, header!);
            case "RULE_007":
                return new Rule007Exporter().export(rows, header!);
            case "RULE_008":
                return new Rule008Exporter().export(rows, header!);
            case "RULE_009":
                return new Rule009Exporter().export(rows, header!);
            case "RULE_010":
                return new Rule010Exporter().export(rows, header!);
            case "RULE_011":
                return new Rule011Exporter().export(rows, header!);
            case "RULE_012":
                return new Rule012Exporter().export(rows, header!);
            case "RULE_013":
                return new Rule013Exporter().export(rows, header!);
            case "RULE_014":
                return new Rule014Exporter().export(rows, header!);
            default:
                throw new BadRequestException(
                    `Exportador no implementado para la regla ${rows[0].audit_rules!.code}`
                );
        }
    }

    async getExcel(auditJobId: string, ruleId: string) {
        const workbook = await this.buildWorkbookForRule(auditJobId, ruleId);
        if (!workbook) {
            throw new BadRequestException("No existen resultados para la regla.");
        }
        return workbook;
    }

    //todo: subir el zip o excel al S3
    private async uploadReportBuffer(
        reportId: string,
        s3Key: string,
        buffer: Buffer,
        contentType: string
    ) {
        await this.s3.send(new PutObjectCommand({
            Bucket: process.env.AWS_S3_BUCKET,
            Key: s3Key,
            Body: buffer,
            ContentType: contentType
        }));
        await this.repository.markReportReady(reportId, s3Key);
    }

    //todo: ejecutar el build y mapear el error en caso falle la exportacion
    private async runBackgroundReport(
        auditJobId: string,
        reportId: string,
        build: () => Promise<void>
    ) {
        try {
            await build();
        } catch (error: any) {
            await this.repository.markReportError(
                reportId,
                error?.message ?? "Error desconocido generando el reporte."
            );
            this.logger.error(
                `Fallo generando el reporte ${reportId} de la auditoria ${auditJobId}`,
                error
            );
        }
    }

    //todo: zip de las reglas -> 14
    async startZipReport(auditJobId: string) {
        const report = await this.repository.createReport(auditJobId, "FULL_ZIP");

        this.runBackgroundReport(
            auditJobId,
            report.id,
            () => this.buildZipInBackground(auditJobId, report.id)
        );

        return { reportId: report.id, status: report.status };
    }

    //todo: armar el zip
    private async buildZipInBackground(auditJobId: string, reportId: string) {
        //iteramos sobre las reglas de la auditoria
        const rulesByCode = new Map(
            (await this.repository.getRules(auditJobId)).map(r => [r.code, r])
        );

        const archive = new ZipArchive({ zlib: { level: 9 } });
        const chunks: Buffer[] = [];
        archive.on("data", chunk => chunks.push(chunk));
        const archiveFinished = new Promise<void>((resolve, reject) => {
            archive.on("end", () => resolve());
            archive.on("error", reject);
        });

        let includedRules = 0;
        for (const code of RULE_CODES) {
            const rule = rulesByCode.get(code as string);
            //si no hay regla continuamos :D
            if (!rule?.id) {
                continue;
            }
            const workbook = await this.buildWorkbookForRule(auditJobId, rule.id);
            if (!workbook) {
                continue;
            }
            const buffer = await workbook.xlsx.writeBuffer();
            archive.append(Buffer.from(buffer), { name: `${code}.xlsx` });
            includedRules++;
        }

        if (includedRules === 0) {
            throw new BadRequestException(
                "La auditoría no tiene hallazgos en ninguna regla - no hay nada que exportar."
            );
        }

        await archive.finalize();
        await archiveFinished;

        const zipBuffer = Buffer.concat(chunks);
        const s3Key = `audits/reports/${auditJobId}/${reportId}.zip`;

        await this.uploadReportBuffer(reportId, s3Key, zipBuffer, "application/zip");
    }

    //todo:generar el reporte par auna sola regla .xlsx
    async startRuleReport(auditJobId: string, ruleId: string) {
        const report = await this.repository.createReport(auditJobId, "SINGLE_RULE");

        this.runBackgroundReport(
            auditJobId,
            report.id,
            () => this.buildRuleReportInBackground(auditJobId, ruleId, report.id)
        );

        return { reportId: report.id, status: report.status };
    }

    //todo: armar el workbook 
    private async buildRuleReportInBackground(
        auditJobId: string,
        ruleId: string,
        reportId: string
    ) {
        const workbook = await this.buildWorkbookForRule(auditJobId, ruleId);
        if (!workbook) {
            throw new BadRequestException("No existen resultados para la regla.");
        }
        const buffer = await workbook.xlsx.writeBuffer();
        const s3Key = `audits/reports/${auditJobId}/${reportId}.xlsx`;

        await this.uploadReportBuffer(
            reportId,
            s3Key,
            Buffer.from(buffer),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        );
    }

    //todo: obtener el estado del reporte
    async getReportStatus(reportId: string) {
        const report = await this.repository.getReport(reportId);
        if (!report) {
            throw new BadRequestException("Reporte no encontrado.");
        }
        if (report.status !== "READY" || !report.s3_key) {
            return {
                reportId: report.id,
                status: report.status,
                errorMessage: report.error_message ?? undefined
            };
        }
        const downloadUrl = await getSignedUrl(
            this.s3,
            new GetObjectCommand({
                Bucket: process.env.AWS_S3_BUCKET,
                Key: report.s3_key
            }),
            { expiresIn: 300 }
        );
        return {
            reportId: report.id,
            status: report.status,
            downloadUrl
        };
    }
}
