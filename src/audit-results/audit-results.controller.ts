import {
    Controller,
    Get,
    Post,
    Put,
    Param,
    Body,
    Query,
    Res
} from "@nestjs/common";
import type { Response } from "express";

import { AuditResultsService } from "./audit-results.service";
import { FindingsQueryDto } from "./dto/findings-query.dto";
import { UpdateAuditFollowUpDto } from './dto/update-audit.dto';

@Controller("auditsresult")
export class AuditResultsController {

    constructor(
        private readonly service: AuditResultsService,
    ) {}

    @Get("")
    getAudits() {
        return this.service.getAudits();
    }

    @Get(":id/dashboard")
    getDashboard(
        @Param("id") id: string,
    ) {
        return this.service.getDashboard(id);
    }

    @Get(":id/rules")
    getRules(
        @Param("id") id: string,
    ) {
        return this.service.getRules(id);
    }

    @Get(":id/findings")
    getFindings(
        @Param("id") id: string,
        @Query() dto: FindingsQueryDto
    ) {
        return this.service.getFindings(id, dto);
    }

    @Get("findings/:id")
    getFinding(
        @Param("id") id: string
    ) {
        return this.service.getFinding(id);
    }

    @Put(":id/follow-up")
    putAuditResult(
        @Param("id") id: string,
        @Body("dto") dto: UpdateAuditFollowUpDto,
    ) {
        console.log(dto);
        return this.service.putAuditResult(id, dto);
    }

    @Get(":id/rules/:ruleId/excel")
    async getExcel(
        @Param("id") auditJobId: string,
        @Param("ruleId") ruleId: string,
        @Res({ passthrough: true }) res: Response
    ) {
        console.log("CONTROLLER");
        const workbook = await this.service.getExcel(
            auditJobId,
            ruleId
        );

        const buffer = await workbook.xlsx.writeBuffer();
        console.log(buffer.byteLength);
        res.setHeader(
            "Content-Type",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        );

        res.setHeader(
            "Content-Disposition",
            `attachment; filename=${ruleId}.xlsx`
        );
        res.send(Buffer.from(buffer));
    }

    //todo: disparar la generacion el zip
    @Post(":id/reports/zip")
    startZipReport(
        @Param("id") auditJobId: string
    ) {
        return this.service.startZipReport(auditJobId);
    }

    //todo: polling para una regla
    @Post(":id/rules/:ruleId/reports")
    startRuleReport(
        @Param("id") auditJobId: string,
        @Param("ruleId") ruleId: string
    ) {
        return this.service.startRuleReport(auditJobId, ruleId);
    }

    //todo: verificar si ya está ok el archivo
    @Get("reports/:reportId")
    getReportStatus(
        @Param("reportId") reportId: string
    ) {
        return this.service.getReportStatus(reportId);
    }

}