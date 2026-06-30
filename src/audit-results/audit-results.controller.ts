import {
    Controller,
    Get,
    Param,
    Query,
    Res
} from "@nestjs/common";
import type { Response } from "express";

import { AuditResultsService } from "./audit-results.service";
import { FindingsQueryDto } from "./dto/findings-query.dto";

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

    @Get(":id/rules/:ruleId/excel")
    async getExcel(
        @Param("id") auditJobId: string,
        @Param("ruleId") ruleId: string,
        @Res({ passthrough: true }) res: Response
    ) {

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

}