import {
    Controller,
    Get,
    Param,
    Query
} from '@nestjs/common';
import { AuditResultsService } from './audit-results.service';
import { FindingsQueryDto } from './dto/findings-query.dto';

@Controller("auditsresult")
export class AuditResultsController {

    constructor(
        private readonly service: AuditResultsService,
    ) {}

    @Get("")
    getAudits(){
        return this.service.getAudits();
    }

    @Get(':id/dashboard')
    async getDashboard(
        @Param('id') id: string,
    ) {
        return this.service.getDashboard(id);
    }
    @Get(':id/rules')
    async getRules(
        @Param('id') id: string,
    ) {
        return this.service.getRules(id);
    }

    @Get(":id/findings")
    async getFindings(

        @Param("id") id: string,

        @Query() dto: FindingsQueryDto

    ) {

        return this.service.getFindings(id, dto);

    }
    @Get("findings/:id")
    async getFinding(
        @Param("id") id: string
    ) {

        return this.service.getFinding(id);

    }
}