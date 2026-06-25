import {
  Body,
  Controller,
  Get,
  Param,
  Post,
} from '@nestjs/common';

import { AuditsService }
from './audits.service';

import { CreateAuditDto }
from './dto/create-audit.dto';

@Controller('audits')
export class AuditsController {

  constructor(
    private readonly auditsService: AuditsService,
  ) {}

  @Post()
  create(
    @Body()
    createAuditDto: CreateAuditDto,
  ) {
    return this.auditsService.create(
      createAuditDto,
    );
  }

  @Post(':id/run')
  runAudit(
    @Param('id')
    id: string,
  ) {
    return this.auditsService.run(id);
  }

  @Get()
  findAll() {

    return this.auditsService.findAll();

  }

  @Get(':id')
  findOne(
    @Param('id')
    id: string,
  ) {

    return this.auditsService.findOne(
      id,
    );

  }

}