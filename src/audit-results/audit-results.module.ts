import { Module } from '@nestjs/common';
import { AuditResultsController } from './audit-results.controller';
import { AuditResultsService } from './audit-results.service';
import { AuditResultsRepository } from './repositories/audit-results.repository';
import { PrismaModule } from '../prisma/prisma.module';

@Module({
  imports: [
    PrismaModule,
  ],
  controllers: [AuditResultsController],
  providers: [AuditResultsService, AuditResultsRepository]
})
export class AuditResultsModule {}
