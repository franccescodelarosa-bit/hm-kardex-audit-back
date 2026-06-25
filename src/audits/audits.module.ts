import { Module } from '@nestjs/common';
import { AuditsController } from './audits.controller';
import { AuditsService } from './audits.service';
import { PrismaModule } from '../prisma/prisma.module';
import { AuditEngineModule } from '../audit-engine/audit-engine.module';
@Module({
  imports: [ PrismaModule, AuditEngineModule ],
  controllers: [ AuditsController ],
  providers: [ AuditsService ],
})
export class AuditsModule {}