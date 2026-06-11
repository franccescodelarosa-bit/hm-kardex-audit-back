import { Module } from '@nestjs/common';
import { ConfigModule } from '@nestjs/config';
import { AppController } from './app.controller';
import { AppService } from './app.service';
import { PrismaModule } from './prisma/prisma.module';
import { AuthModule } from './auth/auth.module';
import { UsersModule } from './users/users.module';
import { AdminModule } from './admin/admin.module';
import { CompaniesModule } from './companies/companies.module';
import { UploadsModule } from './uploads/uploads.module';
import { AuditsModule } from './audits/audits.module';
import { AccessLogsModule } from './access-logs/access-logs.module';

@Module({
  imports: [PrismaModule, AuthModule, UsersModule, AdminModule, CompaniesModule, UploadsModule, AuditsModule, AccessLogsModule],
  controllers: [AppController],
  providers: [AppService],
})
export class AppModule {}
