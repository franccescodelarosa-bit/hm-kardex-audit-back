import { Module } from '@nestjs/common';
import { ClientsService } from './clients.service';
import { PrismaModule }
from '../prisma/prisma.module';
import { ClientsController } from './clients.controller';

@Module({
  imports: [PrismaModule],
  controllers: [ClientsController],
  providers: [ClientsService],
})
export class ClientsModule {}
