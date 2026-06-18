import { Module } from '@nestjs/common';
import { PassportModule } from '@nestjs/passport';

import { CognitoStrategy }
from './strategies/cognito.strategy';
import { AuthController } from './auth.controller';
import { AuthService } from './auth.service';

import { PrismaModule } from '../prisma/prisma.module';

@Module({
  imports: [
    PassportModule.register({
      defaultStrategy: 'jwt',
    }),
    PrismaModule],

  controllers: [AuthController],

  providers: [AuthService, CognitoStrategy],
})
export class AuthModule {}