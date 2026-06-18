import {
  Injectable,
  UnauthorizedException,
} from '@nestjs/common';

import { PrismaService } from '../prisma/prisma.service';

import * as jwt from 'jsonwebtoken';

@Injectable()
export class AuthService {
  constructor(
    private readonly prisma: PrismaService
  ) {}

  async me(sub: string) {
    const user =
      await this.prisma.users.findFirst({
        where: {
          cognito_sub: sub,
        },
      });

    if (!user) {
      throw new UnauthorizedException(
        'Usuario no encontrado',
      );
    }
    if (
      user &&
      user.status === 'PENDING'
    ) {
      await this.prisma.users.update({
        where: {
          id: user.id,
        },
        data: {
          status: 'ACTIVE',
        },
      });

      user.status = 'ACTIVE';
    }

    return user;
  }
}