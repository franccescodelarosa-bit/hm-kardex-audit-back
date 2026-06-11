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

  async me(token: string) {
    try {
      const decoded: any =
        jwt.decode(token);

      const sub =
        decoded?.sub;

      if (!sub) {
        throw new Error();
      }

      const user =
        await this.prisma.users.findFirst({
          where: {
            cognito_sub: sub,
          },
        });

      if (!user) {
        throw new UnauthorizedException(
          'Usuario no encontrado'
        );
      }

      return user;
    } catch (error) {
      throw new UnauthorizedException(
        'Token inválido'
      );
    }
  }
}