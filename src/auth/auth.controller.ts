import {
  Controller,
  Get,
  Headers,
  UnauthorizedException,
} from '@nestjs/common';

import { AuthService } from './auth.service';

@Controller('auth')
export class AuthController {
  constructor(
    private readonly authService: AuthService
  ) {}

  @Get('me')
  async me(
    @Headers('authorization')
    authHeader?: string
  ) {
    if (!authHeader) {
      throw new UnauthorizedException(
        'Token requerido'
      );
    }

    const token =
      authHeader.replace(
        'Bearer ',
        ''
      );

    return this.authService.me(
      token
    );
  }
}