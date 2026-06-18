import {
  Controller,
  Get,
  Req,
  UseGuards,
} from '@nestjs/common';

import { Request } from 'express';

import { AuthService } from './auth.service';

import { JwtAuthGuard }
from './guards/jwt-auth.guard';

@Controller('auth')
export class AuthController {
  constructor(
    private readonly authService: AuthService,
  ) {}

  @Get('me')
  @UseGuards(JwtAuthGuard)
  async me(
    @Req() req: any,
  ) {
    const user =
      req.user as any;

    return this.authService.me(
      user.sub,
    );
  }
}