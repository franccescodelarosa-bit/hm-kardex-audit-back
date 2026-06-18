import {
  Body,
  Controller,
  Post,
  Get,
} from '@nestjs/common';

import { UsersService } from './users.service';

@Controller('users')
export class UsersController {
  constructor(
    private readonly usersService: UsersService,
  ) {}

  @Post()
  async create(
    @Body() body: any,
  ) {
    return this.usersService.create(
      body,
    );
  }

  @Get()
  async findAll() {
    return this.usersService.findAll();
  }
}