import {
  Injectable,
} from '@nestjs/common';

import { PrismaService }
from '../prisma/prisma.service';

import { CreateClientDto }
from './dto/create-client.dto';

import { UpdateClientDto }
from './dto/update-client.dto';

@Injectable()
export class ClientsService {

  constructor(
    private readonly prisma: PrismaService,
  ) {}

  async create(
    createClientDto: CreateClientDto,
  ) {

    const company =
      await this.prisma.companies.create({
        data: {
          name:
            createClientDto.companyName,

          ruc:
            createClientDto.ruc,

          email:
            createClientDto.email,

          phone:
            createClientDto.phone,

          status:
            'ACTIVE',
        },
      });

    return this.prisma.clients.create({
      data: {

        company_id:
          company.id,

        business_name:
          createClientDto.businessName,

        ruc:
          createClientDto.ruc,

        email:
          createClientDto.email,

        phone:
          createClientDto.phone,

        status:
          'ACTIVE',
      },
    });

  }

  async findAll() {

    return this.prisma.clients.findMany({
      orderBy: {
        created_at: 'desc',
      },
    });

  }

  async findOne(id: string) {

    return this.prisma.clients.findUnique({
      where: {
        id,
      },
    });

  }

  async update(
    id: string,
    updateClientDto: UpdateClientDto,
  ) {

    return this.prisma.clients.update({
      where: {
        id,
      },
      data: updateClientDto,
    });

  }

  async remove(id: string) {

    return this.prisma.clients.delete({
      where: {
        id,
      },
    });

  }
}