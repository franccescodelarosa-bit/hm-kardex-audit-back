import {
  Injectable,
  BadRequestException,
} from '@nestjs/common';

import { PrismaService }
from '../prisma/prisma.service';

import {
  CognitoIdentityProviderClient,
  AdminCreateUserCommand,
  AdminGetUserCommand,
} from '@aws-sdk/client-cognito-identity-provider';

@Injectable()
export class UsersService {

  constructor(
    private readonly prisma: PrismaService,
  ) {}

  private cognito =
    new CognitoIdentityProviderClient({
      region:
        process.env.AWS_REGION,
    });

  async create(body: any) {
    try {

      const response =
        await this.cognito.send(
          new AdminCreateUserCommand({
            UserPoolId:
              process.env.COGNITO_USER_POOL_ID,

            Username:
              body.email,

            UserAttributes: [
              {
                Name: 'email',
                Value: body.email,
              },
              {
                Name: 'email_verified',
                Value: 'true',
              },
              {
                Name: 'name',
                Value:
                  body.fullName,
              },
            ],
          }),
        );

      const cognitoUser =
        await this.cognito.send(
          new AdminGetUserCommand({
            UserPoolId:
              process.env.COGNITO_USER_POOL_ID,

            Username:
              body.email,
          }),
        );

      const sub =
        cognitoUser.UserAttributes?.find(
          (x) => x.Name === 'sub',
        )?.Value;

      const user =
        await this.prisma.users.create({
          data: {
            cognito_sub:
              sub,

            email:
              body.email,

            full_name:
              body.fullName,

            role:
              'AUDITOR',

            status:
              'PENDING',
          },
        });

      console.log(
        'USER CREATED',
        user,
      );

      return {
        success: true,
        message:
          'Usuario creado correctamente',
        user,
      };

    } catch (error: any) {

      console.error(error);

      throw new BadRequestException(
        error.message,
      );
    }
  }

  async findAll() {

    return this.prisma.users.findMany({
      orderBy: {
        created_at:
          'desc',
      },
    });

  }
}