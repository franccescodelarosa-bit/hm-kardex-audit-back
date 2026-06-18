import { Injectable } from '@nestjs/common';
import { PassportStrategy } from '@nestjs/passport';

import { Strategy, ExtractJwt } from 'passport-jwt';

import * as jwksRsa from 'jwks-rsa';
const COGNITO_URL = `https://cognito-idp.us-east-1.amazonaws.com/${process.env.COGNITO_USER_POOL_ID}`;

@Injectable()
export class CognitoStrategy extends PassportStrategy(
  Strategy,
  'jwt',
) {
  constructor() {
    super({
      jwtFromRequest:
        ExtractJwt.fromAuthHeaderAsBearerToken(),

      issuer: COGNITO_URL,

      algorithms: ['RS256'],

      secretOrKeyProvider:
        jwksRsa.passportJwtSecret({
          cache: true,
          rateLimit: true,

          jwksRequestsPerMinute: 5,

          jwksUri: `${COGNITO_URL}/.well-known/jwks.json`,
        }) as any,
    });
  }

  async validate(payload: any) {
    return payload;
  }
}