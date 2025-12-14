import { Injectable, UnauthorizedException } from '@nestjs/common';
import { PassportStrategy } from '@nestjs/passport';
import { ExtractJwt, Strategy } from 'passport-jwt';
import { ConfigService } from '@nestjs/config';
import { PrismaService } from '../prisma/prisma.service';

@Injectable()
export class JwtStrategy extends PassportStrategy(Strategy) {
  constructor(
    private readonly configService: ConfigService,
    private readonly prisma: PrismaService,
  ) {
    super({
      jwtFromRequest: ExtractJwt.fromAuthHeaderAsBearerToken(),
      ignoreExpiration: false,
      secretOrKey: configService.get<string>('JWT_SECRET'),
    });
  }

  async validate(payload: any) {
    let user = await this.prisma.user.findUnique({
      where: { id: payload.sub }
    });

    // If user not found by ID, try to find by GHL user ID
    if (!user && payload.ghlUserId) {
      user = await this.prisma.user.findUnique({
        where: { ghlUserId: payload.ghlUserId }
      });
      
      if (user) {
        console.log(`Found user by GHL user ID: ${user.username || user.name} (${user.id})`);
      }
    }

    if (!user || user.status !== 'ACTIVE') {
      throw new UnauthorizedException('User not found or inactive');
    }

    return {
      sub: user.id,
      username: user.username,
      email: user.email,
      role: user.role,
      ghlUserId: user.ghlUserId, // Include the ghlUserId in the token
    };
  }
} 