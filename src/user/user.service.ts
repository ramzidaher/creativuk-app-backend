import { Injectable } from '@nestjs/common';
import { PrismaService } from '../prisma/prisma.service';
import { User } from '@prisma/client';

@Injectable()
export class UserService {
  constructor(private prisma: PrismaService) {}

  async findById(id: string): Promise<User | null> {
    return this.prisma.user.findUnique({ where: { id } });
  }

  async findByGhlUserId(ghlUserId: string): Promise<User | null> {
    return this.prisma.user.findUnique({ where: { ghlUserId: ghlUserId } });
  }

  async upsertByGhlUserId(data: {
    ghlUserId?: string | null;
    name?: string | null;
    email?: string | null;
    ghlAccessToken?: string | null;
    ghlRefreshToken?: string | null;
    tokenExpiresAt?: Date | null;
  }): Promise<User> {
    if (!data.ghlUserId) {
      throw new Error('ghlUserId is required for upsert operation');
    }
    
    return this.prisma.user.upsert({
      where: { ghlUserId: data.ghlUserId },
      update: {
        name: data.name ?? undefined,
        email: data.email ?? `user-${data.ghlUserId}@temp.com`, // Provide default email
        ghlAccessToken: data.ghlAccessToken ?? '',
        ghlRefreshToken: data.ghlRefreshToken ?? '',
        tokenExpiresAt: data.tokenExpiresAt ?? undefined,
      },
      create: {
        ghlUserId: data.ghlUserId,
        name: data.name ?? undefined,
        email: data.email ?? `user-${data.ghlUserId}@temp.com`, // Provide default email
        username: `user-${data.ghlUserId}`, // Provide default username
        password: 'temp-password-hash', // Temporary password hash
        ghlAccessToken: data.ghlAccessToken ?? '',
        ghlRefreshToken: data.ghlRefreshToken ?? '',
        tokenExpiresAt: data.tokenExpiresAt ?? undefined,
      },
    });
  }

  async update(id: string, data: Partial<User>): Promise<User> {
    return this.prisma.user.update({
      where: { id },
      data,
    });
  }

  async findByRole(role: string): Promise<User[]> {
    return this.prisma.user.findMany({ where: { role: role as any } });
  }

  async findAll(): Promise<User[]> {
    return this.prisma.user.findMany();
  }
}
