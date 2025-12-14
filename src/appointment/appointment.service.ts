import { Injectable, NotFoundException, ForbiddenException } from '@nestjs/common';
import { ConfigService } from '@nestjs/config';
import { PrismaService } from '../prisma/prisma.service';
import { CreateAppointmentDto, UpdateAppointmentDto } from './dto';
import { GoHighLevelService } from '../integrations/gohighlevel.service';
import { UserService } from '../user/user.service';
import { Logger } from '@nestjs/common';

@Injectable()
export class AppointmentService {
  private readonly logger = new Logger(AppointmentService.name);

  constructor(
    private prisma: PrismaService,
    private goHighLevelService: GoHighLevelService,
    private userService: UserService,
    private readonly configService: ConfigService,
  ) {}

  async create(data: CreateAppointmentDto) {
    return this.prisma.appointment.create({ data });
  }

  async findAll() {
    return this.prisma.appointment.findMany();
  }

  async findAllByUser(userId: string) {
    return this.prisma.appointment.findMany({
      where: { userId }
    });
  }

  async findOne(id: string) {
    const appointment = await this.prisma.appointment.findUnique({ where: { id } });
    if (!appointment) throw new NotFoundException('Appointment not found');
    return appointment;
  }

  async findOneByUser(id: string, userId: string) {
    const appointment = await this.prisma.appointment.findFirst({
      where: { 
        id,
        userId 
      }
    });
    if (!appointment) throw new NotFoundException('Appointment not found');
    return appointment;
  }

  async update(id: string, data: UpdateAppointmentDto) {
    return this.prisma.appointment.update({ where: { id }, data });
  }

  async updateByUser(id: string, data: UpdateAppointmentDto, userId: string) {
    const appointment = await this.findOneByUser(id, userId);
    return this.prisma.appointment.update({ where: { id }, data });
  }

  async remove(id: string) {
    return this.prisma.appointment.delete({ where: { id } });
  }

  async removeByUser(id: string, userId: string) {
    const appointment = await this.findOneByUser(id, userId);
    return this.prisma.appointment.delete({ where: { id } });
  }

  async getGhlAppointmentsForUser(userId: string) {
    try {
      // Get user details to find their GHL user ID
      const user = await this.userService.findById(userId);
      if (!user) {
        throw new NotFoundException('User not found');
      }

      this.logger.log(`Fetching GHL appointments for user: ${user.name} (${user.role})`);

      // Get GHL credentials
      const credentials = this.getGhlCredentials();
      if (!credentials) {
        this.logger.warn('GHL credentials not configured - returning empty response');
        return { appointments: [], total: 0 };
      }

      // Validate and get GHL user ID
      let ghlUserId = await this.validateAndGetGhlUserId(user, credentials);
      if (!ghlUserId) {
        this.logger.warn(`Could not determine GHL user ID for ${user.name} - returning empty response`);
        return { appointments: [], total: 0 };
      }

      // Calculate date range (last 30 days to next 30 days)
      const now = new Date();
      const startDate = new Date(now.getTime() - (30 * 24 * 60 * 60 * 1000)); // 30 days ago
      const endDate = new Date(now.getTime() + (30 * 24 * 60 * 60 * 1000)); // 30 days from now

      // Get all appointments from GHL using the user's GHL user ID
      let allAppointments: any[] = [];
      
      if (ghlUserId) {
        // Try to get appointments using the user's GHL user ID
        try {
          allAppointments = await this.goHighLevelService.getAppointmentsByUserId(
            credentials.accessToken,
            credentials.locationId,
            ghlUserId
          );
          this.logger.log(`Fetched ${allAppointments.length} appointments using user's GHL user ID: ${ghlUserId}`);
        } catch (error) {
          this.logger.warn(`Failed to fetch appointments using user's GHL user ID: ${error.message}`);
          // Fallback to getting all appointments and filtering
          this.logger.warn(`⚠️ FALLBACK MODE: Fetching ALL appointments and filtering for user ${user.name} (${ghlUserId}) - this may cause appointment mixing issues!`);
          try {
            allAppointments = await this.goHighLevelService.getAppointments(
              credentials.accessToken,
              credentials.locationId
            );
            this.logger.log(`Fetched ${allAppointments.length} total appointments from GHL (fallback)`);
          } catch (fallbackError) {
            this.logger.error(`Failed to fetch appointments (fallback): ${fallbackError.message}`);
            allAppointments = [];
          }
        }
      } else {
        // User doesn't have a GHL user ID, try to get all appointments
        try {
          allAppointments = await this.goHighLevelService.getAppointments(
            credentials.accessToken,
            credentials.locationId
          );
          this.logger.log(`Fetched ${allAppointments.length} total appointments from GHL`);
        } catch (error) {
          this.logger.error(`Failed to fetch appointments: ${error.message}`);
          allAppointments = [];
        }
      }

      // Filter appointments based on user role and assignment with improved precision
      let filteredAppointments = allAppointments;

      // Special debugging for Koch users
      if (user.name && user.name.toLowerCase().includes('koch')) {
        this.logger.warn(`🔍 DEBUGGING KOCH USER: ${user.name} (${user.username}) - GHL ID: ${ghlUserId}, Team ID: ${user.ghlTeamId}`);
        this.logger.warn(`Total appointments before filtering: ${allAppointments.length}`);
      }

      if (user.role === 'SURVEYOR') {
        // For surveyors, filter appointments assigned to them or their team
        filteredAppointments = allAppointments.filter((appointment: any) => {
          const assignedUserId = appointment.userId || appointment.assignedTo;
          const assignedTeamId = appointment.teamId; // Only use teamId, not teamMemberId
          const contactAssignedTo = appointment.contact?.assignedTo;
          
          // Log appointment assignment details for debugging
          this.logger.debug(`Appointment ${appointment.id} assignment check:`, {
            assignedUserId,
            assignedTeamId,
            contactAssignedTo,
            userGhlId: ghlUserId,
            userTeamId: user.ghlTeamId
          });
          
          // Check if appointment is assigned to this user
          const isAssignedToUser = assignedUserId === ghlUserId;
          
          // Check if appointment is assigned to this user's team (only if both team IDs exist and are not null)
          const isAssignedToTeam = assignedTeamId && user.ghlTeamId && assignedTeamId === user.ghlTeamId;
          
          // Check if contact is assigned to this user
          const isContactAssignedToUser = contactAssignedTo === ghlUserId;
          
          // Prioritize appointment assignment over contact assignment to prevent cross-user visibility
          // If appointment is explicitly assigned to a user, only show it to that user
          // Only use contact assignment if appointment assignment is not specified
          let isAssigned;
          if (assignedUserId) {
            // Appointment has explicit assignment - only show to assigned user or team
            isAssigned = isAssignedToUser || isAssignedToTeam;
          } else {
            // No explicit appointment assignment - use contact assignment
            isAssigned = isContactAssignedToUser || isAssignedToTeam;
          }
          
          if (isAssigned) {
            this.logger.debug(`✅ Appointment ${appointment.id} assigned to user ${user.name} (${ghlUserId})`);
          } else {
            this.logger.debug(`❌ Appointment ${appointment.id} NOT assigned to user ${user.name} (${ghlUserId})`);
          }
          
          return isAssigned;
        });
        
        // Special debugging for Koch users - show final count
        if (user.name && user.name.toLowerCase().includes('koch')) {
          this.logger.warn(`🔍 KOCH USER FILTERING RESULT: ${user.name} - ${filteredAppointments.length} appointments after filtering`);
        }
      } else if (user.role === 'ADMIN') {
        // Admins can see all appointments
        filteredAppointments = allAppointments;
      } else {
        // For other roles, only show appointments assigned to them
        filteredAppointments = allAppointments.filter((appointment: any) => {
          const assignedUserId = appointment.userId || appointment.assignedTo;
          const contactAssignedTo = appointment.contact?.assignedTo;
          
          // Log appointment assignment details for debugging
          this.logger.debug(`Appointment ${appointment.id} assignment check:`, {
            assignedUserId,
            contactAssignedTo,
            userGhlId: ghlUserId
          });
          
          // Prioritize appointment assignment over contact assignment to prevent cross-user visibility
          let isAssigned;
          if (assignedUserId) {
            // Appointment has explicit assignment - only show to assigned user
            isAssigned = assignedUserId === ghlUserId;
          } else {
            // No explicit appointment assignment - use contact assignment
            isAssigned = contactAssignedTo === ghlUserId;
          }
          
          if (isAssigned) {
            this.logger.debug(`✅ Appointment ${appointment.id} assigned to user ${user.name} (${ghlUserId})`);
          } else {
            this.logger.debug(`❌ Appointment ${appointment.id} NOT assigned to user ${user.name} (${ghlUserId})`);
          }
          
          return isAssigned;
        });
      }

      this.logger.log(`Filtered to ${filteredAppointments.length} appointments for user ${user.name}`);

      // Transform appointments to match our expected format
      const transformedAppointments = filteredAppointments.map((appointment: any) => ({
        id: appointment.id,
        customerName: appointment.contact?.firstName + ' ' + appointment.contact?.lastName || 
                     appointment.contact?.email || 
                     appointment.title || 
                     'Unknown Customer',
        customerPhone: appointment.contact?.phone || '',
        customerEmail: appointment.contact?.email || '',
        address: appointment.location || appointment.address || '',
        scheduledAt: appointment.startTime,
        status: this.mapGhlStatusToAppStatus(appointment.status || appointment.appoinmentStatus),
        notes: appointment.notes || appointment.calendarNotes || '',
        sourceChannel: 'manual', // Default to manual since we don't have this info from GHL
        ghlAppointmentId: appointment.id,
        userId: user.id,
        // Include additional GHL data
        ghlData: {
          contact: appointment.contact,
          calendarId: appointment.calendarId,
          calendarServiceId: appointment.calendarServiceId,
          startTime: appointment.startTime,
          endTime: appointment.endTime,
          selectedTimezone: appointment.selectedTimezone,
          isRecurring: appointment.isRecurring,
          locationId: appointment.locationId
        }
      }));

      return {
        appointments: transformedAppointments,
        total: transformedAppointments.length,
        user: {
          id: user.id,
          name: user.name,
          role: user.role
        }
      };

    } catch (error) {
      this.logger.error(`Error fetching GHL appointments for user ${userId}: ${error.message}`);
      throw error;
    }
  }

  private getGhlCredentials() {
    const accessToken = this.configService.get<string>('GOHIGHLEVEL_API_TOKEN');
    
    if (!accessToken) {
      this.logger.warn('GHL API token not configured - returning empty response');
      return null;
    }
    
    // Extract location ID from the JWT token (same as OpportunitiesService)
    try {
      const tokenData = this.extractTokenData(accessToken);
      const locationId = tokenData.locationId;
      
      if (!locationId) {
        this.logger.warn('Location ID not found in GHL token - returning empty response');
        return null;
      }
      
      this.logger.log(`Using GHL credentials - Location ID: ${locationId}`);
      return { accessToken, locationId };
    } catch (error) {
      this.logger.error('Error extracting location ID from GHL token:', error);
      return null;
    }
  }

  private extractTokenData(accessToken: string) {
    try {
      const payload = JSON.parse(Buffer.from(accessToken.split('.')[1], 'base64').toString());
      return {
        locationId: payload.location_id || payload.locationId,
        userId: payload.sub || payload.userId,
        companyId: payload.companyId
      };
    } catch (error) {
      this.logger.error('Error extracting token data:', error);
      throw new Error('Invalid access token format');
    }
  }

  private mapGhlStatusToAppStatus(ghlStatus: string): 'scheduled' | 'completed' | 'cancelled' {
    const status = ghlStatus?.toLowerCase() || '';
    
    if (status.includes('completed') || status.includes('done')) {
      return 'completed';
    } else if (status.includes('cancelled') || status.includes('cancelled')) {
      return 'cancelled';
    } else {
      return 'scheduled';
    }
  }

  /**
   * Validate and get GHL user ID for a user with improved error handling
   */
  private async validateAndGetGhlUserId(user: any, credentials: { accessToken: string; locationId: string }): Promise<string | null> {
    try {
      // If user already has a GHL user ID, validate it exists
      if (user.ghlUserId) {
        this.logger.log(`User ${user.name} has GHL user ID: ${user.ghlUserId}`);
        
        // Verify the GHL user ID is valid by checking if the user exists in GHL
        try {
          const ghlUsers = await this.goHighLevelService.getAllUsers(credentials.accessToken, credentials.locationId);
          const ghlUser = ghlUsers.find((u: any) => u.id === user.ghlUserId);
          
          if (ghlUser) {
            this.logger.log(`✅ Verified GHL user ID for ${user.name}: ${ghlUser.firstName} ${ghlUser.lastName} (${user.ghlUserId})`);
            return user.ghlUserId;
          } else {
            this.logger.warn(`❌ GHL user ID ${user.ghlUserId} not found in GHL for user ${user.name}`);
          }
        } catch (error) {
          this.logger.error(`Error verifying GHL user ID for ${user.name}: ${error.message}`);
        }
      }

      // Try to find GHL user by name
      if (user.name) {
        this.logger.log(`User ${user.name} doesn't have a valid GHL user ID, trying to find it...`);
        try {
          const ghlUser = await this.goHighLevelService.findUserByName(
            credentials.accessToken,
            credentials.locationId,
            user.name
          );
          if (ghlUser) {
            this.logger.log(`Found GHL user ID for ${user.name}: ${ghlUser.id} (${ghlUser.firstName} ${ghlUser.lastName})`);
            
            // Update the user's GHL user ID in the database
            await this.userService.update(user.id, { ghlUserId: ghlUser.id });
            this.logger.log(`Updated user ${user.name} with GHL user ID: ${ghlUser.id}`);
            return ghlUser.id;
          } else {
            this.logger.warn(`Could not find GHL user for: ${user.name}`);
          }
        } catch (error) {
          this.logger.error(`Error finding GHL user for ${user.name}: ${error.message}`);
        }
      }

      this.logger.error(`Could not determine GHL user ID for user: ${user.name} (${user.username})`);
      return null;
    } catch (error) {
      this.logger.error(`Error validating GHL user ID for ${user.name}: ${error.message}`);
      return null;
    }
  }
} 