import { Injectable } from '@nestjs/common';
import { exec } from 'child_process';
import { promisify } from 'util';
import { writeFileSync, unlinkSync } from 'fs';
import { join } from 'path';
import { BookAppointmentDto } from './dto/book-appointment.dto';
import { UserService } from '../user/user.service';

const execAsync = promisify(exec);

@Injectable()
export class CalendarService {
  constructor(private readonly userService: UserService) {}
  // Calendar mapping
  private readonly CALENDAR_MAPPING = {
    'Phil': 'Philip  Edmondson',
    'Darren': 'Darren  Powell', 
    'Nick': 'Nicholas  Goldson',
    'Owen': 'Owen Shannon',
    'Richard': 'Richard Orchard'
  };

  // User to Calendar mapping (updated based on requirements)
  private readonly USER_CALENDAR_MAPPING = {
    'Andrew': ['Phil', 'Darren', 'Nick'],
    'Ion': ['Phil', 'Darren', 'Nick'],
    'Jordan': ['Phil', 'Darren'],
    'Onur': ['Phil', 'Darren'],
    'Kanji': ['Phil'], // Updated: only Phil
    'Kenji': ['Darren', 'Nick'], // Updated: Darren and Nick
    'Alex': ['Phil', 'Darren', 'Nick'],
    'James': ['Owen', 'Richard'] // Updated: James now has Owen and Richard instead of Jon
  };

  async getCurrentUserCalendarEvents(startDate?: string, endDate?: string, user?: any) {
    console.log('🔍 CalendarService: getCurrentUserCalendarEvents called with:', { startDate, endDate, user: user?.username || user?.name });
    
    // Get the full user details from the database
    const fullUser = await this.userService.findById(user?.sub);
    const authenticatedUserName = fullUser?.name || fullUser?.username || user?.username || user?.name || 'Unknown User';
    console.log('🔍 Using authenticated user name:', authenticatedUserName, 'from user:', fullUser?.name);
    
    // Map user name to calendar name - need to add Miles Kent mapping
    const userToCalendarMapping = {
      'Miles Kent': 'Miles Kent', // Add this mapping
      'Phil': 'Philip  Edmondson',
      'Darren': 'Darren  Powell', 
      'Nick': 'Nicholas  Goldson',
      'Owen': 'Owen Shannon',
      'Richard': 'Richard Orchard'
    };
    
    const outlookCalendarName = userToCalendarMapping[authenticatedUserName] || authenticatedUserName;
    console.log('🔍 Mapped calendar name:', outlookCalendarName, 'for user:', authenticatedUserName);
    
    // Create PowerShell script to get current user's calendar events using proper calendar path
    const script = `
$ErrorActionPreference = "Stop"
$outlook = $null
$namespace = $null
$calendar = $null

try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    
    # Use the proper calendar path structure like the working endpoints
    $rootFolder = $namespace.Folders.Item("calendars@creativuk.co.uk")
    $calendarFolder = $rootFolder.Folders.Item("Calendar")
    $calendar = $calendarFolder.Folders.Item("${outlookCalendarName}")
    
    # Use the authenticated user name from the JWT token
    $userName = "${authenticatedUserName}"
    $userEmail = $namespace.CurrentUser.Address
    
    # Get date range
    $startDate = "${startDate || '2025-09-01'}"
    $endDate = "${endDate || '2025-12-31'}"
    
    $startDateTime = [DateTime]::ParseExact($startDate, "yyyy-MM-dd", $null)
    $endDateTime = [DateTime]::ParseExact($endDate, "yyyy-MM-dd", $null).AddDays(1)
    
    # Get all items and filter for date range
    $allItems = $calendar.Items
    $events = @()
    
    Write-Host "🔍 Debug: Calendar path: calendars@creativuk.co.uk/Calendar/${outlookCalendarName}"
    Write-Host "🔍 Debug: Date range: $startDate to $endDate"
    Write-Host "🔍 Debug: Total items in calendar: $($allItems.Count)"
    
    foreach ($item in $allItems) {
        if ($item.Class -eq 26) {  # 26 = olAppointmentItem
            $itemStart = $item.Start
            Write-Host "🔍 Debug: Found event: $($item.Subject) on $($itemStart.ToString('yyyy-MM-dd'))"
            
            if ($itemStart -ge $startDateTime -and $itemStart -lt $endDateTime) {
                $event = @{
                    id = $item.EntryID
                    title = $item.Subject
                    startTime = $item.Start.ToString('HH:mm')
                    endTime = $item.End.ToString('HH:mm')
                    date = $item.Start.ToString('yyyy-MM-dd')
                    location = if ($item.Location) { $item.Location } else { "" }
                    status = switch ($item.BusyStatus) {
                        0 { "free" }
                        1 { "tentative" }
                        2 { "busy" }
                        3 { "out-of-office" }
                        default { "unknown" }
                    }
                    isAllDay = $item.AllDayEvent
                    isRecurring = $item.IsRecurring
                }
                $events += $event
                Write-Host "🔍 Debug: Added event to results: $($item.Subject)"
            }
        }
    }
    
    Write-Host "🔍 Debug: Total events found: $($events.Count)"
    
    $result = @{
        calendarName = "Current User Calendar"
        userName = $userName
        userEmail = $userEmail
        events = $events
        totalEvents = $events.Count
        dateRange = @{
            start = $startDate
            end = $endDate
        }
        debug = @{
            authenticatedUserName = "${authenticatedUserName}"
            currentUserAddress = $userEmail
            outlookCalendarName = "${outlookCalendarName}"
            calendarPath = "calendars@creativuk.co.uk/Calendar/${outlookCalendarName}"
        }
    }
    
    $resultJson = $result | ConvertTo-Json -Depth 3
    Write-Output $resultJson
    
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
    Write-Host "🔍 Debug: Error details: $($_.Exception.Message)"
    Write-Host "🔍 Debug: Stack trace: $($_.ScriptStackTrace)"
    exit 1
} finally {
    if ($calendar) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($calendar) | Out-Null }
    if ($namespace) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($namespace) | Out-Null }
    if ($outlook) { 
        $outlook.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null 
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
`;

    // Write script to temporary file
    const tempScriptPath = join(process.cwd(), `temp-current-calendar-script-${Date.now()}.ps1`);
    writeFileSync(tempScriptPath, script);
    
    try {
      // Execute PowerShell script from file
      const { stdout, stderr } = await execAsync(`powershell -ExecutionPolicy Bypass -File "${tempScriptPath}"`, { timeout: 30000 });
      
      console.log('PowerShell stdout:', stdout);
      console.log('PowerShell stderr:', stderr);
      
      if (stdout.includes('ERROR:')) {
        throw new Error(stdout.replace('ERROR:', '').trim());
      }
      
      // Clean the output - remove any non-JSON content
      const cleanOutput = stdout.trim();
      const jsonStart = cleanOutput.indexOf('{');
      const jsonEnd = cleanOutput.lastIndexOf('}');
      
      if (jsonStart === -1 || jsonEnd === -1) {
        throw new Error(`No valid JSON found in PowerShell output: ${cleanOutput}`);
      }
      
      const jsonString = cleanOutput.substring(jsonStart, jsonEnd + 1);
      
      try {
        const result = JSON.parse(jsonString);
        console.log('🔍 CalendarService: Successfully retrieved current user calendar events:', result);
        return result;
      } catch (parseError) {
        console.error('JSON parse error:', parseError);
        console.error('Raw output:', cleanOutput);
        console.error('Extracted JSON:', jsonString);
        throw new Error(`Failed to parse JSON: ${parseError.message}`);
      }
    } finally {
      // Clean up temporary file
      try {
        unlinkSync(tempScriptPath);
      } catch (cleanupError) {
        console.warn('Failed to clean up temporary script file:', cleanupError);
      }
    }
  }

  async getCalendarEvents(calendarName: string, startDate?: string, endDate?: string) {
    console.log('🔍 CalendarService: getCalendarEvents called with:', { calendarName, startDate, endDate });
    
    if (!this.CALENDAR_MAPPING[calendarName]) {
      console.error('🔍 CalendarService: Invalid calendar name:', calendarName);
      throw new Error('Invalid calendar name');
    }

    const outlookCalendarName = this.CALENDAR_MAPPING[calendarName];
    console.log('🔍 CalendarService: Using Outlook calendar name:', outlookCalendarName);
    console.log('🔍 CalendarService: Multi-day event handling enabled for all calendars');
    
    // Create PowerShell script to get calendar events
    const script = `
$ErrorActionPreference = "Stop"
$outlook = $null
$namespace = $null
$calendar = $null

try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $rootFolder = $namespace.Folders.Item("calendars@creativuk.co.uk")
    $calendarFolder = $rootFolder.Folders.Item("Calendar")
    $calendar = $calendarFolder.Folders.Item("${outlookCalendarName}")
    
    # Get date range
    $startDate = "${startDate || '2025-09-01'}"
    $endDate = "${endDate || '2025-12-31'}"
    
    $startDateTime = [DateTime]::ParseExact($startDate, "yyyy-MM-dd", $null)
    $endDateTime = [DateTime]::ParseExact($endDate, "yyyy-MM-dd", $null).AddDays(1)
    
    # Get all items and filter for date range
    $allItems = $calendar.Items
    $events = @()
    
    foreach ($item in $allItems) {
        if ($item.Class -eq 26) {
            $itemStart = $item.Start
            $itemEnd = $item.End
            
            if ($itemStart -lt $endDateTime -and $itemEnd -gt $startDateTime) {
                $title = $item.Subject
                $busyStatus = $item.BusyStatus
                
                if ($title -match "Not available|Unavailable|Blocked|Holiday|Leave") {
                    $status = "busy"
                } else {
                    $status = switch ($busyStatus) {
                        0 { "free" }
                        1 { "tentative" }
                        2 { "busy" }
                        3 { "out-of-office" }
                        default { "unknown" }
                    }
                }
                
                $isMultiDay = $item.Start.Date -ne $item.End.Date
                $eventStartDate = $item.Start.ToString('yyyy-MM-dd')
                $eventEndDate = $item.End.ToString('yyyy-MM-dd')
                
                if ($isMultiDay) {
                    Write-Host "Processing multi-day event: $($item.Subject) from $eventStartDate to $eventEndDate"
                    $currentDate = $item.Start.Date
                    $endDateOnly = $item.End.Date
                    
                    while ($currentDate -le $endDateOnly) {
                        $event = @{
                            id = $item.EntryID + "_" + $currentDate.ToString('yyyy-MM-dd')
                            title = $item.Subject
                            startTime = if ($currentDate -eq $item.Start.Date) { $item.Start.ToString('HH:mm') } else { "00:00" }
                            endTime = if ($currentDate -eq $item.End.Date) { $item.End.ToString('HH:mm') } else { "23:59" }
                            date = $currentDate.ToString('yyyy-MM-dd')
                            location = if ($item.Location) { $item.Location } else { "" }
                            status = $status
                            isAllDay = $item.AllDayEvent
                            isRecurring = $item.IsRecurring
                            startDate = $eventStartDate
                            endDate = $eventEndDate
                            isMultiDay = $true
                        }
                        $events += $event
                        $currentDate = $currentDate.AddDays(1)
                    }
                } else {
                    $event = @{
                        id = $item.EntryID
                        title = $item.Subject
                        startTime = $item.Start.ToString('HH:mm')
                        endTime = $item.End.ToString('HH:mm')
                        date = $item.Start.ToString('yyyy-MM-dd')
                        location = if ($item.Location) { $item.Location } else { "" }
                        status = $status
                        isAllDay = $item.AllDayEvent
                        isRecurring = $item.IsRecurring
                        startDate = $eventStartDate
                        endDate = $eventEndDate
                        isMultiDay = $false
                    }
                    $events += $event
                }
            }
        }
    }
    
    $result = @{
        calendarName = "${calendarName}"
        displayName = "${outlookCalendarName}"
        events = $events
        totalEvents = $events.Count
    }
    
    $resultJson = $result | ConvertTo-Json -Depth 3
    Write-Output $resultJson

} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
    exit 1
} finally {
    if ($calendar) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($calendar) | Out-Null }
    if ($namespace) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($namespace) | Out-Null }
    if ($outlook) { 
        $outlook.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null 
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
`;

    // Write script to temporary file
    const tempScriptPath = join(process.cwd(), `temp-calendar-script-${Date.now()}.ps1`);
    writeFileSync(tempScriptPath, script);
    
    try {
      // Execute PowerShell script from file
      const { stdout, stderr } = await execAsync(`powershell -ExecutionPolicy Bypass -File "${tempScriptPath}"`, { timeout: 30000 });
      
      console.log('PowerShell stdout:', stdout);
      console.log('PowerShell stderr:', stderr);
      
      if (stdout.includes('ERROR:')) {
        throw new Error(stdout.replace('ERROR:', '').trim());
      }
      
      // Clean the output - remove any non-JSON content
      const cleanOutput = stdout.trim();
      const jsonStart = cleanOutput.indexOf('{');
      const jsonEnd = cleanOutput.lastIndexOf('}');
      
      if (jsonStart === -1 || jsonEnd === -1) {
        throw new Error(`No valid JSON found in PowerShell output: ${cleanOutput}`);
      }
      
      const jsonString = cleanOutput.substring(jsonStart, jsonEnd + 1);
      
      try {
        const result = JSON.parse(jsonString);
        console.log('🔍 CalendarService: Successfully parsed calendar events:', result);
        return result;
      } catch (parseError) {
        console.error('JSON parse error:', parseError);
        console.error('Raw output:', cleanOutput);
        console.error('Extracted JSON:', jsonString);
        throw new Error(`Failed to parse JSON: ${parseError.message}`);
      }
    } finally {
      // Clean up temporary file
      try {
        unlinkSync(tempScriptPath);
      } catch (cleanupError) {
        console.warn('Failed to clean up temporary script file:', cleanupError);
      }
    }
  }

  async getUserCalendars(username: string) {
    const userCalendars = this.USER_CALENDAR_MAPPING[username] || [];
    
    const calendarInfo = userCalendars.map(calendar => ({
      id: calendar,
      name: calendar,
      displayName: this.CALENDAR_MAPPING[calendar] || calendar
    }));
    
    return {
      username,
      calendars: calendarInfo
    };
  }

  async bookAppointment(bookAppointmentDto: BookAppointmentDto) {
    const { 
      opportunityId, 
      customerName, 
      customerAddress, 
      calendar, 
      date, 
      timeSlot,
      installer,
      surveyor,
      surveyorEmail 
    } = bookAppointmentDto;
    
    if (!calendar || !date || !timeSlot) {
      throw new Error('Missing required fields');
    }
    
    if (!this.CALENDAR_MAPPING[calendar]) {
      throw new Error('Invalid calendar name');
    }
    
    const outlookCalendarName = this.CALENDAR_MAPPING[calendar];
    
    // Parse date and time
    const appointmentDate = new Date(date);
    const [hours, minutes] = timeSlot.split(':');
    appointmentDate.setHours(parseInt(hours), parseInt(minutes), 0, 0);
    
    // For installers, create all-day booking from 9am to end of day, then all day next day
    const endTime = new Date(appointmentDate);
    endTime.setHours(23, 59, 59, 999); // End of day
    
    // Next day - all day
    const nextDay = new Date(appointmentDate);
    nextDay.setDate(nextDay.getDate() + 1);
    nextDay.setHours(0, 0, 0, 0); // Start of next day
    
    const nextDayEndTime = new Date(nextDay);
    nextDayEndTime.setHours(23, 59, 59, 999); // End of next day
    
    // Format dates for PowerShell
    const startDateStr = appointmentDate.toISOString().replace('T', ' ').substring(0, 19);
    const endDateStr = endTime.toISOString().replace('T', ' ').substring(0, 19);
    const nextDayStartStr = nextDay.toISOString().replace('T', ' ').substring(0, 19);
    const nextDayEndStr = nextDayEndTime.toISOString().replace('T', ' ').substring(0, 19);
    
    // Create PowerShell script to book appointment
    const script = `
$ErrorActionPreference = "Stop"
$outlook = $null
$namespace = $null
$calendar = $null
$appointment = $null

try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    
    # Use the main calendar instead of individual installer calendars
    # This avoids permission issues with individual installer folders
    $calendar = $namespace.GetDefaultFolder(9)  # 9 = olFolderCalendar
    
    # Create appointment for Day 1 (9am to end of day)
    $appointment = $calendar.Items.Add(1)  # 1 = olAppointmentItem
    $appointment.Subject = "Solar Installation - ${customerName || 'Customer'} (${installer || 'Installer'}) - Day 1 (9am to end of day)"
    $appointment.Start = "${startDateStr}"
    $appointment.End = "${endDateStr}"
    $appointment.Location = "${customerAddress || 'Customer Address'}"
    $appointment.Body = "Solar Installation Appointment - Day 1 (9am to end of day)
    
Customer: ${customerName || 'N/A'}
Address: ${customerAddress || 'N/A'}
Opportunity ID: ${opportunityId || 'N/A'}
Installer: ${installer || 'N/A'}${surveyor ? `
Surveyor: ${surveyor}` : ''}
Assigned Calendar: ${outlookCalendarName}
Date: ${date}
Time: ${timeSlot} to end of day
Duration: 2 days (Day 1 of 2)

This appointment was created automatically by the Solar App."
    $appointment.ReminderMinutesBeforeStart = 60
    $appointment.BusyStatus = 2  # 2 = olBusy
    
    # Create appointment for Day 2 (all day)
    $appointment2 = $calendar.Items.Add(1)  # 1 = olAppointmentItem
    $appointment2.Subject = "Solar Installation - ${customerName || 'Customer'} (${installer || 'Installer'}) - Day 2 (All Day)"
    $appointment2.Start = "${nextDayStartStr}"
    $appointment2.End = "${nextDayEndStr}"
    $appointment2.Location = "${customerAddress || 'Customer Address'}"
    $appointment2.Body = "Solar Installation Appointment - Day 2 (All Day)
    
Customer: ${customerName || 'N/A'}
Address: ${customerAddress || 'N/A'}
Opportunity ID: ${opportunityId || 'N/A'}
Installer: ${installer || 'N/A'}${surveyor ? `
Surveyor: ${surveyor}` : ''}
Assigned Calendar: ${outlookCalendarName}
Date: ${date} (Day 2)
Time: All day
Duration: 2 days (Day 2 of 2)

This appointment was created automatically by the Solar App."
    $appointment2.ReminderMinutesBeforeStart = 60
    $appointment2.BusyStatus = 2  # 2 = olBusy
    
    # Add surveyor as attendee if provided (before saving)
    ${surveyor && surveyorEmail ? `
    try {
        Write-Host "Attempting to add surveyor: ${surveyor} with email: ${surveyorEmail}"
        # Add surveyor as attendee using their email address
        $recipient = $appointment.Recipients.Add("${surveyorEmail}")
        $recipient.Type = 1  # 1 = olRequired
        $resolved = $recipient.Resolve()
        Write-Host "Recipient resolve result: $resolved"
        Write-Host "Recipients count: $($appointment.Recipients.Count)"
        Write-Host "Added surveyor ${surveyor} (${surveyorEmail}) as required attendee"
    } catch {
        Write-Host "Could not add surveyor as attendee: $($_.Exception.Message)"
        Write-Host "Error details: $($_.Exception.ToString())"
    }` : surveyor ? `
    try {
        Write-Host "Attempting to add surveyor: ${surveyor} (name only)"
        # Fallback: try with name only if no email provided
        $recipient = $appointment.Recipients.Add("${surveyor}")
        $recipient.Type = 1  # 1 = olRequired
        $resolved = $recipient.Resolve()
        Write-Host "Recipient resolve result: $resolved"
        Write-Host "Recipients count: $($appointment.Recipients.Count)"
        Write-Host "Added surveyor ${surveyor} (using name) as required attendee"
    } catch {
        Write-Host "Could not add surveyor as attendee: $($_.Exception.Message)"
        Write-Host "Error details: $($_.Exception.ToString())"
    }` : ''}
    
    # Debug: Check appointment state before saving
    Write-Host "Before saving - Recipients count: $($appointment.Recipients.Count)"
    Write-Host "Before saving - MeetingStatus: $($appointment.MeetingStatus)"
    
    # Save the appointment with attendees
    try {
        $appointment.Save()
        $appointment2.Save()
        Write-Host "After saving - Recipients count: $($appointment.Recipients.Count)"
        Write-Host "After saving - MeetingStatus: $($appointment.MeetingStatus)"
        Write-Host "Saved both appointments (Day 1 and Day 2) with $($appointment.Recipients.Count) attendees"
        
        # If we have attendees, convert to meeting and send invitations
        ${surveyor && (surveyorEmail || surveyor) ? `
        if ($appointment.Recipients.Count -gt 0) {
            # Make it a meeting request
            $appointment.MeetingStatus = 1  # olMeeting
            Write-Host "Set MeetingStatus to 1 (olMeeting)"
            $appointment.Send()
            Write-Host "After Send() - Recipients count: $($appointment.Recipients.Count)"
            Write-Host "After Send() - MeetingStatus: $($appointment.MeetingStatus)"
            Write-Host "Converted to meeting and sent invitations to attendees"
        } else {
            Write-Host "No recipients found, not sending meeting invitation"
        }` : ''}
    } catch {
        Write-Host "Error saving/sending appointment: $($_.Exception.Message)"
        Write-Host "Error details: $($_.Exception.ToString())"
    }
    
    $result = @{
        success = $true
        appointmentId = $appointment.EntryID
        subject = $appointment.Subject
        start = $appointment.Start
        end = $appointment.End
        location = $appointment.Location
    }
    
    $resultJson = $result | ConvertTo-Json -Depth 3
    Write-Output $resultJson
    
} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
    exit 1
} finally {
    if ($appointment) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($appointment) | Out-Null }
    if ($appointment2) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($appointment2) | Out-Null }
    if ($calendar) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($calendar) | Out-Null }
    if ($namespace) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($namespace) | Out-Null }
    if ($outlook) { 
        $outlook.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null 
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
`;

    // Write script to temporary file
    const tempScriptPath = join(process.cwd(), `temp-booking-script-${Date.now()}.ps1`);
    writeFileSync(tempScriptPath, script);
    
    try {
      // Execute PowerShell script from file
      const { stdout, stderr } = await execAsync(`powershell -ExecutionPolicy Bypass -File "${tempScriptPath}"`, { timeout: 30000 });
      
      console.log('PowerShell stdout:', stdout);
      console.log('PowerShell stderr:', stderr);
      
      if (stdout.includes('ERROR:')) {
        throw new Error(stdout.replace('ERROR:', '').trim());
      }
      
      // Clean the output - remove any non-JSON content
      const cleanOutput = stdout.trim();
      const jsonStart = cleanOutput.indexOf('{');
      const jsonEnd = cleanOutput.lastIndexOf('}');
      
      if (jsonStart === -1 || jsonEnd === -1) {
        throw new Error(`No valid JSON found in PowerShell output: ${cleanOutput}`);
      }
      
      const jsonString = cleanOutput.substring(jsonStart, jsonEnd + 1);
      
      try {
        const result = JSON.parse(jsonString);
        
        if (result.success) {
          return {
            success: true,
            message: 'Appointment booked successfully',
            appointment: result
          };
        } else {
          throw new Error('Failed to book appointment');
        }
      } catch (parseError) {
        console.error('JSON parse error:', parseError);
        console.error('Raw output:', cleanOutput);
        console.error('Extracted JSON:', jsonString);
        throw new Error(`Failed to parse JSON: ${parseError.message}`);
      }
    } finally {
      // Clean up temporary file
      try {
        unlinkSync(tempScriptPath);
      } catch (cleanupError) {
        console.warn('Failed to clean up temporary script file:', cleanupError);
      }
    }
  }

  async getAvailability(calendarName: string, date: string) {
    if (!this.CALENDAR_MAPPING[calendarName]) {
      throw new Error('Invalid calendar name');
    }
    
    const outlookCalendarName = this.CALENDAR_MAPPING[calendarName];
    
    // Create PowerShell script to check availability
    const script = `
$ErrorActionPreference = "Stop"
$outlook = $null
$namespace = $null
$calendar = $null

try {
    $outlook = New-Object -ComObject Outlook.Application
    $namespace = $outlook.GetNamespace("MAPI")
    $rootFolder = $namespace.Folders.Item("calendars@creativuk.co.uk")
    $calendarFolder = $rootFolder.Folders.Item("Calendar")
    $calendar = $calendarFolder.Folders.Item("${outlookCalendarName}")
    
    # Parse date
    $targetDate = [DateTime]::ParseExact("${date}", "yyyy-MM-dd", $null)
    $startOfDay = $targetDate.Date
    $endOfDay = $targetDate.Date.AddDays(1)
    
    # Get all items for the date
    $allItems = $calendar.Items
    $dayEvents = @()
    
    foreach ($item in $allItems) {
        if ($item.Class -eq 26) {  # 26 = olAppointmentItem
            $itemStart = $item.Start
            $itemEnd = $item.End
            
            # Check if event overlaps with the target day (handles multi-day events)
            # Event overlaps if: (itemStart < endOfDay) AND (itemEnd > startOfDay)
            if ($itemStart -lt $endOfDay -and $itemEnd -gt $startOfDay) {
                $dayEvents += $item
            }
        }
    }
    
    # Generate available time slots (8 AM to 5 PM, 1-hour slots, including 9am for installers)
    $availableSlots = @()
    $busySlots = @()
    
    for ($hour = 8; $hour -le 17; $hour++) {
        $slotStart = $targetDate.Date.AddHours($hour)
        $slotEnd = $slotStart.AddHours(1)
        $slotTime = $slotStart.ToString('HH:mm')
        
        # Check if this slot conflicts with any events
        $isBusy = $false
        foreach ($event in $dayEvents) {
            $eventStart = $event.Start
            $eventEnd = $event.End
            $eventTitle = $event.Subject
            $eventBusyStatus = $event.BusyStatus
            $isAllDay = $event.AllDayEvent
            
            # If it's an all-day event, block all slots
            if ($isAllDay) {
                $isBusy = $true
                break
            }
            
            # If title contains "Not available" or similar, block all slots
            if ($eventTitle -match "Not available|Unavailable|Blocked|Holiday|Leave") {
                $isBusy = $true
                break
            }
            
            # Check for time overlap with busy events
            if ($eventBusyStatus -eq 2 -or $eventBusyStatus -eq 3) { # busy or out-of-office
                if (($slotStart -lt $eventEnd) -and ($slotEnd -gt $eventStart)) {
                    $isBusy = $true
                    break
                }
            }
        }
        
        if ($isBusy) {
            $busySlots += $slotTime
        } else {
            $availableSlots += $slotTime
        }
    }
    
    $result = @{
        date = "${date}"
        calendarName = "${calendarName}"
        displayName = "${outlookCalendarName}"
        availableSlots = $availableSlots
        busySlots = $busySlots
        totalEvents = $dayEvents.Count
        events = $dayEvents | ForEach-Object { 
            $title = $_.Subject
            $busyStatus = $_.BusyStatus
            
            # If title contains "Not available" or similar, mark as busy regardless of BusyStatus
            if ($title -match "Not available|Unavailable|Blocked|Holiday|Leave") {
                $status = "busy"
            } else {
                $status = switch ($busyStatus) {
                    0 { "free" }
                    1 { "tentative" }
                    2 { "busy" }
                    3 { "out-of-office" }
                    default { "unknown" }
                }
            }
            
            @{
                title = $title
                startTime = $_.Start.ToString('HH:mm')
                endTime = $_.End.ToString('HH:mm')
                location = if ($_.Location) { $_.Location } else { "" }
                status = $status
                isAllDay = $_.AllDayEvent
            }
        }
    }
    
    $resultJson = $result | ConvertTo-Json -Depth 3
    Write-Output $resultJson

} catch {
    Write-Output "ERROR: $($_.Exception.Message)"
    exit 1
} finally {
    if ($calendar) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($calendar) | Out-Null }
    if ($namespace) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($namespace) | Out-Null }
    if ($outlook) { 
        $outlook.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook) | Out-Null 
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
`;

    // Write script to temporary file
    const tempScriptPath = join(process.cwd(), `temp-availability-script-${Date.now()}.ps1`);
    writeFileSync(tempScriptPath, script);
    
    try {
      // Execute PowerShell script from file
      const { stdout, stderr } = await execAsync(`powershell -ExecutionPolicy Bypass -File "${tempScriptPath}"`, { timeout: 30000 });
      
      console.log('PowerShell stdout:', stdout);
      console.log('PowerShell stderr:', stderr);
      
      if (stdout.includes('ERROR:')) {
        throw new Error(stdout.replace('ERROR:', '').trim());
      }
      
      // Clean the output - remove any non-JSON content
      const cleanOutput = stdout.trim();
      const jsonStart = cleanOutput.indexOf('{');
      const jsonEnd = cleanOutput.lastIndexOf('}');
      
      if (jsonStart === -1 || jsonEnd === -1) {
        throw new Error(`No valid JSON found in PowerShell output: ${cleanOutput}`);
      }
      
      const jsonString = cleanOutput.substring(jsonStart, jsonEnd + 1);
      
      try {
        const result = JSON.parse(jsonString);
        console.log('🔍 CalendarService: Successfully parsed calendar events:', result);
        return result;
      } catch (parseError) {
        console.error('JSON parse error:', parseError);
        console.error('Raw output:', cleanOutput);
        console.error('Extracted JSON:', jsonString);
        throw new Error(`Failed to parse JSON: ${parseError.message}`);
      }
    } finally {
      // Clean up temporary file
      try {
        unlinkSync(tempScriptPath);
      } catch (cleanupError) {
        console.warn('Failed to clean up temporary script file:', cleanupError);
      }
    }
  }
}
