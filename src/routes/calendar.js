const express = require('express');
const { exec } = require('child_process');
const { promisify } = require('util');
const path = require('path');

const router = express.Router();
const execAsync = promisify(exec);

// Calendar mapping
const CALENDAR_MAPPING = {
  'Phil': 'Philip  Edmondson',
  'Darren': 'Darren  Powell', 
  'Nick': 'Nicholas  Goldson',
  'Owen': 'Owen Shannon',
  'Richard': 'Richard Orchard'
};

// User to Calendar mapping
const USER_CALENDAR_MAPPING = {
  'Andrew': ['Phil', 'Darren'],
  'Ion': ['Phil', 'Darren', 'Nick'],
  'Jordan': ['Phil', 'Darren'],
  'Onur': ['Phil', 'Darren'],
  'Kanji': ['Phil', 'Nick'],
  'Alex': ['Phil', 'Darren', 'Nick'],
  'James': ['Owen', 'Richard'],
  'Kenji': ['Nick']
};

// Get current logged-in user's calendar events
router.get('/current/events', async (req, res) => {
  try {
    const { startDate, endDate } = req.query;
    
    console.log('🔍 Getting current user calendar events for date range:', startDate, 'to', endDate);
    
    // For now, we'll use a default mapping - in production this should come from the JWT token
    // This is a temporary fix - the proper solution is to get the user from the JWT token
    const authenticatedUserName = 'Miles Kent'; // This should be extracted from JWT token
    
    // Map user name to calendar name
    const userToCalendarMapping = {
      'Miles Kent': 'Miles Kent',
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
    
    # Get current user info
    $currentUser = $namespace.CurrentUser
    $userName = $currentUser.Name
    $userEmail = $currentUser.Address
    
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
    const tempScriptPath = path.join(process.cwd(), `temp-current-calendar-script-${Date.now()}.ps1`);
    require('fs').writeFileSync(tempScriptPath, script);
    
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
        console.log('🔍 Successfully retrieved current user calendar events:', result);
        res.json(result);
      } catch (parseError) {
        console.error('JSON parse error:', parseError);
        console.error('Raw output:', cleanOutput);
        console.error('Extracted JSON:', jsonString);
        throw new Error(`Failed to parse JSON: ${parseError.message}`);
      }
    } finally {
      // Clean up temporary file
      try {
        require('fs').unlinkSync(tempScriptPath);
      } catch (cleanupError) {
        console.warn('Failed to clean up temporary script file:', cleanupError);
      }
    }
  } catch (error) {
    console.error('Error getting current user calendar events:', error);
    res.status(500).json({ error: error.message });
  }
});

// Get calendar events for a specific calendar (legacy endpoint)
router.get('/:calendarName/events', async (req, res) => {
  try {
    const { calendarName } = req.params;
    const { startDate, endDate } = req.query;
    
    if (!CALENDAR_MAPPING[calendarName]) {
      return res.status(400).json({ error: 'Invalid calendar name' });
    }

    const outlookCalendarName = CALENDAR_MAPPING[calendarName];
    
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
        if ($item.Class -eq 26) {  # 26 = olAppointmentItem
            $itemStart = $item.Start
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

    // Execute PowerShell script
    const { stdout } = await execAsync(`powershell -Command "${script}"`, { timeout: 30000 });
    
    if (stdout.includes('ERROR:')) {
      throw new Error(stdout.replace('ERROR:', '').trim());
    }
    
    const result = JSON.parse(stdout.trim());
    res.json(result);
    
  } catch (error) {
    console.error('Error getting calendar events:', error);
    res.status(500).json({ 
      error: 'Failed to get calendar events',
      details: error.message 
    });
  }
});

// Get available calendars for a user
router.get('/user/:username/calendars', async (req, res) => {
  try {
    const { username } = req.params;
    
    const userCalendars = USER_CALENDAR_MAPPING[username] || [];
    
    const calendarInfo = userCalendars.map(calendar => ({
      id: calendar,
      name: calendar,
      displayName: CALENDAR_MAPPING[calendar] || calendar
    }));
    
    res.json({
      username,
      calendars: calendarInfo
    });
    
  } catch (error) {
    console.error('Error getting user calendars:', error);
    res.status(500).json({ 
      error: 'Failed to get user calendars',
      details: error.message 
    });
  }
});

// Book an appointment
router.post('/book-appointment', async (req, res) => {
  try {
    const { 
      opportunityId, 
      customerName, 
      customerAddress, 
      calendar, 
      date, 
      timeSlot,
      installer 
    } = req.body;
    
    if (!calendar || !date || !timeSlot) {
      return res.status(400).json({ error: 'Missing required fields' });
    }
    
    if (!CALENDAR_MAPPING[calendar]) {
      return res.status(400).json({ error: 'Invalid calendar name' });
    }
    
    const outlookCalendarName = CALENDAR_MAPPING[calendar];
    
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
    $rootFolder = $namespace.Folders.Item("calendars@creativuk.co.uk")
    $calendarFolder = $rootFolder.Folders.Item("Calendar")
    $calendar = $calendarFolder.Folders.Item("${outlookCalendarName}")
    
    # Create appointment for Day 1 (9am to end of day)
    $appointment = $calendar.Items.Add(1)  # 1 = olAppointmentItem
    $appointment.Subject = "Solar Installation - ${customerName || 'Customer'} - Day 1 (9am to end of day)"
    $appointment.Start = "${appointmentDate.toISOString().replace('T', ' ').substring(0, 19)}"
    $appointment.End = "${endTime.toISOString().replace('T', ' ').substring(0, 19)}"
    $appointment.Location = "${customerAddress || 'Customer Address'}"
    $appointment.Body = "Solar Installation Appointment - Day 1 (9am to end of day)
    
Customer: ${customerName || 'N/A'}
Address: ${customerAddress || 'N/A'}
Opportunity ID: ${opportunityId || 'N/A'}
Installer: ${installer || 'N/A'}
Date: ${date}
Time: ${timeSlot} to end of day
Duration: 2 days (Day 1 of 2)

This appointment was created automatically by the Solar App."
    $appointment.ReminderMinutesBeforeStart = 60
    $appointment.BusyStatus = 2  # 2 = olBusy
    
    # Create appointment for Day 2 (all day)
    $appointment2 = $calendar.Items.Add(1)  # 1 = olAppointmentItem
    $appointment2.Subject = "Solar Installation - ${customerName || 'Customer'} - Day 2 (All Day)"
    $appointment2.Start = "${nextDay.toISOString().replace('T', ' ').substring(0, 19)}"
    $appointment2.End = "${nextDayEndTime.toISOString().replace('T', ' ').substring(0, 19)}"
    $appointment2.Location = "${customerAddress || 'Customer Address'}"
    $appointment2.Body = "Solar Installation Appointment - Day 2 (All Day)
    
Customer: ${customerName || 'N/A'}
Address: ${customerAddress || 'N/A'}
Opportunity ID: ${opportunityId || 'N/A'}
Installer: ${installer || 'N/A'}
Date: ${date} (Day 2)
Time: All day
Duration: 2 days (Day 2 of 2)

This appointment was created automatically by the Solar App."
    $appointment2.ReminderMinutesBeforeStart = 60
    $appointment2.BusyStatus = 2  # 2 = olBusy
    
    # Save both appointments
    $appointment.Save()
    $appointment2.Save()
    
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

    // Execute PowerShell script
    const { stdout } = await execAsync(`powershell -Command "${script}"`, { timeout: 30000 });
    
    if (stdout.includes('ERROR:')) {
      throw new Error(stdout.replace('ERROR:', '').trim());
    }
    
    const result = JSON.parse(stdout.trim());
    
    if (result.success) {
      res.json({
        success: true,
        message: 'Appointment booked successfully',
        appointment: result
      });
    } else {
      throw new Error('Failed to book appointment');
    }
    
  } catch (error) {
    console.error('Error booking appointment:', error);
    res.status(500).json({ 
      error: 'Failed to book appointment',
      details: error.message 
    });
  }
});

// Get availability for a specific date
router.get('/:calendarName/availability/:date', async (req, res) => {
  try {
    const { calendarName, date } = req.params;
    
    if (!CALENDAR_MAPPING[calendarName]) {
      return res.status(400).json({ error: 'Invalid calendar name' });
    }
    
    const outlookCalendarName = CALENDAR_MAPPING[calendarName];
    
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
            if ($itemStart -ge $startOfDay -and $itemStart -lt $endOfDay) {
                $dayEvents += $item
            }
        }
    }
    
    # Generate available time slots (8 AM to 6 PM, 1-hour slots)
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
            
            # Check for overlap
            if (($slotStart -lt $eventEnd) -and ($slotEnd -gt $eventStart)) {
                $isBusy = $true
                break
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
        events = $dayEvents | ForEach-Object { @{
            title = $_.Subject
            startTime = $_.Start.ToString('HH:mm')
            endTime = $_.End.ToString('HH:mm')
            location = if ($_.Location) { $_.Location } else { "" }
            status = switch ($_.BusyStatus) {
                0 { "free" }
                1 { "tentative" }
                2 { "busy" }
                3 { "out-of-office" }
                default { "unknown" }
            }
        }}
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

    // Execute PowerShell script
    const { stdout } = await execAsync(`powershell -Command "${script}"`, { timeout: 30000 });
    
    if (stdout.includes('ERROR:')) {
      throw new Error(stdout.replace('ERROR:', '').trim());
    }
    
    const result = JSON.parse(stdout.trim());
    res.json(result);
    
  } catch (error) {
    console.error('Error getting availability:', error);
    res.status(500).json({ 
      error: 'Failed to get availability',
      details: error.message 
    });
  }
});

module.exports = router;
