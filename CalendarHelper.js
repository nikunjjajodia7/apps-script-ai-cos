/**
 * Calendar Helper Utilities
 * Functions for scheduling meetings and managing calendar events
 */

/**
 * Find available time slot for multiple attendees
 */
function findAvailableSlot(attendeeEmails, durationMinutes, startDate, endDate) {
  try {
    const calendar = CalendarApp.getDefaultCalendar();
    const duration = durationMinutes || CONFIG.DEFAULT_MEETING_DURATION_MINUTES();
    
    // Get free/busy information
    const freebusy = CalendarApp.getDefaultCalendar().getEvents(startDate, endDate);
    
    // Simple algorithm: find first available slot
    // For production, use Calendar API freebusy query for better accuracy
    let currentTime = new Date(startDate);
    const endTime = new Date(endDate);
    
    while (currentTime < endTime) {
      const slotEnd = new Date(currentTime.getTime() + duration * 60 * 1000);
      
      // Check if this slot is free for all attendees
      let isAvailable = true;
      
      // Check default calendar
      const conflictingEvents = calendar.getEvents(currentTime, slotEnd);
      if (conflictingEvents.length > 0) {
        isAvailable = false;
      }
      
      // Check other attendees' calendars (if we have access)
      for (let i = 0; i < attendeeEmails.length && isAvailable; i++) {
        try {
          const attendeeCalendar = CalendarApp.getCalendarById(attendeeEmails[i]);
          if (attendeeCalendar) {
            const attendeeEvents = attendeeCalendar.getEvents(currentTime, slotEnd);
            if (attendeeEvents.length > 0) {
              isAvailable = false;
            }
          }
        } catch (e) {
          // Don't have access to this calendar, skip
          Logger.log(`Cannot access calendar for ${attendeeEmails[i]}`);
        }
      }
      
      if (isAvailable) {
        return {
          start: currentTime,
          end: slotEnd,
        };
      }
      
      // Move to next 30-minute slot
      currentTime = new Date(currentTime.getTime() + 30 * 60 * 1000);
    }
    
    return null; // No available slot found
    
  } catch (error) {
    logError(ERROR_TYPE.API_ERROR, 'findAvailableSlot', error.toString(), null, error.stack);
    return null;
  }
}

/**
 * Schedule a one-on-one meeting
 */
function scheduleOneOnOne(taskId) {
  try {
    const task = getTask(taskId);
    if (!task || !task.Assignee_Email) {
      throw new Error(`Task ${taskId} has no assignee`);
    }
    
    const bossEmail = CONFIG.BOSS_EMAIL();
    const assigneeEmail = task.Assignee_Email;
    const duration = CONFIG.DEFAULT_MEETING_DURATION_MINUTES();
    
    // Find available slot in next 2 weeks
    const startDate = new Date();
    const endDate = new Date(startDate.getTime() + 14 * 24 * 60 * 60 * 1000);
    
    const slot = findAvailableSlot([bossEmail, assigneeEmail], duration, startDate, endDate);
    
    if (!slot) {
      // No slot found, set status to conflict
      updateTask(taskId, {
        Status: TASK_STATUS.SCHEDULING_CONFLICT,
      });
      logInteraction(taskId, 'Could not find available slot for 1-on-1 meeting');
      return null;
    }
    
    // Create calendar event
    const calendar = CalendarApp.getDefaultCalendar();
    const eventTitle = `1-on-1: ${task.Task_Name}`;
    const eventDescription = `Task: ${task.Task_Name}\n\nContext: ${task.Context_Hidden || task.Interaction_Log || 'No additional context'}`;
    
    const event = calendar.createEvent(
      eventTitle,
      slot.start,
      slot.end,
      {
        description: eventDescription,
        guests: assigneeEmail,
        sendInvites: true,
      }
    );
    
    // Update task
    updateTask(taskId, {
      Status: TASK_STATUS.SCHEDULED,
      Meeting_Action: '', // Clear the action
    });
    
    logInteraction(taskId, `1-on-1 meeting scheduled: ${Utilities.formatDate(slot.start, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')}`);
    
    return event.getId();
    
  } catch (error) {
    logError(ERROR_TYPE.API_ERROR, 'scheduleOneOnOne', error.toString(), taskId, error.stack);
    updateTask(taskId, {
      Status: TASK_STATUS.SCHEDULING_CONFLICT,
    });
    return null;
  }
}

/**
 * Add task to weekly agenda
 */
function addToWeeklyAgenda(taskId) {
  try {
    const task = getTask(taskId);
    if (!task) {
      throw new Error(`Task ${taskId} not found`);
    }
    
    const weeklyMeetingTitle = CONFIG.WEEKLY_MEETING_TITLE();
    const calendar = CalendarApp.getDefaultCalendar();
    
    // Find next weekly meeting (next 30 days)
    const startDate = new Date();
    const endDate = new Date(startDate.getTime() + 30 * 24 * 60 * 60 * 1000);
    const events = calendar.getEvents(startDate, endDate);
    
    let weeklyEvent = null;
    for (let i = 0; i < events.length; i++) {
      if (events[i].getTitle().includes(weeklyMeetingTitle)) {
        weeklyEvent = events[i];
        break;
      }
    }
    
    if (!weeklyEvent) {
      // No weekly meeting found, create one for next week
      const nextWeek = new Date(startDate.getTime() + 7 * 24 * 60 * 60 * 1000);
      // Find next Monday at 10am
      const daysUntilMonday = (8 - nextWeek.getDay()) % 7 || 7;
      const meetingDate = new Date(nextWeek);
      meetingDate.setDate(nextWeek.getDate() + daysUntilMonday);
      meetingDate.setHours(10, 0, 0, 0);
      
      weeklyEvent = calendar.createEvent(
        weeklyMeetingTitle,
        meetingDate,
        new Date(meetingDate.getTime() + 60 * 60 * 1000), // 1 hour
        {
          recurrence: CalendarApp.newRecurrence().addWeeklyRule().until(endDate),
        }
      );
    }
    
    // Append task to event description
    const currentDescription = weeklyEvent.getDescription() || '';
    const newDescription = currentDescription + 
      `\n\n---\nTask: ${task.Task_Name}\nContext: ${task.Context_Hidden || ''}`;
    
    weeklyEvent.setDescription(newDescription);
    
    // Update task
    updateTask(taskId, {
      Status: TASK_STATUS.SCHEDULED,
      Meeting_Action: '', // Clear the action
    });
    
    logInteraction(taskId, `Added to weekly agenda: ${weeklyEvent.getTitle()}`);
    
    return weeklyEvent.getId();
    
  } catch (error) {
    logError(ERROR_TYPE.API_ERROR, 'addToWeeklyAgenda', error.toString(), taskId, error.stack);
    return null;
  }
}

/**
 * Schedule focus time for Boss
 */
function scheduleFocusTime(taskId) {
  try {
    const task = getTask(taskId);
    if (!task) {
      throw new Error(`Task ${taskId} not found`);
    }
    
    const duration = CONFIG.FOCUS_TIME_DURATION_MINUTES();
    const calendar = CalendarApp.getDefaultCalendar();
    
    // Find available slot in next week
    const startDate = new Date();
    const endDate = new Date(startDate.getTime() + 7 * 24 * 60 * 60 * 1000);
    
    const slot = findAvailableSlot([CONFIG.BOSS_EMAIL()], duration, startDate, endDate);
    
    if (!slot) {
      updateTask(taskId, {
        Status: TASK_STATUS.SCHEDULING_CONFLICT,
      });
      logInteraction(taskId, 'Could not find available slot for focus time');
      return null;
    }
    
    // Create focus time event
    const eventTitle = `Deep Work: ${task.Task_Name}`;
    const eventDescription = `Focus time for task: ${task.Task_Name}\n\nContext: ${task.Context_Hidden || ''}`;
    
    const event = calendar.createEvent(
      eventTitle,
      slot.start,
      slot.end,
      {
        description: eventDescription,
      }
    );
    
    // Update task
    updateTask(taskId, {
      Status: TASK_STATUS.SCHEDULED,
      Meeting_Action: '', // Clear the action
    });
    
    logInteraction(taskId, `Focus time scheduled: ${Utilities.formatDate(slot.start, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')}`);
    
    return event.getId();
    
  } catch (error) {
    logError(ERROR_TYPE.API_ERROR, 'scheduleFocusTime', error.toString(), taskId, error.stack);
    updateTask(taskId, {
      Status: TASK_STATUS.SCHEDULING_CONFLICT,
    });
    return null;
  }
}

/**
 * Handle meeting action when set on a task
 */
function onMeetingActionSet(taskId) {
  try {
    const task = getTask(taskId);
    if (!task || !task.Meeting_Action) {
      return;
    }
    
    const action = task.Meeting_Action;
    
    if (action === MEETING_ACTION.ONE_ON_ONE) {
      scheduleOneOnOne(taskId);
    } else if (action === MEETING_ACTION.WEEKLY) {
      addToWeeklyAgenda(taskId);
    } else if (action === MEETING_ACTION.SELF) {
      scheduleFocusTime(taskId);
    }
    
  } catch (error) {
    logError(ERROR_TYPE.UNKNOWN_ERROR, 'onMeetingActionSet', error.toString(), taskId, error.stack);
  }
}

// ============================================
// TEST FUNCTIONS
// ============================================

/**
 * Test function: Schedule 1-on-1 meeting for a task
 * Usage: testScheduleOneOnOne('TASK-20251221031304')
 */
function testScheduleOneOnOne(taskId) {
  try {
    Logger.log(`=== Testing 1-on-1 Scheduling for Task: ${taskId} ===`);
    
    if (!taskId) {
      Logger.log('ERROR: No task ID provided');
      Logger.log('Usage: testScheduleOneOnOne("TASK-20251221031304")');
      Logger.log('Or run testScheduleOneOnOneAuto() to find a task automatically');
      return;
    }
    
    const task = getTask(taskId);
    if (!task) {
      Logger.log(`ERROR: Task ${taskId} not found`);
      return;
    }
    
    Logger.log(`Task: ${task.Task_Name}`);
    Logger.log(`Assignee: ${task.Assignee_Email || 'NOT SET'}`);
    Logger.log(`Status: ${task.Status}`);
    
    if (!task.Assignee_Email) {
      Logger.log('ERROR: Task has no assignee email');
      Logger.log('Please set Assignee_Email in the Tasks sheet');
      return;
    }
    
    Logger.log('\nScheduling 1-on-1 meeting...');
    const eventId = scheduleOneOnOne(taskId);
    
    if (eventId) {
      Logger.log(`✓ Meeting scheduled successfully!`);
      Logger.log(`Event ID: ${eventId}`);
      Logger.log('Check your Google Calendar for the new event');
    } else {
      Logger.log('✗ Could not schedule meeting (scheduling conflict)');
      Logger.log('Check task status - it should be set to "Scheduling_Conflict"');
    }
    
  } catch (error) {
    Logger.log(`ERROR: ${error.toString()}`);
    Logger.log(`Stack: ${error.stack || 'No stack trace'}`);
  }
}

/**
 * Test function: Automatically find a task and schedule 1-on-1
 */
function testScheduleOneOnOneAuto() {
  try {
    Logger.log('=== Finding Task for 1-on-1 Scheduling ===');
    
    // Find a task with an assignee
    const tasks = getSheetData(SHEETS.TASKS_DB);
    const taskWithAssignee = tasks.find(t => 
      t.Assignee_Email && 
      t.Assignee_Email.trim() !== '' && 
      t.Status !== TASK_STATUS.SCHEDULED &&
      t.Status !== TASK_STATUS.DONE
    );
    
    if (!taskWithAssignee) {
      Logger.log('ERROR: No suitable task found');
      Logger.log('Please create a task with Assignee_Email set');
      return;
    }
    
    Logger.log(`Found task: ${taskWithAssignee.Task_ID} - ${taskWithAssignee.Task_Name}`);
    testScheduleOneOnOne(taskWithAssignee.Task_ID);
    
  } catch (error) {
    Logger.log(`ERROR: ${error.toString()}`);
  }
}

/**
 * Test function: Add task to weekly agenda
 * Usage: testAddToWeeklyAgenda('TASK-20251221031304')
 */
function testAddToWeeklyAgenda(taskId) {
  try {
    Logger.log(`=== Testing Weekly Agenda for Task: ${taskId} ===`);
    
    if (!taskId) {
      Logger.log('ERROR: No task ID provided');
      Logger.log('Usage: testAddToWeeklyAgenda("TASK-20251221031304")');
      return;
    }
    
    const task = getTask(taskId);
    if (!task) {
      Logger.log(`ERROR: Task ${taskId} not found`);
      return;
    }
    
    Logger.log(`Task: ${task.Task_Name}`);
    Logger.log(`Status: ${task.Status}`);
    
    Logger.log('\nAdding to weekly agenda...');
    const eventId = addToWeeklyAgenda(taskId);
    
    if (eventId) {
      Logger.log(`✓ Task added to weekly agenda!`);
      Logger.log(`Event ID: ${eventId}`);
      Logger.log('Check your Google Calendar for the weekly meeting');
      Logger.log('The task has been added to the meeting description');
    } else {
      Logger.log('✗ Could not add to weekly agenda');
    }
    
  } catch (error) {
    Logger.log(`ERROR: ${error.toString()}`);
    Logger.log(`Stack: ${error.stack || 'No stack trace'}`);
  }
}

/**
 * Test function: Schedule focus time for a task
 * Usage: testScheduleFocusTime('TASK-20251221031304')
 */
function testScheduleFocusTime(taskId) {
  try {
    Logger.log(`=== Testing Focus Time Scheduling for Task: ${taskId} ===`);
    
    if (!taskId) {
      Logger.log('ERROR: No task ID provided');
      Logger.log('Usage: testScheduleFocusTime("TASK-20251221031304")');
      return;
    }
    
    const task = getTask(taskId);
    if (!task) {
      Logger.log(`ERROR: Task ${taskId} not found`);
      return;
    }
    
    Logger.log(`Task: ${task.Task_Name}`);
    Logger.log(`Status: ${task.Status}`);
    
    Logger.log('\nScheduling focus time...');
    const eventId = scheduleFocusTime(taskId);
    
    if (eventId) {
      Logger.log(`✓ Focus time scheduled successfully!`);
      Logger.log(`Event ID: ${eventId}`);
      Logger.log('Check your Google Calendar for the new focus time block');
    } else {
      Logger.log('✗ Could not schedule focus time (scheduling conflict)');
      Logger.log('Check task status - it should be set to "Scheduling_Conflict"');
    }
    
  } catch (error) {
    Logger.log(`ERROR: ${error.toString()}`);
    Logger.log(`Stack: ${error.stack || 'No stack trace'}`);
  }
}

/**
 * Test function: Find available time slots
 * Usage: testFindAvailableSlot(['email1@example.com', 'email2@example.com'], 30)
 */
function testFindAvailableSlot(attendeeEmails, durationMinutes) {
  try {
    Logger.log('=== Testing Available Slot Finder ===');
    
    if (!attendeeEmails || attendeeEmails.length === 0) {
      attendeeEmails = [CONFIG.BOSS_EMAIL()];
      Logger.log(`Using default: ${attendeeEmails[0]}`);
    }
    
    if (!durationMinutes) {
      durationMinutes = CONFIG.DEFAULT_MEETING_DURATION_MINUTES();
      Logger.log(`Using default duration: ${durationMinutes} minutes`);
    }
    
    Logger.log(`Attendees: ${attendeeEmails.join(', ')}`);
    Logger.log(`Duration: ${durationMinutes} minutes`);
    
    const startDate = new Date();
    const endDate = new Date(startDate.getTime() + 7 * 24 * 60 * 60 * 1000); // Next 7 days
    
    Logger.log(`Searching for slots between ${startDate} and ${endDate}...`);
    
    const slot = findAvailableSlot(attendeeEmails, durationMinutes, startDate, endDate);
    
    if (slot) {
      Logger.log(`✓ Available slot found!`);
      Logger.log(`Start: ${slot.start}`);
      Logger.log(`End: ${slot.end}`);
      Logger.log(`Duration: ${(slot.end - slot.start) / 1000 / 60} minutes`);
    } else {
      Logger.log('✗ No available slot found in the next 7 days');
    }
    
  } catch (error) {
    Logger.log(`ERROR: ${error.toString()}`);
    Logger.log(`Stack: ${error.stack || 'No stack trace'}`);
  }
}

/**
 * Test function: List upcoming calendar events
 */
function testListUpcomingEvents() {
  try {
    Logger.log('=== Upcoming Calendar Events ===');
    
    const calendar = CalendarApp.getDefaultCalendar();
    const startDate = new Date();
    const endDate = new Date(startDate.getTime() + 7 * 24 * 60 * 60 * 1000); // Next 7 days
    
    Logger.log(`Calendar: ${calendar.getName()}`);
    Logger.log(`Date range: ${startDate} to ${endDate}\n`);
    
    const events = calendar.getEvents(startDate, endDate);
    
    if (events.length === 0) {
      Logger.log('No events found in the next 7 days');
      return;
    }
    
    Logger.log(`Found ${events.length} event(s):\n`);
    
    events.forEach((event, index) => {
      Logger.log(`${index + 1}. ${event.getTitle()}`);
      Logger.log(`   Start: ${event.getStartTime()}`);
      Logger.log(`   End: ${event.getEndTime()}`);
      Logger.log(`   Guests: ${event.getGuestList().map(g => g.getEmail()).join(', ') || 'None'}`);
      Logger.log('');
    });
    
  } catch (error) {
    Logger.log(`ERROR: ${error.toString()}`);
    Logger.log(`Stack: ${error.stack || 'No stack trace'}`);
  }
}

/**
 * Test function: Check calendar access
 */
function testCalendarAccess() {
  try {
    Logger.log('=== Testing Calendar Access ===');
    
    const calendar = CalendarApp.getDefaultCalendar();
    Logger.log(`✓ Calendar access: OK`);
    Logger.log(`Calendar name: ${calendar.getName()}`);
    Logger.log(`Calendar ID: ${calendar.getId()}`);
    
    const bossEmail = CONFIG.BOSS_EMAIL();
    Logger.log(`Boss email: ${bossEmail}`);
    
    // Try to get events
    const startDate = new Date();
    const endDate = new Date(startDate.getTime() + 1 * 24 * 60 * 60 * 1000);
    const events = calendar.getEvents(startDate, endDate);
    Logger.log(`✓ Can read events: Found ${events.length} event(s) today`);
    
    Logger.log('\n=== Calendar Access Test Complete ===');
    
  } catch (error) {
    Logger.log(`ERROR: ${error.toString()}`);
    Logger.log(`Stack: ${error.stack || 'No stack trace'}`);
  }
}

