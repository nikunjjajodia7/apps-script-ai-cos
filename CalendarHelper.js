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
 * @param {string} taskId - Task ID
 * @param {string} preferredDate - Optional: Preferred date in YYYY-MM-DD format
 * @param {string} preferredTime - Optional: Preferred time in HH:MM format (24-hour)
 * @param {number} durationMinutes - Optional: Duration in minutes (default from config)
 */
function scheduleOneOnOne(taskId, preferredDate, preferredTime, durationMinutes) {
  try {
    const task = getTask(taskId);
    if (!task || !task.Assignee_Email) {
      throw new Error(`Task ${taskId} has no assignee`);
    }
    
    const bossEmail = CONFIG.BOSS_EMAIL();
    const assigneeEmail = task.Assignee_Email;
    const duration = durationMinutes || CONFIG.DEFAULT_MEETING_DURATION_MINUTES();
    
    let slot = null;
    
    // If date and time are provided, use them directly (check for conflicts)
    if (preferredDate && preferredTime) {
      try {
        const [year, month, day] = preferredDate.split('-').map(Number);
        const [hours, minutes] = preferredTime.split(':').map(Number);
        const eventStart = new Date(year, month - 1, day, hours, minutes, 0, 0);
        const eventEnd = new Date(eventStart.getTime() + duration * 60 * 1000);
        
        // Check if slot is available
        const calendar = CalendarApp.getDefaultCalendar();
        const existingEvents = calendar.getEvents(eventStart, eventEnd);
        
        // Check assignee calendar if accessible
        let hasConflict = existingEvents.length > 0;
        try {
          const assigneeCalendar = CalendarApp.getCalendarById(assigneeEmail);
          if (assigneeCalendar) {
            const assigneeEvents = assigneeCalendar.getEvents(eventStart, eventEnd);
            hasConflict = hasConflict || assigneeEvents.length > 0;
          }
        } catch (e) {
          Logger.log(`Cannot check assignee calendar: ${e.toString()}`);
        }
        
        if (!hasConflict) {
          slot = { start: eventStart, end: eventEnd };
        } else {
          Logger.log(`Preferred time slot has conflicts, will find alternative`);
        }
      } catch (e) {
        Logger.log(`Error parsing preferred date/time: ${e.toString()}`);
      }
    }
    
    // If no preferred slot or it had conflicts, find available slot
    if (!slot) {
      const startDate = preferredDate ? new Date(preferredDate + 'T00:00:00') : new Date();
      const endDate = new Date(startDate.getTime() + 14 * 24 * 60 * 60 * 1000);
      
      slot = findAvailableSlot([bossEmail, assigneeEmail], duration, startDate, endDate);
    }
    
    if (!slot) {
      // No slot found, set status to ai_assist for review
      updateTask(taskId, {
        Status: TASK_STATUS.AI_ASSIST,
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
    
    // Update task with calendar event info for bi-directional sync
    updateTask(taskId, {
      Status: TASK_STATUS.ON_TIME,  // Scheduled tasks are on_time
      Meeting_Action: '', // Clear the action
      Calendar_Event_ID: event.getId(),
      Scheduled_Time: slot.start,
      Previous_Status: task.Status, // Store previous status for reverting if event deleted
    });
    
    logInteraction(taskId, `1-on-1 meeting scheduled: ${Utilities.formatDate(slot.start, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')} (Event ID: ${event.getId()})`);
    
    return event.getId();
    
  } catch (error) {
    logError(ERROR_TYPE.API_ERROR, 'scheduleOneOnOne', error.toString(), taskId, error.stack);
    updateTask(taskId, {
      Status: TASK_STATUS.AI_ASSIST,  // Conflict needs review
    });
    return null;
  }
}

/**
 * Add task to weekly agenda
 * @param {string} taskId - Task ID
 * @param {string} preferredDate - Optional: Preferred date in YYYY-MM-DD format (used to find nearest weekly meeting)
 */
function addToWeeklyAgenda(taskId, preferredDate) {
  try {
    const task = getTask(taskId);
    if (!task) {
      throw new Error(`Task ${taskId} not found`);
    }
    
    const weeklyMeetingTitle = CONFIG.WEEKLY_MEETING_TITLE();
    const calendar = CalendarApp.getDefaultCalendar();
    
    // Find next weekly meeting (next 30 days)
    const startDate = preferredDate ? new Date(preferredDate + 'T00:00:00') : new Date();
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
      `\n\n---\nTask: ${task.Task_Name}\nTask ID: ${taskId}\nContext: ${task.Context_Hidden || ''}`;
    
    weeklyEvent.setDescription(newDescription);
    
    // Update task with calendar event info for bi-directional sync
    updateTask(taskId, {
      Status: TASK_STATUS.ON_TIME,  // Added to weekly is on_time
      Meeting_Action: '', // Clear the action
      Calendar_Event_ID: weeklyEvent.getId(),
      Scheduled_Time: weeklyEvent.getStartTime(),
      Previous_Status: task.Status, // Store previous status for reverting if event deleted
    });
    
    logInteraction(taskId, `Added to weekly agenda: ${weeklyEvent.getTitle()} (Event ID: ${weeklyEvent.getId()})`);
    
    return weeklyEvent.getId();
    
  } catch (error) {
    logError(ERROR_TYPE.API_ERROR, 'addToWeeklyAgenda', error.toString(), taskId, error.stack);
    return null;
  }
}

/**
 * Schedule focus time for Boss
 * @param {string} taskId - Task ID
 * @param {string} preferredDate - Optional: Preferred date in YYYY-MM-DD format
 * @param {string} preferredTime - Optional: Preferred time in HH:MM format (24-hour)
 * @param {number} durationMinutes - Optional: Duration in minutes (default from config)
 */
function scheduleFocusTime(taskId, preferredDate, preferredTime, durationMinutes) {
  try {
    const task = getTask(taskId);
    if (!task) {
      throw new Error(`Task ${taskId} not found`);
    }
    
    const duration = durationMinutes || CONFIG.FOCUS_TIME_DURATION_MINUTES();
    const calendar = CalendarApp.getDefaultCalendar();
    
    let slot = null;
    
    // If date and time are provided, use them directly (check for conflicts)
    if (preferredDate && preferredTime) {
      try {
        const [year, month, day] = preferredDate.split('-').map(Number);
        const [hours, minutes] = preferredTime.split(':').map(Number);
        const eventStart = new Date(year, month - 1, day, hours, minutes, 0, 0);
        const eventEnd = new Date(eventStart.getTime() + duration * 60 * 1000);
        
        // Check if slot is available
        const existingEvents = calendar.getEvents(eventStart, eventEnd);
        
        if (existingEvents.length === 0) {
          slot = { start: eventStart, end: eventEnd };
        } else {
          Logger.log(`Preferred time slot has conflicts, will find alternative`);
        }
      } catch (e) {
        Logger.log(`Error parsing preferred date/time: ${e.toString()}`);
      }
    }
    
    // If no preferred slot or it had conflicts, find available slot
    if (!slot) {
      const startDate = preferredDate ? new Date(preferredDate + 'T00:00:00') : new Date();
      const endDate = new Date(startDate.getTime() + 7 * 24 * 60 * 60 * 1000);
      
      slot = findAvailableSlot([CONFIG.BOSS_EMAIL()], duration, startDate, endDate);
    }
    
    if (!slot) {
      updateTask(taskId, {
        Status: TASK_STATUS.AI_ASSIST,  // Conflict needs review
      });
      logInteraction(taskId, 'Could not find available slot for focus time');
      return null;
    }
    
    // Create focus time event
    const eventTitle = `Deep Work: ${task.Task_Name}`;
    const eventDescription = `Focus time for task: ${task.Task_Name}\nTask ID: ${taskId}\n\nContext: ${task.Context_Hidden || ''}`;
    
    const event = calendar.createEvent(
      eventTitle,
      slot.start,
      slot.end,
      {
        description: eventDescription,
      }
    );
    
    // Update task with calendar event info for bi-directional sync
    updateTask(taskId, {
      Status: TASK_STATUS.ON_TIME,  // Focus time scheduled is on_time
      Meeting_Action: '', // Clear the action
      Calendar_Event_ID: event.getId(),
      Scheduled_Time: slot.start,
      Previous_Status: task.Status, // Store previous status for reverting if event deleted
    });
    
    logInteraction(taskId, `Focus time scheduled: ${Utilities.formatDate(slot.start, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')} (Event ID: ${event.getId()})`);
    
    return event.getId();
    
  } catch (error) {
    logError(ERROR_TYPE.API_ERROR, 'scheduleFocusTime', error.toString(), taskId, error.stack);
    updateTask(taskId, {
      Status: TASK_STATUS.AI_ASSIST,  // Conflict needs review
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
      t.Status !== TASK_STATUS.ON_TIME &&  // Already active/scheduled
      t.Status !== TASK_STATUS.CLOSED
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

// ============================================
// CALENDAR BI-DIRECTIONAL SYNC
// ============================================

/**
 * Sync calendar changes back to tasks
 * Detects when calendar events are deleted, rescheduled, or modified
 * Run this periodically (every 15-30 minutes) via time-driven trigger
 */
function syncCalendarChangesToTasks() {
  try {
    Logger.log('=== Syncing Calendar Changes to Tasks ===');
    
    const calendar = CalendarApp.getDefaultCalendar();
    const tz = Session.getScriptTimeZone();
    
    // Get all tasks with a Calendar_Event_ID
    const allTasks = getSheetData(SHEETS.TASKS_DB);
    const scheduledTasks = allTasks.filter(task => 
      task.Calendar_Event_ID && 
      task.Calendar_Event_ID.trim() !== '' &&
      task.Status === TASK_STATUS.ON_TIME  // Tasks with calendar events
    );
    
    Logger.log(`Found ${scheduledTasks.length} task(s) with calendar events to check`);
    
    let deletedCount = 0;
    let rescheduledCount = 0;
    let unchangedCount = 0;
    
    scheduledTasks.forEach(task => {
      try {
        const eventId = task.Calendar_Event_ID;
        const taskId = task.Task_ID;
        
        Logger.log(`\nChecking task ${taskId}: Event ID ${eventId}`);
        
        // Try to get the calendar event
        let event = null;
        try {
          event = calendar.getEventById(eventId);
        } catch (e) {
          Logger.log(`Could not retrieve event: ${e.toString()}`);
        }
        
        if (!event) {
          // Event was deleted in Google Calendar
          Logger.log(`Event DELETED for task ${taskId}`);
          handleCalendarEventDeleted(task);
          deletedCount++;
        } else {
          // Event exists - check if it was rescheduled
          const eventStartTime = event.getStartTime();
          const storedScheduledTime = task.Scheduled_Time ? new Date(task.Scheduled_Time) : null;
          
          // Check if event is cancelled
          if (event.getGuestList && event.getGuestList().length > 0) {
            const myStatus = event.getMyStatus();
            if (myStatus === CalendarApp.GuestStatus.NO) {
              Logger.log(`Event DECLINED for task ${taskId}`);
              handleCalendarEventDeleted(task);
              deletedCount++;
              return;
            }
          }
          
          // Compare times (allow 1 minute tolerance for timezone issues)
          if (storedScheduledTime) {
            const timeDiff = Math.abs(eventStartTime.getTime() - storedScheduledTime.getTime());
            const oneMinute = 60 * 1000;
            
            if (timeDiff > oneMinute) {
              // Event was rescheduled
              Logger.log(`Event RESCHEDULED for task ${taskId}`);
              Logger.log(`  Old time: ${storedScheduledTime}`);
              Logger.log(`  New time: ${eventStartTime}`);
              handleCalendarEventRescheduled(task, eventStartTime);
              rescheduledCount++;
            } else {
              // No changes
              unchangedCount++;
            }
          } else {
            // No stored time, update it now
            updateTask(taskId, { Scheduled_Time: eventStartTime });
            unchangedCount++;
          }
        }
        
      } catch (error) {
        Logger.log(`Error checking task ${task.Task_ID}: ${error.toString()}`);
      }
    });
    
    Logger.log(`\n=== Sync Complete ===`);
    Logger.log(`  Deleted: ${deletedCount}`);
    Logger.log(`  Rescheduled: ${rescheduledCount}`);
    Logger.log(`  Unchanged: ${unchangedCount}`);
    
    return {
      success: true,
      deleted: deletedCount,
      rescheduled: rescheduledCount,
      unchanged: unchangedCount
    };
    
  } catch (error) {
    Logger.log(`ERROR in syncCalendarChangesToTasks: ${error.toString()}`);
    Logger.log(`Stack: ${error.stack || 'No stack trace'}`);
    logError(ERROR_TYPE.API_ERROR, 'syncCalendarChangesToTasks', error.toString(), null, error.stack);
    return { success: false, error: error.toString() };
  }
}

/**
 * Handle when a calendar event is deleted
 * Reverts the task to its previous status
 */
function handleCalendarEventDeleted(task) {
  try {
    const taskId = task.Task_ID;
    const previousStatus = task.Previous_Status || TASK_STATUS.AI_ASSIST;
    
    Logger.log(`Reverting task ${taskId} from SCHEDULED to ${previousStatus}`);
    
    // Update task - clear calendar fields and revert status
    updateTask(taskId, {
      Status: previousStatus,
      Calendar_Event_ID: '',
      Scheduled_Time: '',
      Previous_Status: '',
    });
    
    logInteraction(taskId, `Calendar event was deleted. Task reverted to status: ${previousStatus}`);
    
    // Notify boss about the deletion
    notifyBossOfCalendarChange(task, 'deleted');
    
  } catch (error) {
    Logger.log(`Error handling deleted event for task ${task.Task_ID}: ${error.toString()}`);
  }
}

/**
 * Handle when a calendar event is rescheduled
 * Updates the task's scheduled time
 */
function handleCalendarEventRescheduled(task, newStartTime) {
  try {
    const taskId = task.Task_ID;
    const tz = Session.getScriptTimeZone();
    const oldTimeStr = task.Scheduled_Time 
      ? Utilities.formatDate(new Date(task.Scheduled_Time), tz, 'yyyy-MM-dd HH:mm')
      : 'Unknown';
    const newTimeStr = Utilities.formatDate(newStartTime, tz, 'yyyy-MM-dd HH:mm');
    
    Logger.log(`Updating task ${taskId} scheduled time: ${oldTimeStr} → ${newTimeStr}`);
    
    // Update task with new scheduled time
    updateTask(taskId, {
      Scheduled_Time: newStartTime,
    });
    
    logInteraction(taskId, `Calendar event rescheduled: ${oldTimeStr} → ${newTimeStr}`);
    
    // Notify boss about the reschedule
    notifyBossOfCalendarChange(task, 'rescheduled', oldTimeStr, newTimeStr);
    
  } catch (error) {
    Logger.log(`Error handling rescheduled event for task ${task.Task_ID}: ${error.toString()}`);
  }
}

/**
 * Notify boss about calendar changes
 */
function notifyBossOfCalendarChange(task, changeType, oldTime = null, newTime = null) {
  try {
    const bossEmail = CONFIG.BOSS_EMAIL();
    if (!bossEmail) return;
    
    let subject = '';
    let body = '';
    
    if (changeType === 'deleted') {
      subject = `📅 Calendar Event Deleted: ${task.Task_Name}`;
      body = `The calendar event for the following task was deleted:\n\n`;
      body += `Task: ${task.Task_Name}\n`;
      body += `Task ID: ${task.Task_ID}\n`;
      body += `Assignee: ${task.Assignee_Email || 'Self'}\n\n`;
      body += `The task status has been reverted to: ${task.Previous_Status || 'New'}\n\n`;
      body += `If this was intentional, no action is needed.\n`;
      body += `If you'd like to reschedule, please update the task in your dashboard.`;
    } else if (changeType === 'rescheduled') {
      subject = `📅 Meeting Rescheduled: ${task.Task_Name}`;
      body = `A calendar event for the following task was rescheduled:\n\n`;
      body += `Task: ${task.Task_Name}\n`;
      body += `Task ID: ${task.Task_ID}\n`;
      body += `Assignee: ${task.Assignee_Email || 'Self'}\n\n`;
      body += `Previous Time: ${oldTime}\n`;
      body += `New Time: ${newTime}\n\n`;
      body += `The task has been updated with the new scheduled time.`;
    }
    
    body += `\n\n---\nThis notification was sent by Chief of Staff AI`;
    
    GmailApp.sendEmail(
      bossEmail,
      subject,
      body,
      {
        name: 'Chief of Staff AI',
      }
    );
    
    Logger.log(`Boss notified of calendar ${changeType} for task ${task.Task_ID}`);
    
  } catch (error) {
    Logger.log(`Error notifying boss of calendar change: ${error.toString()}`);
  }
}

/**
 * Test function: Manually sync calendar changes
 */
function testSyncCalendarChanges() {
  Logger.log('=== Testing Calendar Sync ===');
  const result = syncCalendarChangesToTasks();
  Logger.log(`Result: ${JSON.stringify(result)}`);
}

/**
 * Test function: List tasks with calendar events
 */
function testListScheduledTasks() {
  try {
    Logger.log('=== Tasks with Calendar Events ===');
    
    const allTasks = getSheetData(SHEETS.TASKS_DB);
    const scheduledTasks = allTasks.filter(task => 
      task.Calendar_Event_ID && 
      task.Calendar_Event_ID.trim() !== ''
    );
    
    if (scheduledTasks.length === 0) {
      Logger.log('No tasks with calendar events found');
      return;
    }
    
    Logger.log(`Found ${scheduledTasks.length} task(s):\n`);
    
    scheduledTasks.forEach((task, index) => {
      Logger.log(`${index + 1}. ${task.Task_ID}: ${task.Task_Name}`);
      Logger.log(`   Status: ${task.Status}`);
      Logger.log(`   Event ID: ${task.Calendar_Event_ID}`);
      Logger.log(`   Scheduled: ${task.Scheduled_Time || 'Not set'}`);
      Logger.log(`   Previous Status: ${task.Previous_Status || 'Not set'}`);
      Logger.log('');
    });
    
  } catch (error) {
    Logger.log(`ERROR: ${error.toString()}`);
  }
}

