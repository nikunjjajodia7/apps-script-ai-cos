/**
 * Email Helper Utilities
 * Functions for sending and processing emails via Gmail
 */

/**
 * Get email template from Config sheet
 */
function getEmailTemplate(templateType, variables = {}) {
  const config = getConfig();
  const templateKey = `EMAIL_TEMPLATE_${templateType.toUpperCase()}`;
  let template = config[templateKey];
  
  if (!template) {
    // Fallback templates
    if (templateType === 'assignment') {
      template = `Hello {{ASSIGNEE_NAME}},

On behalf of {{BOSS_NAME}}, I'm assigning you the following task:

• Task: {{TASK_NAME}}
• Context: {{CONTEXT}}
• Due Date: {{DUE_DATE}}

Please confirm that you can meet this deadline. If you have any questions, need clarification, or foresee any issues, let me know by replying to this email. If the date is not feasible, please propose an alternate date.

Thank you,
{{SIGNATURE}}`;
    } else if (templateType === 'followup') {
      template = `Hello {{ASSIGNEE_NAME}},

I wanted to follow up on the task assignment below. Please confirm receipt and your availability to complete this by {{DUE_DATE}}.

• Task: {{TASK_NAME}}
• Due Date: {{DUE_DATE}}

If you have any concerns, please reply to this email.

Thank you,
{{SIGNATURE}}`;
    } else if (templateType === 'escalation') {
      template = `Hello {{ASSIGNEE_NAME}},

This is a follow-up regarding the task assignment below. We haven't received a response yet, and this requires your attention.

• Task: {{TASK_NAME}}
• Due Date: {{DUE_DATE}}

Please respond as soon as possible to confirm you can complete this task.

Thank you,
{{SIGNATURE}}`;
    }
  }
  
  // Replace variables
  const bossEmail = CONFIG.BOSS_EMAIL();
  const bossName = bossEmail.split('@')[0];
  
  const replacements = {
    '{{ASSIGNEE_NAME}}': variables.assigneeName || variables.assignee_email || 'Team Member',
    '{{BOSS_NAME}}': bossName,
    '{{TASK_NAME}}': variables.taskName || variables.Task_Name || 'Task',
    '{{CONTEXT}}': variables.context || variables.Context_Hidden || 'No additional context',
    '{{DUE_DATE}}': variables.dueDate || variables.Due_Date || 'Not specified',
    '{{SIGNATURE}}': CONFIG.EMAIL_SIGNATURE(),
  };
  
  let emailBody = template;
  Object.keys(replacements).forEach(key => {
    emailBody = emailBody.replace(new RegExp(key, 'g'), replacements[key]);
  });
  
  return emailBody;
}

/**
 * Send task assignment email
 */
function sendTaskAssignmentEmail(taskId) {
  try {
    const task = getTask(taskId);
    if (!task) {
      throw new Error(`Task ${taskId} not found`);
    }
    
    if (!task.Assignee_Email) {
      throw new Error(`Task ${taskId} has no assignee`);
    }
    
    const staff = getStaff(task.Assignee_Email);
    const assigneeName = staff ? staff.Name : task.Assignee_Email.split('@')[0];
    
    // Generate email using AI or template
    let emailBody;
    try {
      emailBody = generateAssignmentEmail(task);
    } catch (error) {
      Logger.log('AI email generation failed, using template: ' + error);
      emailBody = getEmailTemplate('assignment', {
        assigneeName: assigneeName,
        taskName: task.Task_Name,
        context: task.Context_Hidden || '',
        dueDate: task.Due_Date ? Utilities.formatDate(new Date(task.Due_Date), Session.getScriptTimeZone(), 'MMMM d, yyyy') : 'Not specified',
      });
    }
    
    const subject = `Task Assignment: ${task.Task_Name}`;
    
    // Add task correlation token in email headers
    const taskToken = `TASK-${taskId}`;
    const emailOptions = {
      htmlBody: emailBody.replace(/\n/g, '<br>'),
      name: 'Chief of Staff AI',
      replyTo: CONFIG.BOSS_EMAIL(),
    };
    
    // Send email
    GmailApp.sendEmail(
      task.Assignee_Email,
      subject,
      emailBody,
      emailOptions
    );
    
    // Store email thread ID for correlation
    // Note: GmailApp.sendEmail doesn't return thread ID directly
    // We'll search for it or use a different approach
    const emailThread = GmailApp.search(`to:${task.Assignee_Email} subject:"${subject}"`, 0, 1)[0];
    if (emailThread) {
      // Store thread ID in task (we can add a Thread_ID column or use Interaction_Log)
      logInteraction(taskId, `Assignment email sent to ${task.Assignee_Email}. Thread ID: ${emailThread.getId()}`);
    } else {
      logInteraction(taskId, `Assignment email sent to ${task.Assignee_Email}`);
    }
    
    return true;
  } catch (error) {
    logError(ERROR_TYPE.API_ERROR, 'sendTaskAssignmentEmail', error.toString(), taskId, error.stack);
    throw error;
  }
}

/**
 * Send follow-up email
 */
function sendFollowUpEmail(taskId) {
  try {
    const task = getTask(taskId);
    if (!task || !task.Assignee_Email) {
      return false;
    }
    
    const staff = getStaff(task.Assignee_Email);
    const assigneeName = staff ? staff.Name : task.Assignee_Email.split('@')[0];
    
    const emailBody = getEmailTemplate('followup', {
      assigneeName: assigneeName,
      taskName: task.Task_Name,
      dueDate: task.Due_Date ? Utilities.formatDate(new Date(task.Due_Date), Session.getScriptTimeZone(), 'MMMM d, yyyy') : 'Not specified',
    });
    
    const subject = `Follow-up: Task Assignment: ${task.Task_Name}`;
    
    GmailApp.sendEmail(
      task.Assignee_Email,
      subject,
      emailBody,
      {
        htmlBody: emailBody.replace(/\n/g, '<br>'),
        name: 'Chief of Staff AI',
        replyTo: CONFIG.BOSS_EMAIL(),
      }
    );
    
    logInteraction(taskId, `Follow-up email sent to ${task.Assignee_Email}`);
    return true;
  } catch (error) {
    logError(ERROR_TYPE.API_ERROR, 'sendFollowUpEmail', error.toString(), taskId);
    return false;
  }
}

/**
 * Send escalation email
 */
function sendEscalationEmail(taskId) {
  try {
    const task = getTask(taskId);
    if (!task || !task.Assignee_Email) {
      return false;
    }
    
    const staff = getStaff(task.Assignee_Email);
    const assigneeName = staff ? staff.Name : task.Assignee_Email.split('@')[0];
    
    const emailBody = getEmailTemplate('escalation', {
      assigneeName: assigneeName,
      taskName: task.Task_Name,
      dueDate: task.Due_Date ? Utilities.formatDate(new Date(task.Due_Date), Session.getScriptTimeZone(), 'MMMM d, yyyy') : 'Not specified',
    });
    
    const subject = `Urgent: Task Assignment: ${task.Task_Name}`;
    
    GmailApp.sendEmail(
      task.Assignee_Email,
      subject,
      emailBody,
      {
        htmlBody: emailBody.replace(/\n/g, '<br>'),
        name: 'Chief of Staff AI',
        replyTo: CONFIG.BOSS_EMAIL(),
      }
    );
    
    logInteraction(taskId, `Escalation email sent to ${task.Assignee_Email}`);
    return true;
  } catch (error) {
    logError(ERROR_TYPE.API_ERROR, 'sendEscalationEmail', error.toString(), taskId);
    return false;
  }
}

/**
 * Send confirmation email to Boss
 */
function sendBossConfirmation(taskId, summary) {
  try {
    const task = getTask(taskId);
    if (!task) return;
    
    const bossEmail = CONFIG.BOSS_EMAIL();
    const subject = `Task Created: ${task.Task_Name}`;
    
    const emailBody = `Hello,

I've created a new task based on your voice command:

${summary}

Task ID: ${taskId}
Status: ${task.Status}

You can review and manage this task in your dashboard.

Best regards,
Chief of Staff AI`;

    GmailApp.sendEmail(
      bossEmail,
      subject,
      emailBody,
      {
        name: 'Chief of Staff AI',
      }
    );
    
    logInteraction(taskId, `Confirmation email sent to Boss`);
  } catch (error) {
    logError(ERROR_TYPE.API_ERROR, 'sendBossConfirmation', error.toString(), taskId);
  }
}

/**
 * Find email thread for a task
 */
function findTaskEmailThread(taskId) {
  const task = getTask(taskId);
  if (!task || !task.Assignee_Email) {
    return null;
  }
  
  const subject = `Task Assignment: ${task.Task_Name}`;
  const threads = GmailApp.search(`to:${task.Assignee_Email} subject:"${subject}"`, 0, 5);
  
  // Return most recent thread
  return threads.length > 0 ? threads[0] : null;
}

/**
 * Get replies to a task assignment email
 */
function getTaskReplies(taskId) {
  const thread = findTaskEmailThread(taskId);
  if (!thread) {
    return [];
  }
  
  const messages = thread.getMessages();
  // Skip first message (our assignment email)
  return messages.slice(1);
}

