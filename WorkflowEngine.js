/**
 * Workflow Engine
 * Executes workflows based on triggers, conditions, and actions
 */

/**
 * Load all active workflows from Workflows sheet
 * @returns {Array} Array of workflow objects
 */
function loadWorkflows() {
  try {
    const workflows = getSheetData(SHEETS.WORKFLOWS);
    // Filter only active workflows
    return workflows.filter(w => w.Active === true || w.Active === 'TRUE' || w.Active === 'true');
  } catch (error) {
    Logger.log(`Error loading workflows: ${error.toString()}`);
    return [];
  }
}

/**
 * Execute workflows matching a trigger event
 * @param {string} triggerEvent - The event that triggered workflow execution (e.g., 'status_change', 'task_created', 'email_reply')
 * @param {object} context - Context object with task data, event details, etc.
 * @returns {Array} Array of results from executed workflows
 */
function executeWorkflow(triggerEvent, context) {
  try {
    const workflows = loadWorkflows();
    const results = [];
    
    for (const workflow of workflows) {
      // Parse workflow JSON
      let workflowConfig;
      try {
        workflowConfig = typeof workflow.Trigger_Event === 'string' 
          ? JSON.parse(workflow.Trigger_Event) 
          : workflow.Trigger_Event;
      } catch (e) {
        // Try parsing the entire workflow as JSON
        try {
          const workflowJson = workflow.Trigger_Event || workflow.Conditions || workflow.Actions;
          if (typeof workflowJson === 'string') {
            workflowConfig = JSON.parse(workflowJson);
          } else {
            // Workflow might be stored in separate columns
            workflowConfig = {
              trigger: workflow.Trigger_Event,
              conditions: workflow.Conditions ? (typeof workflow.Conditions === 'string' ? JSON.parse(workflow.Conditions) : workflow.Conditions) : {},
              actions: workflow.Actions ? (typeof workflow.Actions === 'string' ? JSON.parse(workflow.Actions) : workflow.Actions) : [],
              timing: workflow.Timing ? (typeof workflow.Timing === 'string' ? JSON.parse(workflow.Timing) : workflow.Timing) : {}
            };
          }
        } catch (e2) {
          Logger.log(`Error parsing workflow ${workflow.Workflow_ID || workflow.Name}: ${e2.toString()}`);
          continue;
        }
      }
      
      // Check if trigger matches
      const workflowTrigger = workflowConfig.trigger || workflow.Trigger_Event;
      if (workflowTrigger !== triggerEvent) {
        continue; // Skip this workflow
      }
      
      // Evaluate conditions
      const conditions = workflowConfig.conditions || (workflow.Conditions ? (typeof workflow.Conditions === 'string' ? JSON.parse(workflow.Conditions) : workflow.Conditions) : {});
      if (!evaluateConditions(conditions, context)) {
        continue; // Conditions not met, skip
      }
      
      // Execute actions
      const actions = workflowConfig.actions || (workflow.Actions ? (typeof workflow.Actions === 'string' ? JSON.parse(workflow.Actions) : workflow.Actions) : []);
      const actionResults = executeActions(actions, context, workflow.Workflow_ID || workflow.Name);
      
      results.push({
        workflowId: workflow.Workflow_ID || workflow.Name,
        workflowName: workflow.Name,
        executed: true,
        actions: actionResults
      });
    }
    
    return results;
  } catch (error) {
    Logger.log(`Error executing workflow for trigger ${triggerEvent}: ${error.toString()}`);
    return [];
  }
}

/**
 * Evaluate workflow conditions against context
 * @param {object} conditions - Conditions object (e.g., {status: 'Assigned', assignee_exists: true})
 * @param {object} context - Context object with task data
 * @returns {boolean} True if all conditions are met
 */
function evaluateConditions(conditions, context) {
  if (!conditions || Object.keys(conditions).length === 0) {
    return true; // No conditions means always execute
  }
  
  try {
    for (const [key, value] of Object.entries(conditions)) {
      let contextValue;
      
      // Handle nested properties (e.g., 'task.status')
      if (key.includes('.')) {
        const parts = key.split('.');
        contextValue = context;
        for (const part of parts) {
          contextValue = contextValue?.[part];
        }
      } else {
        contextValue = context[key] || context.task?.[key];
      }
      
      // Compare values
      if (typeof value === 'object' && value !== null) {
        // Handle operators like {operator: '>', value: 24}
        if (value.operator) {
          const operator = value.operator;
          const compareValue = value.value;
          
          switch (operator) {
            case '>':
              if (!(parseFloat(contextValue) > parseFloat(compareValue))) return false;
              break;
            case '>=':
              if (!(parseFloat(contextValue) >= parseFloat(compareValue))) return false;
              break;
            case '<':
              if (!(parseFloat(contextValue) < parseFloat(compareValue))) return false;
              break;
            case '<=':
              if (!(parseFloat(contextValue) <= parseFloat(compareValue))) return false;
              break;
            case '==':
            case '===':
              if (contextValue != compareValue) return false;
              break;
            case '!=':
            case '!==':
              if (contextValue == compareValue) return false;
              break;
            case 'in':
              if (!Array.isArray(compareValue) || !compareValue.includes(contextValue)) return false;
              break;
            case 'not_in':
              if (Array.isArray(compareValue) && compareValue.includes(contextValue)) return false;
              break;
            default:
              Logger.log(`Unknown operator: ${operator}`);
              return false;
          }
        } else {
          // Object comparison (deep equality check)
          if (JSON.stringify(contextValue) !== JSON.stringify(value)) return false;
        }
      } else {
        // Simple equality check
        if (contextValue != value) return false;
      }
    }
    
    return true; // All conditions met
  } catch (error) {
    Logger.log(`Error evaluating conditions: ${error.toString()}`);
    return false; // Fail safe: don't execute if condition evaluation fails
  }
}

/**
 * Execute workflow actions with timing
 * @param {Array} actions - Array of action objects
 * @param {object} context - Context object
 * @param {string} workflowId - Workflow ID for logging
 * @returns {Array} Array of action execution results
 */
function executeActions(actions, context, workflowId) {
  if (!Array.isArray(actions) || actions.length === 0) {
    return [];
  }
  
  const results = [];
  const taskId = context.taskId || context.task?.Task_ID;
  
  for (const action of actions) {
    try {
      const actionType = action.type || action.action;
      const delayHours = action.delay_hours || action.delayHours || 0;
      const params = action.params || action.parameters || {};
      
      // If delay is specified, schedule the action (for now, we'll execute immediately but log the delay)
      // In a production system, you'd use Apps Script triggers or a queue system
      if (delayHours > 0) {
        Logger.log(`Workflow ${workflowId}: Action ${actionType} scheduled with ${delayHours} hour delay`);
        // For now, we'll execute immediately but note the delay
        // TODO: Implement proper delayed execution using time-based triggers
      }
      
      let result = { type: actionType, executed: false, error: null };
      
      switch (actionType) {
        case 'send_email':
          if (taskId) {
            try {
              sendTaskAssignmentEmail(taskId);
              result.executed = true;
              result.message = 'Email sent';
            } catch (e) {
              result.error = e.toString();
            }
          }
          break;
          
        case 'update_status':
          if (taskId && params.status) {
            try {
              updateTask(taskId, { Status: params.status });
              result.executed = true;
              result.message = `Status updated to ${params.status}`;
            } catch (e) {
              result.error = e.toString();
            }
          }
          break;
          
        case 'log_interaction':
          if (taskId && params.message) {
            try {
              logInteraction(taskId, params.message);
              result.executed = true;
              result.message = 'Interaction logged';
            } catch (e) {
              result.error = e.toString();
            }
          }
          break;
          
        case 'send_followup':
          if (taskId) {
            try {
              sendFollowUpEmail(taskId);
              result.executed = true;
              result.message = 'Follow-up email sent';
            } catch (e) {
              result.error = e.toString();
            }
          }
          break;
          
        case 'escalate_to_boss':
          if (taskId) {
            try {
              sendBossAlert(taskId);
              result.executed = true;
              result.message = 'Boss alerted';
            } catch (e) {
              result.error = e.toString();
            }
          }
          break;
          
        case 'update_priority':
          if (taskId && params.priority) {
            try {
              updateTask(taskId, { Priority: params.priority });
              result.executed = true;
              result.message = `Priority updated to ${params.priority}`;
            } catch (e) {
              result.error = e.toString();
            }
          }
          break;
          
        case 'assign_task':
          if (taskId && params.assigneeEmail) {
            try {
              updateTask(taskId, { 
                Assignee_Email: params.assigneeEmail,
                Assignee_Name: params.assigneeName || '',
                Status: TASK_STATUS.NOT_ACTIVE
              });
              sendTaskAssignmentEmail(taskId);
              result.executed = true;
              result.message = `Task assigned to ${params.assigneeEmail}`;
            } catch (e) {
              result.error = e.toString();
            }
          }
          break;
          
        default:
          result.error = `Unknown action type: ${actionType}`;
          Logger.log(`Workflow ${workflowId}: Unknown action type: ${actionType}`);
      }
      
      results.push(result);
    } catch (error) {
      Logger.log(`Error executing action in workflow ${workflowId}: ${error.toString()}`);
      results.push({ 
        type: action.type || 'unknown', 
        executed: false, 
        error: error.toString() 
      });
    }
  }
  
  return results;
}

/**
 * Test workflow execution with sample context (dry run)
 * @param {object} workflow - Workflow object
 * @param {object} sampleContext - Sample context for testing
 * @returns {object} Test results
 */
function testWorkflow(workflow, sampleContext) {
  try {
    // Parse workflow if needed
    let workflowConfig = workflow;
    if (typeof workflow.Trigger_Event === 'string') {
      try {
        workflowConfig = JSON.parse(workflow.Trigger_Event);
      } catch (e) {
        // Use workflow as-is
      }
    }
    
    const conditions = workflowConfig.conditions || (workflow.Conditions ? (typeof workflow.Conditions === 'string' ? JSON.parse(workflow.Conditions) : workflow.Conditions) : {});
    const actions = workflowConfig.actions || (workflow.Actions ? (typeof workflow.Actions === 'string' ? JSON.parse(workflow.Actions) : workflow.Actions) : []);
    
    const conditionsMet = evaluateConditions(conditions, sampleContext);
    const actionResults = conditionsMet ? executeActions(actions, sampleContext, workflow.Workflow_ID || 'test') : [];
    
    return {
      workflowId: workflow.Workflow_ID || workflow.Name,
      conditionsMet: conditionsMet,
      actions: actionResults,
      wouldExecute: conditionsMet && actions.length > 0
    };
  } catch (error) {
    return {
      error: error.toString(),
      wouldExecute: false
    };
  }
}

