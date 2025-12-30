/**
 * Migration Script: Populate Default Prompts
 * Extracts hardcoded prompts from GeminiAI.gs and populates the prompt sheets
 * Run this once after creating the sheets to migrate existing prompts
 */

/**
 * Migrate all prompts from hardcoded to sheets
 */
function migratePromptsToSheets() {
  try {
    Logger.log('=== Starting Prompt Migration ===');
    
    // Ensure sheets exist
    createAllSheets();
    
    // Migrate Voice Prompts
    migrateVoicePrompts();
    
    // Migrate Email Prompts
    migrateEmailPrompts();
    
    // Migrate MoM Prompts
    migrateMoMPrompts();
    
    Logger.log('=== Prompt Migration Complete ===');
    return { success: true, message: 'All prompts migrated successfully' };
  } catch (error) {
    Logger.log(`Error during prompt migration: ${error.toString()}`);
    return { success: false, error: error.toString() };
  }
}

/**
 * Migrate voice-related prompts
 */
function migrateVoicePrompts() {
  try {
    // parseVoiceCommand prompt (extracted from GeminiAI.gs)
    const parseVoiceCommandPrompt = `You are an expert task management assistant with advanced name recognition and context understanding. Parse the following voice command transcript and extract all relevant information with high accuracy.

Voice command transcript: "{{TRANSCRIPT}}"{{STAFF_CONTEXT}}{{PROJECTS_CONTEXT}}

CRITICAL INSTRUCTIONS FOR NAME MATCHING:
1. When extracting assignee_name, consider PHONETIC VARIATIONS and SOUND-ALIKE names
2. Examples of sound-alike matches:
   - "anaya" heard → match to "anaaya" in staff list
   - "john" heard → match to "Jon" or "Jonathan" in staff list
   - "mike" heard → match to "Michael" in staff list
   - "sarah" heard → match to "Sara" in staff list
3. Use the AVAILABLE STAFF MEMBERS list above to find the BEST MATCH, even if spelling differs slightly
4. If you hear a name that sounds similar to a staff member, use the EXACT name from the staff list
5. Consider common nicknames and variations (e.g., "bob" → "Robert", "jim" → "James")
6. If unsure between multiple matches, prefer the one that sounds most similar phonetically

CRITICAL INSTRUCTIONS FOR PROJECT RECOGNITION:
1. Look for project mentions in the task name, context, or assignee context
2. Use the AVAILABLE PROJECTS list above to match project names
3. Consider partial matches and variations (e.g., "Q4 Report" → "Q4 Financial Report")
4. Project tags can be abbreviations or codes - match both name and tag
5. If project is mentioned indirectly (e.g., "for the website project"), extract it

Extract the following information and return ONLY a valid JSON object (no markdown, no code blocks, no explanations):

{
  "task_name": "concise, clear task title (3-8 words, action-oriented)",
  "assignee_email": "full email address if mentioned explicitly, or null",
  "assignee_name": "full name if mentioned (first and last name if available), or null",
  "due_date": "date in YYYY-MM-DD format if mentioned, or null",
  "due_date_text": "exact original text about due date/timing (e.g., 'next Monday', 'by Friday', 'in 2 days'), or null",
  "due_time": "time if mentioned (HH:MM format in 24-hour), or null",
  "project_tag": "project name or tag if mentioned, or null",
  "context": "all additional details, background, requirements, constraints, or important notes mentioned",
  "tone": "urgent, normal, or calm based on speaker's tone and urgency words",
  "confidence": 0.0-1.0 confidence score (be accurate: 0.9+ if very clear, 0.7-0.9 if mostly clear, 0.5-0.7 if some ambiguity, <0.5 if unclear)",
  "ambiguities": ["specific list of what is unclear or ambiguous (e.g., 'assignee name unclear', 'due date format ambiguous')"]
}

EXTRACTION GUIDELINES:

1. Task Name:
   - Extract the core action/objective (e.g., "Review Q4 report" not "I need to review the Q4 report")
   - Remove filler words like "please", "I want", "can you", "I need", "we should"
   - Keep it concise but descriptive (3-8 words ideal)
   - Use action verbs: "Review", "Update", "Create", "Send", "Schedule", etc.
   - Include key nouns that identify what the task is about

2. Assignee:
   - Look for: "assign to [name]", "give to [name]", "[name] should", "[name] will", "send to [name]", "[name] can", "[name] needs to"
   - IMPORTANT: Match the HEARD name to the EXACT name in AVAILABLE STAFF MEMBERS list using phonetic/sound-alike matching
   - If you hear "anaya" but staff list has "anaaya", use "anaaya" (the exact spelling from staff list)
   - Extract full name if possible (first + last), but prioritize matching to staff list
   - If only first name mentioned, try to match to full name in staff list
   - Consider common name variations: "mike"→"Michael", "bob"→"Robert", "jim"→"James", "sarah"→"Sara"
   - Email: only if explicitly stated (e.g., "john@company.com"), otherwise use null

3. Due Date:
   - Relative dates: "today" = today's date, "tomorrow" = tomorrow, "next Monday" = next Monday, "Friday" = this or next Friday
   - Absolute dates: "January 15th", "15th of January", "01/15/2024" → convert to YYYY-MM-DD
   - Preserve original text in due_date_text for reference
   - If time mentioned (e.g., "by 3 PM", "before noon"), extract in due_time field

4. Tone (indicates urgency level):
   - urgent: ASAP, immediately, critical, important, rush
   - normal: standard, regular (default if not mentioned)
   - calm: whenever, no rush, eventually

5. Project/Tag:
   - Look for: "for [project]", "related to [project]", "under [project]", "in [project]", "for the [project]", project names mentioned
   - IMPORTANT: Match to EXACT project name or tag from AVAILABLE PROJECTS list
   - Consider context: if task mentions "website", "Q4 report", "marketing campaign", match to relevant project
   - Use partial matching: "Q4 Report" can match "Q4 Financial Report" project
   - If project is mentioned indirectly, infer from context (e.g., "update the homepage" → "Website" project)
   - Return the Project_Tag value from the projects list if found, otherwise return project name

6. Context:
   - Include: why the task exists, background info, specific requirements, constraints, dependencies
   - Include: location, tools needed, people involved, budget constraints, etc.
   - Be comprehensive but concise

7. Tone:
   - Urgent: rushed speech, urgent words, time pressure mentioned
   - Normal: standard conversational tone
   - Calm: relaxed, no rush indicated

8. Confidence:
   - 0.9-1.0: All key info clear (task, assignee, date all clear)
   - 0.7-0.9: Most info clear, minor ambiguities
   - 0.5-0.7: Some key info missing or unclear
   - <0.5: Major ambiguities or unclear command

9. Ambiguities:
   - List specific unclear elements: "assignee name unclear", "due date ambiguous", "task scope vague"

EXAMPLES WITH NAME AND PROJECT MATCHING:

Input: "Assign anaya to review the Q4 financial report by next Friday, it's urgent"
Available Staff: - Anaaya Udhas (anaayaudhas92@gmail.com)
Available Projects: - Q4 Financial Report (Tag: Q4_Report)
Output: {
  "task_name": "Review Q4 financial report",
  "assignee_email": null,
  "assignee_name": "Anaaya Udhas",
  "due_date": "2024-01-19",
  "due_date_text": "next Friday",
  "due_time": null,
  "project_tag": "Q4_Report",
  "context": "Q4 financial report needs review, urgent",
  "tone": "urgent",
  "confidence": 0.95,
  "ambiguities": []
}

Input: "I need someone to update the website homepage, maybe by end of week"
Available Projects: - Website Redesign (Tag: Website)
Output: {
  "task_name": "Update website homepage",
  "assignee_email": null,
  "assignee_name": null,
  "due_date": "2024-01-19",
  "due_date_text": "end of week",
  "due_time": null,
  "project_tag": "Website",
  "context": "Website homepage needs updating",
  "tone": "normal",
  "confidence": 0.75,
  "ambiguities": ["assignee not specified"]
}

Input: "Give this to mike, he should handle the marketing campaign launch"
Available Staff: - Michael Johnson (michael.j@company.com)
Available Projects: - Marketing Campaign (Tag: Marketing)
Output: {
  "task_name": "Handle marketing campaign launch",
  "assignee_email": null,
  "assignee_name": "Michael Johnson",
  "due_date": null,
  "due_date_text": null,
  "due_time": null,
  "project_tag": "Marketing",
  "context": "Marketing campaign launch needs to be handled",
  "tone": "normal",
  "confidence": 0.9,
  "ambiguities": []
}

Now parse this voice command: "{{TRANSCRIPT}}"`;

    savePrompt('parseVoiceCommand', 'voice', parseVoiceCommandPrompt, 'Parses voice command transcripts and extracts structured task data');
    
    Logger.log('✓ Migrated voice prompts');
  } catch (error) {
    Logger.log(`Error migrating voice prompts: ${error.toString()}`);
  }
}

/**
 * Migrate email-related prompts
 */
function migrateEmailPrompts() {
  try {
    // classifyReplyType prompt (extracted from GeminiAI.gs)
    const classifyReplyTypePrompt = `You are an expert at analyzing email replies about task assignments. Classify this email reply into one of these categories:

1. ACCEPTANCE - The assignee accepts the task and agrees to the deadline. Examples: "I accept", "I'll do it", "Got it", "Will complete by deadline", "Acknowledged" (if accepting)
2. DATE_CHANGE - The assignee requests a different due date or mentions a date that's not feasible. Look for: "not feasible", "can't make that date", "need more time", "deadline is", "by [date]", "feasible as deadline", "propose [date]", "suggest [date]", "prefer [date]", "10th Jan", "January 10th", "10/01/2025", etc.
3. SCOPE_QUESTION - The assignee has questions about what needs to be done, needs clarification, or asks "what does this mean", "can you clarify", "I need more details"
4. ROLE_REJECTION - The assignee says this is not their job, responsibility, or role. Examples: "not my responsibility", "wrong person", "not my role", "should be assigned to"
5. OTHER - Something else that doesn't fit the above categories

IMPORTANT DATE EXTRACTION RULES:
- IGNORE dates in quoted email sections (text after "On [date] wrote:" or lines starting with ">")
- IGNORE dates in email signatures
- Extract ONLY the PROPOSED/NEW date from the actual reply text (the part the employee wrote)
- If the email mentions ANY date in the actual reply (not quoted text), and that date is different from the original deadline, classify as DATE_CHANGE
- Look for phrases like: "7th of Jan 2026", "10th Jan is feasible", "deadline is 10th Jan", "by 10th Jan", "10th Jan works", "can do it by 10th Jan", "I can do it by the 7th of Jan 2026"
- Convert ALL date formats to YYYY-MM-DD:
  * "7th of Jan 2026" or "7th of January 2026" → 2026-01-07
  * "10th Jan" or "10 Jan" → Assume current year if year not mentioned, or next year if date has passed
  * "Jan 10th" or "January 10" → Same as above
  * "10/01/2025" → 2025-01-10 (assume DD/MM/YYYY format)
  * "2025-01-10" → Keep as is
  * Relative dates like "next Monday", "in 2 weeks" → Calculate actual date
- If multiple dates are mentioned, extract the one that appears to be the NEW/PROPOSED deadline (usually the one mentioned with "by", "deadline", "feasible", "can do")
- If the email says "not feasible" or "can't make that date" and mentions another date, that other date is the proposed date
- PRIORITIZE dates that appear in sentences with words like "by", "deadline", "feasible", "can do", "I can"

Email reply (already cleaned - quoted text removed): "{{REPLY_CONTENT}}"{{DATE_CONTEXT}}

Return ONLY a valid JSON object (no markdown, no code blocks):
{
  "type": "ACCEPTANCE|DATE_CHANGE|SCOPE_QUESTION|ROLE_REJECTION|OTHER",
  "confidence": 0.0-1.0,
  "extracted_date": "proposed date in YYYY-MM-DD format if DATE_CHANGE (MUST be the NEW date being requested, not the original. Convert all formats: '10th Jan' → '2025-01-10', 'Jan 10' → '2025-01-10', '10/01/2025' → '2025-01-10', 'next Monday' → calculate actual date). If no clear date, return null",
  "reasoning": "brief explanation of why this classification was chosen and what date was extracted (if any)"
}`;

    savePrompt('classifyReplyType', 'email', classifyReplyTypePrompt, 'Classifies email replies into categories (ACCEPTANCE, DATE_CHANGE, SCOPE_QUESTION, ROLE_REJECTION, OTHER)');
    
    // generateAssignmentEmail prompt
    const generateAssignmentEmailPrompt = `You are a professional Chief of Staff assistant. Write a formal but friendly email assigning a task to a team member.

Task Name: {{TASK_NAME}}
Assignee: {{ASSIGNEE_NAME}}
Due Date: {{DUE_DATE}}
Context: {{CONTEXT}}

Write an email that:
1. Is polite but authoritative
2. Clearly states the task and expectations
3. Asks for confirmation of the deadline
4. Invites questions or concerns
5. Uses a professional but friendly tone

Return ONLY the email body text (no subject line, no signature - those will be added separately).`;

    savePrompt('generateAssignmentEmail', 'email', generateAssignmentEmailPrompt, 'Generates task assignment emails');
    
    // summarizeReviewRequest prompt (simplified version)
    const summarizeReviewRequestPrompt = `An employee has sent a reply about a task. Summarize their message.

Task: "{{TASK_NAME}}"
Employee's Reply: "{{EMPLOYEE_REPLY}}"
Review Type: {{REVIEW_TYPE}}

Generate a brief summary (2-3 sentences) explaining what the employee is communicating.

Return ONLY the summary text (no markdown, no JSON, just plain text).`;

    savePrompt('summarizeReviewRequest', 'email', summarizeReviewRequestPrompt, 'Summarizes employee review requests');
    
    Logger.log('✓ Migrated email prompts');
  } catch (error) {
    Logger.log(`Error migrating email prompts: ${error.toString()}`);
  }
}

/**
 * Migrate MoM-related prompts
 */
function migrateMoMPrompts() {
  try {
    // analyzeMoMDocument prompt (extracted from GeminiAI.gs)
    const analyzeMoMDocumentPrompt = `You are analyzing a manually created Minutes of Meeting (MoM) document. Extract the following information:

1. Action Items: List all action items mentioned in the document. For each action item, extract:
   - description: What needs to be done
   - owner: Who is responsible (name or email if mentioned)
   - deadline: When it's due (extract date if mentioned, format as YYYY-MM-DD or null)
   - priority: High, Medium, or Low (based on urgency indicators)

2. Projects: Identify all projects mentioned in the document. For each project, extract:
   - project_tag: The project identifier/tag (if mentioned, or infer from project name)
   - project_name: The full project name
   - summary: Brief summary of what was discussed about this project
   - key_points: Array of key points discussed (max 10 points)
   - decisions: Array of decisions made about the project (max 10 decisions)
   - include_full_context: Boolean - true if full document text should be stored

3. Executive Summary: A brief summary of the entire meeting (2-3 sentences)

Return your response as a JSON object with this structure:
{
  "action_items": [
    {
      "description": "string",
      "owner": "string or null",
      "deadline": "YYYY-MM-DD or null",
      "priority": "High|Medium|Low"
    }
  ],
  "projects": [
    {
      "project_tag": "string or null",
      "project_name": "string",
      "summary": "string",
      "key_points": ["string"],
      "decisions": ["string"],
      "include_full_context": true
    }
  ],
  "executive_summary": "string"
}

MoM document text:
"{{MOM_TEXT}}"

Return your response as a JSON object with this structure.`;

    savePrompt('analyzeMoMDocument', 'mom', analyzeMoMDocumentPrompt, 'Analyzes manually created MoM documents and extracts action items and project knowledge');
    
    Logger.log('✓ Migrated MoM prompts');
  } catch (error) {
    Logger.log(`Error migrating MoM prompts: ${error.toString()}`);
  }
}

