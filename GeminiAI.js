/**
 * Vertex AI Gemini Integration
 * Functions for calling Gemini models via Vertex AI API
 */

/**
 * Call Vertex AI API using UrlFetchApp
 * Note: This requires proper authentication setup in Google Cloud
 */
function callVertexAI(modelName, prompt, options = {}) {
  try {
    const projectId = CONFIG.VERTEX_AI_PROJECT_ID();
    const location = CONFIG.VERTEX_AI_LOCATION();
    
    if (!projectId) {
      throw new Error('VERTEX_AI_PROJECT_ID not configured');
    }
    
    const url = `https://${location}-aiplatform.googleapis.com/v1/projects/${projectId}/locations/${location}/publishers/google/models/${modelName}:generateContent`;
    
    const payload = {
      contents: [{
        role: 'user',
        parts: [{
          text: prompt
        }]
      }],
      generationConfig: {
        temperature: options.temperature || 0.7,
        topK: options.topK || 40,
        topP: options.topP || 0.95,
        maxOutputTokens: options.maxOutputTokens || 8192,
      }
    };
    
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      headers: {
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken(),
        'Content-Type': 'application/json',
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    });
    
    if (response.getResponseCode() !== 200) {
      const errorText = response.getContentText();
      logError(ERROR_TYPE.API_ERROR, 'callVertexAI', `Vertex AI API error: ${errorText}`);
      throw new Error(`Vertex AI API error: ${errorText}`);
    }
    
    const result = JSON.parse(response.getContentText());
    return result.candidates[0].content.parts[0].text;
    
  } catch (error) {
    logError(ERROR_TYPE.API_ERROR, 'callVertexAI', error.toString(), null, error.stack);
    throw error;
  }
}

/**
 * Call Gemini Flash for quick tasks
 */
function callGeminiFlash(prompt, options = {}) {
  return callVertexAI(CONFIG.GEMINI_FLASH_MODEL, prompt, {
    ...options,
    temperature: options.temperature || 0.3, // Lower temperature for more consistent results
  });
}

/**
 * Call Gemini Pro for complex reasoning
 */
function callGeminiPro(prompt, options = {}) {
  return callVertexAI(CONFIG.GEMINI_PRO_MODEL, prompt, {
    ...options,
    temperature: options.temperature || 0.7,
  });
}

/**
 * Transcribe audio file using Gemini (multimodal)
 * Note: This requires the audio file to be accessible
 */
function transcribeAudio(audioFileId) {
  try {
    Logger.log('Starting audio transcription...');
    const file = DriveApp.getFileById(audioFileId);
    const fileName = file.getName();
    Logger.log(`File: ${fileName}`);
    
    const audioBlob = file.getBlob();
    const mimeType = audioBlob.getContentType();
    Logger.log(`MIME type: ${mimeType}`);
    
    // Check file size (Vertex AI has limits)
    const fileSize = audioBlob.getBytes().length;
    Logger.log(`File size: ${(fileSize / 1024 / 1024).toFixed(2)} MB`);
    
    if (fileSize > 20 * 1024 * 1024) { // 20MB limit
      throw new Error('Audio file too large (max 20MB)');
    }
    
    const base64Audio = Utilities.base64Encode(audioBlob.getBytes());
    Logger.log('Audio encoded to base64');
    
    const projectId = CONFIG.VERTEX_AI_PROJECT_ID();
    const location = CONFIG.VERTEX_AI_LOCATION();
    
    if (!projectId) {
      throw new Error('VERTEX_AI_PROJECT_ID not configured in Config sheet');
    }
    
    Logger.log(`Project ID: ${projectId}, Location: ${location}`);
    
    const url = `https://${location}-aiplatform.googleapis.com/v1/projects/${projectId}/locations/${location}/publishers/google/models/${CONFIG.GEMINI_FLASH_MODEL}:generateContent`;
    Logger.log(`API URL: ${url}`);
    
    const payload = {
      contents: [{
        role: 'user',
        parts: [
          {
            text: 'Transcribe this audio file. Return only the spoken text, no additional commentary.'
          },
          {
            inlineData: {
              mimeType: mimeType,
              data: base64Audio
            }
          }
        ]
      }],
      generationConfig: {
        temperature: 0.3,
        maxOutputTokens: 4096,
      }
    };
    
    Logger.log('Sending request to Vertex AI...');
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      headers: {
        'Authorization': 'Bearer ' + ScriptApp.getOAuthToken(),
        'Content-Type': 'application/json',
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    });
    
    const responseCode = response.getResponseCode();
    Logger.log(`Response code: ${responseCode}`);
    
    if (responseCode !== 200) {
      const errorText = response.getContentText();
      Logger.log(`Error response: ${errorText}`);
      logError(ERROR_TYPE.API_ERROR, 'transcribeAudio', `Audio transcription error (${responseCode}): ${errorText}`);
      throw new Error(`Audio transcription error (${responseCode}): ${errorText}`);
    }
    
    const result = JSON.parse(response.getContentText());
    Logger.log('Response parsed successfully');
    
    if (!result.candidates || !result.candidates[0] || !result.candidates[0].content || !result.candidates[0].content.parts || !result.candidates[0].content.parts[0]) {
      Logger.log('Unexpected response structure: ' + JSON.stringify(result));
      throw new Error('Unexpected response structure from Vertex AI');
    }
    
    const transcript = result.candidates[0].content.parts[0].text;
    Logger.log(`Transcript received (${transcript.length} characters)`);
    
    return transcript;
    
  } catch (error) {
    Logger.log('ERROR in transcribeAudio: ' + error.toString());
    Logger.log('Stack: ' + (error.stack || 'No stack trace'));
    logError(ERROR_TYPE.API_ERROR, 'transcribeAudio', error.toString(), null, error.stack);
    throw error;
  }
}

/**
 * Parse voice command and extract structured data
 */
function parseVoiceCommand(transcript) {
  const prompt = `You are a task management assistant. Parse the following voice command and extract structured information.

Voice command transcript: "${transcript}"

Extract the following information and return ONLY a valid JSON object (no markdown, no code blocks):
{
  "task_name": "concise task title",
  "assignee_email": "email if mentioned, or null",
  "assignee_name": "name if mentioned, or null",
  "due_date": "date in YYYY-MM-DD format if mentioned, or null",
  "due_date_text": "original text about due date, or null",
  "context": "additional context or details",
  "tone": "urgent, normal, or calm",
  "confidence": 0.0-1.0 confidence score,
  "ambiguities": ["list of any ambiguous elements"]
}

If the assignee is mentioned by name but not email, try to infer the email from common patterns, or set to null.
If due date is relative (e.g., "next Monday"), calculate the actual date.
Be conservative with confidence scores - if anything is unclear, lower the confidence.`;

  try {
    const response = callGeminiFlash(prompt);
    
    // Extract JSON from response (handle markdown code blocks if present)
    let jsonText = response.trim();
    if (jsonText.startsWith('```')) {
      jsonText = jsonText.replace(/^```json\n?/, '').replace(/```$/, '');
    }
    if (jsonText.startsWith('```')) {
      jsonText = jsonText.replace(/^```\n?/, '').replace(/```$/, '');
    }
    
    const parsed = JSON.parse(jsonText);
    
    // Process due date if it's text
    if (parsed.due_date_text && !parsed.due_date) {
      parsed.due_date = parseRelativeDate(parsed.due_date_text);
    }
    
    return parsed;
  } catch (error) {
    logError(ERROR_TYPE.API_ERROR, 'parseVoiceCommand', error.toString());
    return {
      task_name: transcript,
      assignee_email: null,
      assignee_name: null,
      due_date: null,
      context: '',
      tone: 'normal',
      confidence: 0.3,
      ambiguities: ['Failed to parse command']
    };
  }
}

/**
 * Parse relative date text to actual date
 */
function parseRelativeDate(dateText) {
  const today = new Date();
  const lowerText = dateText.toLowerCase();
  
  // Simple relative date parsing
  if (lowerText.includes('today')) {
    return Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  if (lowerText.includes('tomorrow')) {
    const tomorrow = new Date(today);
    tomorrow.setDate(tomorrow.getDate() + 1);
    return Utilities.formatDate(tomorrow, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  if (lowerText.includes('next monday') || lowerText.includes('next week')) {
    const daysUntilMonday = (8 - today.getDay()) % 7 || 7;
    const nextMonday = new Date(today);
    nextMonday.setDate(today.getDate() + daysUntilMonday);
    return Utilities.formatDate(nextMonday, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  
  // Try to parse as date
  try {
    const parsed = new Date(dateText);
    if (!isNaN(parsed.getTime())) {
      return Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
  } catch (e) {
    // Ignore
  }
  
  return null;
}

/**
 * Generate task assignment email using Gemini Pro
 */
function generateAssignmentEmail(task) {
  const staff = getStaff(task.Assignee_Email);
  const assigneeName = staff ? staff.Name : task.Assignee_Email;
  
  const prompt = `You are a professional Chief of Staff assistant. Write a formal but friendly email assigning a task to a team member.

Task Name: ${task.Task_Name}
Assignee: ${assigneeName}
Due Date: ${task.Due_Date || 'Not specified'}
Context: ${task.Context_Hidden || task.Interaction_Log || 'No additional context'}

Write an email that:
1. Is polite but authoritative
2. Clearly states the task and expectations
3. Asks for confirmation of the deadline
4. Invites questions or concerns
5. Uses a professional but friendly tone

Return ONLY the email body text (no subject line, no signature - those will be added separately).`;

  try {
    const emailBody = callGeminiPro(prompt, { temperature: 0.7 });
    return emailBody.trim();
  } catch (error) {
    logError(ERROR_TYPE.API_ERROR, 'generateAssignmentEmail', error.toString(), task.Task_ID);
    // Fallback to template
    return getEmailTemplate('assignment', task);
  }
}

/**
 * Classify email reply type
 */
function classifyReplyType(replyContent) {
  const prompt = `Analyze this email reply and classify it into one of these categories:
1. ACCEPTANCE - The assignee accepts the task and deadline
2. DATE_CHANGE - The assignee requests a different due date
3. SCOPE_QUESTION - The assignee has questions about scope or needs clarification
4. ROLE_REJECTION - The assignee says this is not their responsibility
5. OTHER - Something else

Email reply: "${replyContent}"

Return ONLY a JSON object:
{
  "type": "one of the categories above",
  "confidence": 0.0-1.0,
  "extracted_date": "proposed date in YYYY-MM-DD format if DATE_CHANGE, or null",
  "reasoning": "brief explanation"
}`;

  try {
    const response = callGeminiPro(prompt, { temperature: 0.3 });
    let jsonText = response.trim();
    if (jsonText.startsWith('```')) {
      jsonText = jsonText.replace(/^```json\n?/, '').replace(/```$/, '');
    }
    return JSON.parse(jsonText);
  } catch (error) {
    logError(ERROR_TYPE.API_ERROR, 'classifyReplyType', error.toString());
    return {
      type: 'OTHER',
      confidence: 0.5,
      extracted_date: null,
      reasoning: 'Failed to classify reply'
    };
  }
}

/**
 * Analyze meeting transcript and extract structured data
 */
function analyzeMeetingTranscript(transcript) {
  const prompt = `Analyze this meeting transcript and extract structured information.

Meeting Transcript:
"${transcript}"

Extract and return ONLY a valid JSON object:
{
  "executive_summary": "brief overview of meeting purpose and conclusions",
  "decisions_made": ["decision 1", "decision 2", ...],
  "action_items": [
    {
      "description": "action item description",
      "owner": "person responsible (name or email if mentioned)",
      "deadline": "deadline if mentioned (YYYY-MM-DD or null)"
    }
  ],
  "risks_sentiment": "notable risks or general sentiment/tone of meeting"
}`;

  try {
    const response = callGeminiPro(prompt, { 
      temperature: 0.7,
      maxOutputTokens: 8192 
    });
    
    let jsonText = response.trim();
    if (jsonText.startsWith('```')) {
      jsonText = jsonText.replace(/^```json\n?/, '').replace(/```$/, '');
    }
    
    return JSON.parse(jsonText);
  } catch (error) {
    logError(ERROR_TYPE.API_ERROR, 'analyzeMeetingTranscript', error.toString());
    throw error;
  }
}

