# Changelog

All notable changes to the AI COS Apps Script project.

## [v1.0-voice-processing] - 2025-12-28

### 🎯 Milestone: Voice Processing with Gemini AI

First stable release with complete voice-to-task pipeline.

### Features

#### Voice Processing
- **Gemini AI Transcription**: Convert voice notes to text with high accuracy
- **Smart Name Matching**: Phonetic matching for Indian/South Asian names
  - Handles variations like "Anaaya" vs "Anaya" vs "अनाया"
  - Generates phonetic hints for common names
- **Hindi/Hinglish Support**: Understands date expressions like "kal", "parson", "agle hafte"

#### Date Handling
- **dd-MM-yyyy Format**: Consistent date format throughout (Indian format)
- **Relative Date Parsing**: "tomorrow", "next Monday", "end of week"
- **Date Context in Prompts**: AI knows current date for accurate parsing

#### Database Integration
- **Staff_DB**: Auto-populate assignee name from email lookup
- **Projects_DB**: Match projects by name or tag
- **Tasks_DB**: Store voice transcripts, confidence scores, tone detection

#### Email Workflow
- **Assignment Emails**: Auto-send when task is assigned
- **Negotiation Handling**: Date change requests, scope questions
- **Reply Processing**: Parse employee responses

#### Dashboard API
- **CRUD Operations**: Create, read, update, delete tasks
- **Action Handlers**: Approve, reject, negotiate, reassign
- **Staff/Project Endpoints**: List and manage staff and projects

### Files
- `VoiceIngestion.js` - Voice note processing pipeline
- `GeminiAI.js` - AI transcription and parsing
- `SheetsHelper.js` - Google Sheets CRUD operations
- `DashboardActions.js` - API action handlers
- `EmailHelper.js` - Email sending utilities
- `EmailNegotiation.js` - Reply parsing and negotiation
- `WorkflowEngine.js` - Trigger-based workflows
- `Config.js` - Configuration management
- `Code.js` - Main entry points and setup

---

## How to Return to This Version

```bash
# View this version
git checkout v1.0-voice-processing

# Return to latest
git checkout main
```


