#!/bin/bash
# Sync .gs files from apps-script/ to .js files in apps-script-repo/
# Run this before clasp push if you edit .gs files

echo "Syncing .gs files to .js files..."

# Copy and rename files
cp ../apps-script/GeminiAI.gs ./GeminiAI.js
cp ../apps-script/SheetsHelper.gs ./SheetsHelper.js
cp ../apps-script/VoiceIngestion.gs ./VoiceIngestion.js
cp ../apps-script/EmailNegotiation.gs ./EmailNegotiation.js
cp ../apps-script/DashboardActions.gs ./DashboardActions.js
cp ../apps-script/Code.gs ./Code.js
cp ../apps-script/MeetingIntelligence.gs ./MeetingIntelligence.js
cp ../apps-script/CalendarHelper.gs ./CalendarHelper.js
cp ../apps-script/EmailHelper.gs ./EmailHelper.js
cp ../apps-script/Config.gs ./Config.js

echo "✅ Files synced! Now run: clasp push"

