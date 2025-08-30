#!/bin/bash

echo "?? GeminiSheetsChat Setup Script"
echo "================================"

# Check if .NET 9 is installed
if ! command -v dotnet &> /dev/null; then
    echo "? .NET is not installed. Please install .NET 9 SDK first."
    echo "   Download from: https://dotnet.microsoft.com/download/dotnet/9.0"
    exit 1
fi

echo "? .NET found"

# Initialize user secrets if not already done
echo "?? Initializing user secrets..."
dotnet user-secrets init

# Prompt for Gemini API key
echo ""
echo "?? Setting up Gemini API Key"
echo "Please get your API key from: https://aistudio.google.com/"
echo ""
read -p "Enter your Gemini API key: " GEMINI_KEY

if [ -z "$GEMINI_KEY" ]; then
    echo "??  No API key provided. You can set it later with:"
    echo "   dotnet user-secrets set \"GEMINI_API_KEY\" \"your-key-here\""
else
    dotnet user-secrets set "GEMINI_API_KEY" "$GEMINI_KEY"
    echo "? Gemini API key saved"
fi

# Prompt for Google Spreadsheet ID
echo ""
echo "?? Setting up Google Spreadsheet ID"
echo "Please create a Google Sheet and copy its ID from the URL"
echo "Example: https://docs.google.com/spreadsheets/d/[SPREADSHEET_ID]/edit"
echo ""
read -p "Enter your Google Spreadsheet ID: " SPREADSHEET_ID

if [ -z "$SPREADSHEET_ID" ]; then
    echo "??  No Spreadsheet ID provided. You can set it later with:"
    echo "   dotnet user-secrets set \"GOOGLE_SPREADSHEET_ID\" \"your-spreadsheet-id\""
else
    dotnet user-secrets set "GOOGLE_SPREADSHEET_ID" "$SPREADSHEET_ID"
    echo "? Google Spreadsheet ID saved"
fi

echo ""
echo "?? Google Sheets Setup Checklist:"
echo "1. Go to Google Cloud Console (https://console.cloud.google.com/)"
echo "2. Create a new project or select existing one"
echo "3. Enable the Google Sheets API"
echo "4. Create a Service Account and download credentials JSON"
echo "5. Rename the file to 'geminisheetschat.json' and place in this directory"
echo "6. Create a Google Sheet and copy its ID from the URL"
echo "7. Share the sheet with your service account email (found in JSON file)"
echo "8. Grant 'Editor' permissions to the service account"

echo ""
echo "?? Files needed in this directory:"
echo "   - geminisheetschat.json (renamed from Google Cloud Console download)"
echo ""
echo "?? Configuration is now stored securely in user secrets:"
echo "   - GEMINI_API_KEY: Your Gemini AI API key"
echo "   - GOOGLE_SPREADSHEET_ID: Your Google Sheets document ID"

echo ""
echo "?? Setup complete! Run 'dotnet run' to start the application."
echo ""
echo "For detailed instructions, see README.md"