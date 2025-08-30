@echo off
echo ?? GeminiSheetsChat Setup Script
echo ================================

REM Check if .NET 9 is installed
dotnet --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ? .NET is not installed. Please install .NET 9 SDK first.
    echo    Download from: https://dotnet.microsoft.com/download/dotnet/9.0
    pause
    exit /b 1
)

echo ? .NET found

REM Initialize user secrets if not already done
echo ?? Initializing user secrets...
dotnet user-secrets init

REM Prompt for Gemini API key
echo.
echo ?? Setting up Gemini API Key
echo Please get your API key from: https://aistudio.google.com/
echo.
set /p GEMINI_KEY="Enter your Gemini API key: "

if "%GEMINI_KEY%"=="" (
    echo ??  No API key provided. You can set it later with:
    echo    dotnet user-secrets set "GEMINI_API_KEY" "your-key-here"
) else (
    dotnet user-secrets set "GEMINI_API_KEY" "%GEMINI_KEY%"
    echo ? Gemini API key saved
)

REM Prompt for Google Spreadsheet ID
echo.
echo ?? Setting up Google Spreadsheet ID
echo Please create a Google Sheet and copy its ID from the URL
echo Example: https://docs.google.com/spreadsheets/d/[SPREADSHEET_ID]/edit
echo.
set /p SPREADSHEET_ID="Enter your Google Spreadsheet ID: "

if "%SPREADSHEET_ID%"=="" (
    echo ??  No Spreadsheet ID provided. You can set it later with:
    echo    dotnet user-secrets set "GOOGLE_SPREADSHEET_ID" "your-spreadsheet-id"
) else (
    dotnet user-secrets set "GOOGLE_SPREADSHEET_ID" "%SPREADSHEET_ID%"
    echo ? Google Spreadsheet ID saved
)

echo.
echo ?? Google Sheets Setup Checklist:
echo 1. Go to Google Cloud Console ^(https://console.cloud.google.com/^)
echo 2. Create a new project or select existing one
echo 3. Enable the Google Sheets API
echo 4. Create a Service Account and download credentials JSON
echo 5. Rename the file to 'geminisheetschat.json' and place in this directory
echo 6. Create a Google Sheet and copy its ID from the URL
echo 7. Share the sheet with your service account email ^(found in JSON file^)
echo 8. Grant 'Editor' permissions to the service account

echo.
echo ?? Files needed in this directory:
echo    - geminisheetschat.json ^(renamed from Google Cloud Console download^)
echo.
echo ?? Configuration is now stored securely in user secrets:
echo    - GEMINI_API_KEY: Your Gemini AI API key
echo    - GOOGLE_SPREADSHEET_ID: Your Google Sheets document ID

echo.
echo ?? Setup complete! Run 'dotnet run' to start the application.
echo.
echo For detailed instructions, see README.md

pause