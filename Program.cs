using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Microsoft.Extensions.Configuration;
using Microsoft.SemanticKernel;
using System.Text;

// Configuration setup for API keys
var config = new ConfigurationBuilder()
    .AddUserSecrets<Program>()
    .AddEnvironmentVariables()
    .Build();

var kernel = BuildKernel(config);
var sheetsService = SetupGoogleSheetsService();

// Configuration - Get spreadsheet ID from configuration
string? spreadsheetId = config["GOOGLE_SPREADSHEET_ID"];
if (string.IsNullOrEmpty(spreadsheetId))
{
    Console.WriteLine("❌ GOOGLE_SPREADSHEET_ID is not set.");
    Console.WriteLine("Please set it using: dotnet user-secrets set \"GOOGLE_SPREADSHEET_ID\" \"your-spreadsheet-id\"");
    Console.WriteLine("Or set it as an environment variable.");
    return;
}

string sheetName = "Employees"; // Updated to match the actual sheet name

Console.WriteLine("Setting up Google Sheets...");
string actualSheetName = SetupGoogleSheet(sheetsService, spreadsheetId, sheetName);
string schema = GetSheetSchema();

Console.WriteLine("Google Sheets Schema:");
Console.WriteLine(schema);
Console.WriteLine("\nChat with your Google Sheet! Type 'exit' to quit.");
Console.WriteLine("Make sure you've set your GEMINI_API_KEY and configured your Google Sheets credentials.");

var textToQueryFunction = CreateTextToQueryFunction(kernel, schema);
var finalAnswerFunction = CreateFinalAnswerFunction(kernel);

while (true)
{
    Console.Write("> ");
    string userInput = Console.ReadLine() ?? "";
    if (userInput.Equals("exit", StringComparison.OrdinalIgnoreCase))
    {
        break;
    }

    try
    {
        var queryResult = await textToQueryFunction.InvokeAsync(kernel, new() { ["input"] = userInput });
        string queryDescription = queryResult.GetValue<string>()!.Trim();
        Console.WriteLine($"\n?? Query Logic: {queryDescription}");

        string sheetData = await ExecuteSheetQueryAndFormatResults(sheetsService, spreadsheetId, actualSheetName, userInput, queryDescription);
        if (string.IsNullOrWhiteSpace(sheetData))
        {
            Console.WriteLine("?? I couldn't find any data for that query.\n");
            continue;
        }
        Console.WriteLine($"?? Sheet Result:\n{sheetData}");

        var finalAnswerResult = await finalAnswerFunction.InvokeAsync(kernel, new()
        {
            ["input"] = userInput,
            ["data"] = sheetData
        });

        Console.WriteLine($"\n?? Answer: {finalAnswerResult.GetValue<string>()}\n");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"\nAn error occurred: {ex.Message}\n");
        if (ex.Message.Contains("credentials"))
        {
            Console.WriteLine("Make sure you have:");
            Console.WriteLine("1. Set GEMINI_API_KEY: dotnet user-secrets set \"GEMINI_API_KEY\" \"your-key\"");
            Console.WriteLine("2. Set GOOGLE_SPREADSHEET_ID: dotnet user-secrets set \"GOOGLE_SPREADSHEET_ID\" \"your-id\"");
            Console.WriteLine("3. Created geminisheetschat.json file from Google Cloud Console");
            Console.WriteLine("4. Shared your Google Sheet with the service account email");
        }
    }
}

#pragma warning disable SKEXP0070

Kernel BuildKernel(IConfiguration config)
{
    var builder = Kernel.CreateBuilder();

    string? apiKey = config["GEMINI_API_KEY"];
    if (string.IsNullOrEmpty(apiKey))
    {
        throw new InvalidOperationException("GEMINI_API_KEY is not set. Please set it using 'dotnet user-secrets set GEMINI_API_KEY your-key' or as an environment variable.");
    }

    builder.AddGoogleAIGeminiChatCompletion(
        modelId: "gemini-1.5-flash", // You can also use "gemini-1.5-pro" for more advanced capabilities
        apiKey: apiKey
    );

    return builder.Build();
}

SheetsService SetupGoogleSheetsService()
{
    string[] scopes = { SheetsService.Scope.Spreadsheets };
    
    GoogleCredential credential;
    
    // Try to find credentials.json in current directory or parent directory
    string credentialsPath = "../../../../geminisheetschat.json";
    if (!File.Exists(credentialsPath))
    {
        credentialsPath = "../../../geminisheetschat.json";
        if (!File.Exists(credentialsPath))
        {
            throw new FileNotFoundException("credentials.json not found. Please download it from Google Cloud Console and place it in the project directory.");
        }
    }
    
    using (var stream = new FileStream(credentialsPath, FileMode.Open, FileAccess.Read))
    {
        credential = GoogleCredential.FromStream(stream).CreateScoped(scopes);
    }

    var service = new SheetsService(new BaseClientService.Initializer()
    {
        HttpClientInitializer = credential,
        ApplicationName = "GeminiSheetsChat",
    });

    return service;
}

KernelFunction CreateTextToQueryFunction(Kernel kernel, string schema)
{
    const string prompt = @"
Given the following Google Sheets data structure, analyze the user's question and describe what data should be filtered or retrieved.
- Describe the filtering logic clearly
- Mention specific column names and conditions
- Be specific about what data should be returned
- If it's a calculation (like average), mention that

Schema:
---
{{$schema}}
---

User Question: {{$input}}

Query Description:
";

    var executionSettings = new PromptExecutionSettings()
    {
        ExtensionData = new Dictionary<string, object>()
        {
            { "temperature", 0.0 }
        }
    };

    var promptConfig = new PromptTemplateConfig
    {
        Template = prompt.Replace("{{$schema}}", schema),
        ExecutionSettings = new Dictionary<string, PromptExecutionSettings>
        {
            { "default", executionSettings }
        }
    };

    return KernelFunctionFactory.CreateFromPrompt(promptConfig);
}

KernelFunction CreateFinalAnswerFunction(Kernel kernel)
{
    const string prompt = @"
Answer the following user's question based ONLY on the provided data from Google Sheets.
If the data is empty or irrelevant, say you could not find an answer.
Be friendly and concise.

Data:
---
{{$data}}
---

User Question: {{$input}}

Answer:
";

    var executionSettings = new PromptExecutionSettings()
    {
        ExtensionData = new Dictionary<string, object>()
        {
            { "temperature", 0.2 }
        }
    };

    var promptConfig = new PromptTemplateConfig
    {
        Template = prompt,
        ExecutionSettings = new Dictionary<string, PromptExecutionSettings>
        {
            { "default", executionSettings }
        }
    };

    return KernelFunctionFactory.CreateFromPrompt(promptConfig);
}

async Task<string> ExecuteSheetQueryAndFormatResults(SheetsService service, string spreadsheetId, string sheetName, string userQuery, string queryDescription)
{
    var range = $"{sheetName}!A:E"; // Adjust range based on your data
    var request = service.Spreadsheets.Values.Get(spreadsheetId, range);
    var response = await request.ExecuteAsync();
    
    var values = response.Values;
    if (values == null || values.Count == 0)
    {
        return string.Empty;
    }

    // Filter data based on user query and AI description
    var filteredData = FilterDataBasedOnQuery(values, userQuery, queryDescription);
    
    var resultBuilder = new StringBuilder();
    
    // Add headers
    if (filteredData.Count > 0)
    {
        resultBuilder.AppendLine(string.Join("\t", filteredData[0]));
        
        // Add data rows
        for (int i = 1; i < filteredData.Count; i++)
        {
            resultBuilder.AppendLine(string.Join("\t", filteredData[i]));
        }
    }

    return resultBuilder.ToString();
}

List<IList<object>> FilterDataBasedOnQuery(IList<IList<object>> data, string userQuery, string queryDescription)
{
    var result = new List<IList<object>>();
    if (data.Count == 0) return result;
    
    // Always include headers
    result.Add(data[0]);
    
    var query = userQuery.ToLower();
    var description = queryDescription.ToLower();
    
    // Enhanced filtering logic based on user query and AI description
    for (int i = 1; i < data.Count; i++)
    {
        var row = data[i];
        bool includeRow = false;
        
        // Department-based filtering
        if (query.Contains("engineer") || description.Contains("engineering"))
        {
            includeRow = row.Count > 2 && row[2]?.ToString()?.ToLower().Contains("engineering") == true;
        }
        else if (query.Contains("sales") || description.Contains("sales"))
        {
            includeRow = row.Count > 2 && row[2]?.ToString()?.ToLower().Contains("sales") == true;
        }
        else if (query.Contains("hr") || description.Contains("hr"))
        {
            includeRow = row.Count > 2 && row[2]?.ToString()?.ToLower().Contains("hr") == true;
        }
        // Salary-based filtering
        else if (query.Contains("salary") && (query.Contains("more than") || query.Contains(">") || description.Contains("greater")))
        {
            if (row.Count > 3 && int.TryParse(row[3]?.ToString(), out int salary))
            {
                // Extract number from query
                var numbers = System.Text.RegularExpressions.Regex.Matches(query, @"\d+");
                if (numbers.Count > 0 && int.TryParse(numbers[0].Value, out int threshold))
                {
                    includeRow = salary > threshold;
                }
                else
                {
                    includeRow = salary > 90000; // Default threshold
                }
            }
        }
        else if (query.Contains("salary") && (query.Contains("less than") || query.Contains("<") || description.Contains("less")))
        {
            if (row.Count > 3 && int.TryParse(row[3]?.ToString(), out int salary))
            {
                var numbers = System.Text.RegularExpressions.Regex.Matches(query, @"\d+");
                if (numbers.Count > 0 && int.TryParse(numbers[0].Value, out int threshold))
                {
                    includeRow = salary < threshold;
                }
                else
                {
                    includeRow = salary < 80000; // Default threshold
                }
            }
        }
        // Average calculations - include relevant rows
        else if ((query.Contains("average") || description.Contains("average")) && query.Contains("salary"))
        {
            if (query.Contains("sales") || description.Contains("sales"))
            {
                includeRow = row.Count > 2 && row[2]?.ToString()?.ToLower().Contains("sales") == true;
            }
            else if (query.Contains("engineering") || description.Contains("engineering"))
            {
                includeRow = row.Count > 2 && row[2]?.ToString()?.ToLower().Contains("engineering") == true;
            }
            else
            {
                includeRow = true; // Include all for overall average
            }
        }
        // Name-based searches
        else if (query.Contains("who") || description.Contains("names") || description.Contains("employees"))
        {
            includeRow = true; // Show all employees for "who" questions
        }
        // Date-based filtering
        else if (query.Contains("hired") || query.Contains("hire date") || description.Contains("date"))
        {
            if (query.Contains("2022") || description.Contains("2022"))
            {
                includeRow = row.Count > 4 && row[4]?.ToString()?.Contains("2022") == true;
            }
            else if (query.Contains("2023") || description.Contains("2023"))
            {
                includeRow = row.Count > 4 && row[4]?.ToString()?.Contains("2023") == true;
            }
            else
            {
                includeRow = true; // Show all dates
            }
        }
        else
        {
            includeRow = true; // Default: include all rows for general queries
        }
        
        if (includeRow)
        {
            result.Add(row);
        }
    }
    
    return result;
}

string SetupGoogleSheet(SheetsService service, string spreadsheetId, string sheetName)
{
    try
    {
        var spreadsheet = service.Spreadsheets.Get(spreadsheetId).Execute();
        
        // Check if the target sheet exists
        var targetSheet = spreadsheet.Sheets?.FirstOrDefault(s => s.Properties.Title == sheetName);
        string actualSheetName = sheetName;
        
        if (targetSheet == null)
        {
            // Sheet doesn't exist, try to create it
            try
            {
                var addSheetRequest = new AddSheetRequest
                {
                    Properties = new SheetProperties
                    {
                        Title = sheetName
                    }
                };

                var batchUpdateRequest = new BatchUpdateSpreadsheetRequest
                {
                    Requests = new List<Request>
                    {
                        new Request { AddSheet = addSheetRequest }
                    }
                };

                service.Spreadsheets.BatchUpdate(batchUpdateRequest, spreadsheetId).Execute();
                Console.WriteLine($"✓ Created new sheet: {sheetName}");
            }
            catch (Exception createEx)
            {
                // If we can't create the sheet, use the first available sheet
                Console.WriteLine($"⚠ Could not create sheet '{sheetName}': {createEx.Message}");
                var firstSheet = spreadsheet.Sheets?.FirstOrDefault();
                if (firstSheet != null)
                {
                    actualSheetName = firstSheet.Properties.Title;
                    Console.WriteLine($"ℹ Using existing sheet: {actualSheetName}");
                }
                else
                {
                    throw new Exception("No sheets found in the spreadsheet");
                }
            }
        }
        else
        {
            Console.WriteLine($"ℹ Using existing sheet: {actualSheetName}");
        }
        
        // Clear existing data in the range
        var clearRequest = new ClearValuesRequest();
        service.Spreadsheets.Values.Clear(clearRequest, spreadsheetId, $"{actualSheetName}!A:Z").Execute();

        // Prepare sample data
        var values = new List<IList<object>>
        {
            new List<object> { "Id", "Name", "Department", "Salary", "HireDate" },
            new List<object> { "1", "Alice Johnson", "Engineering", "95000", "2022-01-15" },
            new List<object> { "2", "Bob Smith", "Sales", "82000", "2021-11-30" },
            new List<object> { "3", "Charlie Brown", "Engineering", "110000", "2020-05-20" },
            new List<object> { "4", "Diana Prince", "Sales", "78000", "2022-08-01" },
            new List<object> { "5", "Eve Adams", "HR", "65000", "2023-02-10" }
        };

        // Update the sheet with sample data
        var updateRequest = new ValueRange { Values = values };
        var update = service.Spreadsheets.Values.Update(updateRequest, spreadsheetId, $"{actualSheetName}!A1");
        update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
        update.Execute();

        Console.WriteLine($"✓ Google Sheet '{actualSheetName}' populated with sample data.");
        
        // Return the actual sheet name that was used
        if (actualSheetName != sheetName)
        {
            Console.WriteLine($"ℹ Note: Using sheet name '{actualSheetName}' instead of '{sheetName}'");
        }
        
        return actualSheetName;
    }
    catch (Exception ex)
    {
        Console.WriteLine($"❌ Error setting up Google Sheet: {ex.Message}");
        Console.WriteLine("\nSetup checklist:");
        Console.WriteLine("1. Create a Google Sheet and get its ID from the URL");
        Console.WriteLine("2. Share the sheet with your service account email");
        Console.WriteLine("3. Set GOOGLE_SPREADSHEET_ID: dotnet user-secrets set \"GOOGLE_SPREADSHEET_ID\" \"your-id\"");
        Console.WriteLine("4. Make sure geminisheetschat.json is in the project directory");
        
        // Return the original sheet name as fallback
        return sheetName;
    }
}

static string GetSheetSchema()
{
    return @"
Google Sheet Structure:
Sheet Name: Employees
Columns:
- A: Id (INTEGER) - Employee ID number
- B: Name (TEXT) - Employee full name  
- C: Department (TEXT) - Department (Engineering, Sales, HR)
- D: Salary (INTEGER) - Annual salary in USD
- E: HireDate (TEXT) - Date hired (YYYY-MM-DD format)

Sample Data Available:
- Departments: Engineering, Sales, HR
- Salary ranges from $65,000 to $110,000
- Hire dates from 2020 to 2023
- 5 employees total in the dataset

Supported Query Types:
- Filter by department (e.g., 'Who are the engineers?')
- Filter by salary (e.g., 'Who earns more than 90000?')
- Calculate averages (e.g., 'What is the average salary in Sales?')
- Filter by hire date (e.g., 'Who was hired in 2022?')
- General employee information queries
";
}