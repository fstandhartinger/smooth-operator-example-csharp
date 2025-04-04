using DotNetEnv;
using OpenAI;
using OpenAI.Chat;
using SmoothOperator.AgentTools;

// Load environment variables from .env file
Env.TraversePath().Load();

var screengraspApiKey = Environment.GetEnvironmentVariable("SCREENGRASP_API_KEY");
var openaiApiKey = Environment.GetEnvironmentVariable("OPENAI_API_KEY");

if (string.IsNullOrEmpty(screengraspApiKey))
{
    Console.WriteLine("Error: SCREENGRASP_API_KEY not found in .env file or environment variables. Get a free key at https://screengrasp.com/api.html");
    Console.WriteLine("Press any key to exit.");
    Console.ReadKey();
    return; // Exit if ScreenGrasp key is missing
}

if (string.IsNullOrEmpty(openaiApiKey))
{
    Console.WriteLine("Warning: OPENAI_API_KEY not found in .env file or environment variables. OpenAI part will be skipped. Get a key at https://platform.openai.com/api-keys");
}

// Initialize the Smooth Operator Client
// The API key is required for features like ClickByDescriptionAsync (ScreenGrasp)
using var client = new SmoothOperatorClient(screengraspApiKey);

Console.WriteLine("Starting server (can take a while, especially on first run, because it's installing the server)...");
// StartServerAsync ensures the Smooth Operator server process is running in the background.
// It handles the download and extraction of the server if it's not installed or outdated.
await client.StartServerAsync();

Console.WriteLine("Opening calculator...");
// Open the Windows Calculator application.
await client.System.OpenApplicationAsync("calc");

Console.WriteLine("Typing 3+4...");
// Type the string "3+4" into the currently focused window (hopefully the calculator).
await client.Keyboard.TypeAsync("3+4");

Console.WriteLine("Clicking equals sign...");
// Click the UI element described as "the equals sign" using ScreenGrasp.
// Alternatives:
// - await client.Keyboard.TypeAsync("="); // Simpler if the equals key works
// - Using Windows UI Automation method client.Automation.InvokeAsync() (a bit more complex to implement but more robust)
await client.Mouse.ClickByDescriptionAsync("the equals sign");

if (!string.IsNullOrEmpty(openaiApiKey))
{
    Console.WriteLine("Getting window overview...");
    // Get an overview of the current system state, including the focused window's automation tree.
    // Assumes the calculator is still the focused window. Be mindful of focus changes during debugging.
    var overview = await client.System.GetOverviewAsync();

    if (overview?.FocusInfo?.FocusedElementParentWindow != null)
    {
        // Convert the automation tree of the focused window to a JSON string.
        var focusedWindowJson = overview.FocusInfo.FocusedElementParentWindow.ToJsonString();

        // Use GPT-4o (or another AI model) to interpret the automation tree and extract the result.
        // This demonstrates integrating AI for understanding application state.
        Console.WriteLine("Asking OpenAI about the result...");
        try
        {
            var openAIClient = new OpenAIClient(openaiApiKey);
            var chatClient = openAIClient.GetChatClient("chatgpt-4o-latest"); // Use a suitable model
            var result = await chatClient.CompleteChatAsync(
                new ChatMessage[] {
                    ChatMessage.CreateUserMessage($"What result does the calculator display? You can read it from its automation tree: {focusedWindowJson}")
                });

            Console.WriteLine("OpenAI Result: " + result.Value.Content[0].Text);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error calling OpenAI: {ex.Message}");
        }
    }
    else
    {
        Console.WriteLine("Could not get focused window information to send to OpenAI.");
    }
}
else
{
    Console.WriteLine("OpenAI key not provided, skipping result verification.");
}


/*
 * --- Alternative using Screenshot ---
 * Taking a screenshot and analyzing it with AI is another option,
 * but generally less reliable and more costly than using the automation tree.
 * Prefer Automation Tree > Keyboard > Screenshot for robustness and cost-efficiency.

if (!string.IsNullOrEmpty(openaiApiKey))
{
    Console.WriteLine("Taking screenshot...");
    var screenshot = await client.Screenshot.TakeAsync();
    Console.WriteLine("Asking OpenAI about the screenshot...");
     try
     {
        var openAIClient = new OpenAIClient(openaiApiKey);
        var chatClient = openAIClient.GetChatClient("gpt-4o-latest"); // Or "gpt-4-vision-preview" if needed
        var result = await chatClient.CompleteChatAsync(
            new ChatMessage[] {
                 ChatMessage.CreateUserMessage(new ChatMessageContentPart[] {
                     ChatMessageContentPart.CreateTextPart("What result does the calculator display based on the screenshot?"),
                     ChatMessageContentPart.CreateImagePart(BinaryData.FromBytes(screenshot.ImageBytes), "image/jpeg") // Send image data
                 })
            });
        Console.WriteLine("OpenAI Screenshot Result: "+ result.Value.Content[0].Text);
     }
     catch (Exception ex)
     {
         Console.WriteLine($"Error calling OpenAI with screenshot: {ex.Message}");
     }
}
*/

// The SmoothOperatorClient is IDisposable and will handle stopping the server
// connection when the 'using' block ends if StartServerAsync was called.
// Explicitly calling client.StopServer() generally isn't needed here.

Console.WriteLine("Example finished. Press any key to exit.");
Console.ReadKey();