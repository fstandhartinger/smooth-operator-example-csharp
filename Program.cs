using DotNetEnv;
using OpenAI;
using OpenAI.Chat;
using SmoothOperator.AgentTools;


// If you want to run another example that summarizes data from some twitter feeds, comment out the following two lines
// await new TwitterAiNewsChecker().Run();
// return;

// If you want to run the example that collects order data from email into a mock ERP, uncomment the following two lines
// await new CollectOrdersFromEmailIntoErpSystem().Run();
// return;

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
// - await client.Keyboard.TypeAsync("="); // Simpler and faster
// - Using Windows UI Automation, e.g. client.Automation.InvokeAsync() a bit more complex to implement but very robust (not affected by focus changes)
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

        // You can use GPT-4o or other AI models for all sorts of tasks together with the Smooth Operator Agent Tools.
        // In this case we use it to read the result of the calculator from its automation tree.
        // But it can also for example be used to decide which button to click next, what text to type, etc.
        Console.WriteLine("Asking OpenAI about the result...");
        try
        {
            var openAIClient = new OpenAIClient(openaiApiKey);
            var chatClient = openAIClient.GetChatClient("gpt-4o"); // Use a suitable model
            var result = await chatClient.CompleteChatAsync(ChatMessage.CreateUserMessage($"What result does the calculator display? You can read it from its automation tree: {focusedWindowJson}"));  

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
        var chatClient = openAIClient.GetChatClient("gpt-4o");
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