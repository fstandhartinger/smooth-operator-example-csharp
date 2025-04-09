using DotNetEnv;
using OpenAI;
using OpenAI.Chat;
using SmoothOperator.AgentTools;

public class TwitterAiNewsChecker
{
    public async Task Run()
    {
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

        Console.WriteLine("Opening browser...");
        var isBrowserOpen = false;

        var accounts = new[] { "sama", "DataChaz", "mattshumer_", "karpathy", "kimmonismus", "ai_for_success", "slow_developer" };

        var tweetsText = "";

        foreach (var account in accounts)
        {
            var url = $"https://x.com/{account}";

            if (!isBrowserOpen)
            {
                // Open the Windows Calculator application.
                var openChromeResult = await client.Chrome.OpenChromeAsync(url);
                if (openChromeResult.Message.StartsWith("Error"))
                {                    
                    Console.WriteLine("Error: Could not open Chrome. Use ForceClose parameter or ensure there is no open instance before running this example.");
                    return; // Exit if Chrome cannot be opened
                }
                isBrowserOpen = true;
                await Task.Delay(7000); // give the newly opened browser some time to load that page
            }
            else
            {
                await client.Chrome.NavigateAsync(url);                
            }

            // scroll down the timeline a couple of times
            for (var i = 0; i < 3; i++) {
                await client.Mouse.ScrollAsync(200, 200, 20);
                await Task.Delay(1000);
            }

            Console.WriteLine($"Getting text from {url}...");
            var response = await client.Chrome.GetTextAsync();
            tweetsText += response.ResultValue + Environment.NewLine;
            tweetsText += "--------------------" + Environment.NewLine; // separator
            await Task.Delay(1000); // Small delay between accounts
        }

        if (string.IsNullOrWhiteSpace(tweetsText))
        {
             Console.WriteLine("Error: Could not retrieve any tweet text. Skipping OpenAI analysis.");
        }
        else if (string.IsNullOrEmpty(openaiApiKey))
        {
             Console.WriteLine("Skipping OpenAI analysis as API key is missing.");
        }
        else
        {
            Console.WriteLine("Asking OpenAI about the result...");
            try
            {
                var openAiClient = new OpenAIClient(openaiApiKey);
                var chatClient = openAiClient.GetChatClient("gpt-4o"); // Use a suitable model
                var options = new ChatCompletionOptions { ResponseFormat = ChatResponseFormat.CreateJsonObjectFormat() };
                var result = await chatClient.CompleteChatAsync([
                        ChatMessage.CreateUserMessage(
                            "These are the latest tweets of some twitter accounts that are typically very up-to-date on AI news. " +
                            "Give me a summary on the concrete topics they write about (5 bullet points, one short sentence, each) and a rating 0-100 if you have the impression that actual very big breaking news has just occurred within the last hour." +
                            "<tweets>" +
                            $"{tweetsText}" +
                            "</tweets>" +
                            """
                            Answer with a JSON in this form: 
                            { 
                                "summaryBulletPoints": 
                                    [
                                        "bullet point 1", 
                                        "bullet point 2", 
                                        "bullet point 3"
                                    ],
                                "breakingNewsProbabilityInPercent": 50 
                            } 
                            """)
                        ],
                        options);

                Console.WriteLine();
                Console.WriteLine("Result: " + Environment.NewLine + result.Value.Content[0].Text);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error calling OpenAI: {ex.Message}");
            }
        }

        Console.WriteLine();
        Console.WriteLine("Press any key to exit.");
        Console.ReadKey();
    }
}