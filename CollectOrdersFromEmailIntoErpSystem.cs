using DotNetEnv;
using OpenAI;
using OpenAI.Chat;
using SmoothOperator.AgentTools;
using SmoothOperator.AgentTools.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

public class CollectOrdersFromEmailIntoErpSystem
{
    public async Task Run()
    {
        Console.WriteLine("Starting Email-to-ERP Example...");

        // Load environment variables
        Env.TraversePath().Load();
        var screengraspApiKey = Environment.GetEnvironmentVariable("SCREENGRASP_API_KEY");
        var openaiApiKey = Environment.GetEnvironmentVariable("OPENAI_API_KEY");

        if (string.IsNullOrEmpty(screengraspApiKey))
        {
            Console.WriteLine("Error: SCREENGRASP_API_KEY not found.");
            return;
        }
        if (string.IsNullOrEmpty(openaiApiKey))
        {
            Console.WriteLine("Warning: OPENAI_API_KEY not found. AI steps will be skipped.");
        }

        // Initialize SmoothOperator Client
        using var client = new SmoothOperatorClient(screengraspApiKey);
        Console.WriteLine("Starting Smooth Operator server...");
        await client.StartServerAsync();

        ScreenshotResponse emailScreenshot = null;
        try
        {
            // --- Get Order Email Screenshot ---
            // By default, uses Gmail via Chrome.
            // To use local Outlook instead (if installed), comment the Gmail line and uncomment the Outlook line below.
            Console.WriteLine("Attempting to get order email screenshot via Gmail...");
            emailScreenshot = await GetOrderScreenshotFromGmailAsync(client); //with a little extra effort we could instead just get the order email text as well
            // Console.WriteLine("Attempting to get order email screenshot via Outlook...");
            //emailScreenshot = await GetOrderScreenshotFromOutlookAsync(client);

            if (emailScreenshot == null || !emailScreenshot.Success)
            {
                Console.WriteLine("Error: Could not get email screenshot.");
                return;
            }
            Console.WriteLine("Successfully captured email screenshot.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error getting email screenshot: {ex.Message}");
            return;
        }

        // --- Download and Run Mock ERP ---
        string erpExePath = null;
        try
        {
            Console.WriteLine("Downloading mock ERP application...");
            erpExePath = await DownloadMockErpAsync();
            Console.WriteLine($"Mock ERP downloaded to: {erpExePath}");

            Console.WriteLine("Launching mock ERP application...");
            await client.System.OpenApplicationAsync(erpExePath);
            await Task.Delay(5000); // Give the app time to start
            Console.WriteLine("Mock ERP application launched.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error with mock ERP application: {ex.Message}");
            // Optional: Decide whether to continue if ERP fails
        }

        // --- Extract Order Data using AI ---
        Order orderData = null;
        if (!string.IsNullOrEmpty(openaiApiKey) && emailScreenshot != null && emailScreenshot.Success)
        {
            Console.WriteLine("Asking OpenAI to extract order data from screenshot...");
            try
            {
                var openAiClient = new OpenAIClient(openaiApiKey);
                var chatClient = openAiClient.GetChatClient("gpt-4o");
                var options = new ChatCompletionOptions { ResponseFormat = ChatResponseFormat.CreateJsonObjectFormat() };

                var prompt = $@"Extract the order details from the email in the screenshot. Provide the output strictly in the following JSON format:
{{
  ""customerName"": ""name of the customer"",
  ""orderedArticles"": [
    {{
      ""articleName"": ""name of the article"",
      ""quantity"": quantity_as_number,
      ""pricePerUnit"": price_as_number
    }}
    // ... more articles if present
  ]
}}";

                var chatMessages = new List<ChatMessage> {
                    ChatMessage.CreateUserMessage(new ChatMessageContentPart[] {
                        ChatMessageContentPart.CreateTextPart(prompt),
                        ChatMessageContentPart.CreateImagePart(BinaryData.FromBytes(emailScreenshot.ImageBytes), "image/jpeg")
                    })
                };

                var result = await chatClient.CompleteChatAsync(chatMessages, options);
                var jsonResponse = result.Value.Content[0].Text;
                Console.WriteLine($"OpenAI Order Extraction Response: {jsonResponse}");

                orderData = JsonSerializer.Deserialize<Order>(jsonResponse, new JsonSerializerOptions { PropertyNameCaseInsensitive = true });

                if (orderData != null)
                {
                    Console.WriteLine($"Successfully extracted order for customer: {orderData.CustomerName}");
                }
                else
                {
                    Console.WriteLine("Warning: Could not deserialize order data from OpenAI response.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error calling OpenAI for order extraction: {ex.Message}");
            }
        }
        else
        {
            Console.WriteLine("Skipping AI order extraction (OpenAI key missing or screenshot failed).");
        }

        // --- Automate ERP Data Entry ---
        if (!string.IsNullOrEmpty(openaiApiKey) && orderData != null && !string.IsNullOrEmpty(erpExePath))
        {
            Console.WriteLine("Attempting to automate data entry into mock ERP...");
            try
            {
                // 1. Get Overview and Find ERP Window
                Console.WriteLine("Getting system overview...");
                var overview = await client.System.GetOverviewAsync();

                string windowDetailsJson;

                if (overview.FocusInfo?.FocusedElementParentWindow?.Title == "ERP system")
                {
                    windowDetailsJson = overview.FocusInfo.FocusedElementParentWindow.ToJsonString();
                }
                else
                {

                    // Assuming the mock ERP has a unique title. Adjust if necessary.
                    var erpWindow = overview?.Windows?.FirstOrDefault(w => w.Title != null && w.Title.Contains("ERP system", StringComparison.OrdinalIgnoreCase));

                    if (erpWindow == null)
                    {
                        Console.WriteLine("Error: Could not find the Mock ERP window.");
                        return;
                    }
                    Console.WriteLine($"Found Mock ERP window: {erpWindow.Id} - {erpWindow.Title}");

                    Console.WriteLine("Getting ERP window details...");
                    var windowDetails = await client.System.GetWindowDetailsAsync(erpWindow.Id);
                    if (windowDetails?.UserInterfaceElements == null)
                    {
                        Console.WriteLine("Error: Could not get details for the Mock ERP window.");
                        return;
                    }
                    windowDetailsJson = windowDetails.ToJsonString();
                }
                //Console.WriteLine($"Window Details JSON: {windowDetailsJson}"); // Optional: log for debugging

                // 2. Get Element IDs using AI
                Console.WriteLine("Asking OpenAI to identify ERP element IDs...");
                ErpElementIds erpElementIds = null;
                try
                {
                    var openAiClient = new OpenAIClient(openaiApiKey);
                    var chatClient = openAiClient.GetChatClient("gpt-4o");
                    var options = new ChatCompletionOptions { ResponseFormat = ChatResponseFormat.CreateJsonObjectFormat() };
                    var prompt = $@"Based on the following UI automation tree JSON for the 'Mini ERP Mock' application, identify the element IDs for the specified controls. Provide the output strictly in the following JSON format:
{{
  ""elementIdCustomerName"": ""ID_for_customer_name_input"",
  ""elementIdArticleName"": ""ID_for_article_name_input"",
  ""elementIdQuantity"": ""ID_for_quantity_input"",
  ""elementIdPricePerUnit"": ""ID_for_price_input"",
  ""elementIdAddItemButton"": ""ID_for_add_item_button"",
  ""elementIdSaveOrderButton"": ""ID_for_save_order_button""
}}

UI Automation Tree JSON:
{windowDetailsJson}";


                    var result = await chatClient.CompleteChatAsync([ChatMessage.CreateUserMessage(prompt)], options);
                    var jsonResponse = result.Value.Content[0].Text;
                    Console.WriteLine($"OpenAI Element ID Response: {jsonResponse}");

                    erpElementIds = JsonSerializer.Deserialize<ErpElementIds>(jsonResponse, new JsonSerializerOptions { PropertyNameCaseInsensitive = true });

                    if (erpElementIds == null || string.IsNullOrWhiteSpace(erpElementIds.ElementIdCustomerName)) // Basic check
                    {
                        Console.WriteLine("Error: Could not deserialize or find necessary element IDs from OpenAI response.");
                        return; // Stop if we don't have IDs
                    }
                    Console.WriteLine("Successfully identified ERP element IDs.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error calling OpenAI for element ID extraction: {ex.Message}");
                    return; // Stop if AI fails here
                }


                // 3. Enter Data using Automation
                Console.WriteLine($"Entering customer name: {orderData.CustomerName} into element {erpElementIds.ElementIdCustomerName}");
                await client.Automation.SetValueAsync(erpElementIds.ElementIdCustomerName, orderData.CustomerName);
                await Task.Delay(500); // Small delay between actions

                foreach (var article in orderData.OrderedArticles)
                {
                    Console.WriteLine($"Entering article: {article.ArticleName}");
                    await client.Automation.SetValueAsync(erpElementIds.ElementIdArticleName, article.ArticleName);
                    await Task.Delay(200);
                    await client.Automation.SetValueAsync(erpElementIds.ElementIdQuantity, article.Quantity.ToString());
                    await Task.Delay(200);
                    await client.Automation.SetValueAsync(erpElementIds.ElementIdPricePerUnit, article.PricePerUnit.ToString("F2")); // Format price
                    await Task.Delay(200);

                    Console.WriteLine("Clicking 'Add Item' button...");
                    await client.Automation.InvokeAsync(erpElementIds.ElementIdAddItemButton);
                    await Task.Delay(1000); // Delay after adding item
                }

                Console.WriteLine("Clicking 'Save Order' button...");
                await client.Automation.InvokeAsync(erpElementIds.ElementIdSaveOrderButton);
                await Task.Delay(500);

                Console.WriteLine("Data entry automation complete.");

            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during ERP data entry automation: {ex.Message}");
            }
        }
        else
        {
            Console.WriteLine("Skipping ERP data entry (AI steps failed or ERP not running).");
        }


        Console.WriteLine();
        Console.WriteLine("Email-to-ERP Example finished. Press any key to exit.");
        Console.ReadKey();
    }

    // --- Helper Methods ---

    private async Task<ScreenshotResponse> GetOrderScreenshotFromGmailAsync(SmoothOperatorClient client)
    {
        /*
         * Example Email Content to send to your Gmail for testing:
         *
         * Subject: New Computerstuff.com Order
         * Text:
         * Dear you,
         *
         * I just visited our customer Smith & Co. Ltd.
         * They want to order:
         *
         * - Product Name: High-Speed Router X200
         *   Quantity: 5 units
         *   Price per unit: 120.00
         *
         * - Product Name: Cat6 Ethernet Cable (10m)
         *   Quantity: 10 units
         *   Price per unit: 15.00
         *
         * Best regards,
         * John Doe
         * Sales Representative
         */

        Console.WriteLine("Opening Gmail in Chrome...");
        // ForceClose strategy might be disruptive, consider alternatives if needed
        await client.Chrome.OpenChromeAsync("https://mail.google.com/", ExistingChromeInstanceStrategy.ForceClose);
        await Task.Delay(10000); // Generous delay for Gmail load and potential login

        // Basic navigation - might need adjustments based on Gmail's UI state (e.g., login prompts)
        Console.WriteLine("Searching for 'order' in Gmail...");
        // Use description-based click for search bar if direct typing fails
        await client.Mouse.ClickByDescriptionAsync("the search mail input field"); // Adjust description if needed
        await Task.Delay(1000);
        await client.Keyboard.TypeAsync("New Computerstuff.com Order");
        await Task.Delay(500);
        await client.Keyboard.PressAsync("Enter");
        await Task.Delay(5000); // Wait for search results

        Console.WriteLine("Clicking the first email in the search results...");
        // This description is a guess, might need refinement or use of coordinates/automation IDs if unreliable
        await client.Mouse.ClickByDescriptionAsync("the first email result in the list");
        await Task.Delay(5000); // Wait for email to load

        Console.WriteLine("Taking screenshot of the email...");
        var screenshot = await client.Screenshot.TakeAsync();
        return screenshot;
    }

    private async Task<ScreenshotResponse> GetOrderScreenshotFromOutlookAsync(SmoothOperatorClient client)
    {
        Console.WriteLine("Opening Outlook...");
        try
        {
            await client.System.OpenApplicationAsync("outlook");
            await Task.Delay(10000); // Generous delay for Outlook to load
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to open Outlook: {ex.Message}. Make sure Outlook is installed.");
            return new ScreenshotResponse { Success = false, Message = "Failed to open Outlook." };
        }

        Console.WriteLine("Searching for 'order' in Outlook...");
        // Using keyboard shortcuts for search - might vary slightly between Outlook versions
        await client.Keyboard.PressAsync("Ctrl+E"); // Focus search bar shortcut
        await Task.Delay(2000);
        await client.Keyboard.TypeAsync("New Computerstuff.com Order");
        await Task.Delay(2*60*1000);
        await client.Keyboard.PressAsync("Enter");
        await Task.Delay(5000); // Wait for search results

        Console.WriteLine("Clicking the first email in the Outlook search results...");
        // Using description-based click - highly dependent on Outlook's UI and potentially unreliable.
        // Consider using AutomationApi.GetOverviewAsync/GetWindowDetailsAsync and InvokeAsync for robustness.
        await client.Mouse.ClickByDescriptionAsync("the first email shown in the list pane"); // Adjust description as needed
        await Task.Delay(5000); // Wait for email to load in reading pane or open

        Console.WriteLine("Taking screenshot of Outlook...");
        var screenshot = await client.Screenshot.TakeAsync();
        return screenshot;
    }

    private async Task<string> DownloadMockErpAsync()
    {
        string downloadUrl = "https://www.dropbox.com/scl/fi/4qc9w57zrmmisyqu3ojnp/mini-erp-mock.exe?rlkey=x5m3ob810zt1scf0mpfn15l4v&dl=1";
        string tempPath = Path.GetTempPath();
        string fileName = "mini-erp-mock.exe";
        string destinationPath = Path.Combine(tempPath, fileName);

        if (File.Exists(destinationPath))
        {
            Console.WriteLine("Mock ERP already exists, skipping download.");
            return destinationPath;
        }

        using (var httpClient = new HttpClient())
        {
            using (var response = await httpClient.GetAsync(downloadUrl, HttpCompletionOption.ResponseHeadersRead))
            {
                response.EnsureSuccessStatusCode();
                using (var streamToReadFrom = await response.Content.ReadAsStreamAsync())
                {
                    using (var streamToWriteTo = File.Open(destinationPath, FileMode.Create))
                    {
                        await streamToReadFrom.CopyToAsync(streamToWriteTo);
                    }
                }
            }
        }
        return destinationPath;
    }

    // Nested classes for deserializing OpenAI responses
    private class Order
    {
        [JsonPropertyName("customerName")]
        public string CustomerName { get; set; }

        [JsonPropertyName("orderedArticles")]
        public List<OrderedArticle> OrderedArticles { get; set; }
    }

    private class OrderedArticle
    {
        [JsonPropertyName("articleName")]
        public string ArticleName { get; set; }

        [JsonPropertyName("quantity")]
        public int Quantity { get; set; }

        [JsonPropertyName("pricePerUnit")]
        public decimal PricePerUnit { get; set; }
    }

    private class ErpElementIds
    {
        [JsonPropertyName("elementIdCustomerName")]
        public string ElementIdCustomerName { get; set; }

        [JsonPropertyName("elementIdArticleName")]
        public string ElementIdArticleName { get; set; }

        [JsonPropertyName("elementIdQuantity")]
        public string ElementIdQuantity { get; set; }

        [JsonPropertyName("elementIdPricePerUnit")]
        public string ElementIdPricePerUnit { get; set; }

        [JsonPropertyName("elementIdAddItemButton")]
        public string ElementIdAddItemButton { get; set; }

        [JsonPropertyName("elementIdSaveOrderButton")]
        public string ElementIdSaveOrderButton { get; set; }
    }

}