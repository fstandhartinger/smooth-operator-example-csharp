# Smooth Operator C# Client Example

This project provides a simple example of how to use the [`SmoothOperator.AgentTools`](https://www.nuget.org/packages/SmoothOperator.AgentTools/) C# client library to interact with the Smooth Operator server.

## Features

This example demonstrates:

*   Initializing the `SmoothOperatorClient`.
*   Starting the Smooth Operator background server asynchronously (`StartServerAsync`).
*   Opening an application (Windows Calculator) using `System.OpenApplicationAsync`.
*   Using the keyboard to type input (`Keyboard.TypeAsync`).
*   Using the mouse with ScreenGrasp to click UI elements by description (`Mouse.ClickByDescriptionAsync`).
*   Retrieving the UI automation tree of the focused window (`System.GetOverviewAsync`).
*   (Optional) Using the `OpenAI-DotNet` library (specifically a GPT-4o model) to interpret the automation tree and determine the application's state (e.g., the result displayed in the calculator).
*   (Optional, commented out) Taking a screenshot (`Screenshot.TakeAsync`) and using the OpenAI API to analyze it.

## Prerequisites

*   .NET SDK (widely compatible with almost all versions).
*   A ScreenGrasp API key. Get a free key from [https://screengrasp.com/api.html](https://screengrasp.com/api.html).
*   (Optional) An OpenAI API key if you want to use the GPT-4o integration for result verification. Get a key from [https://platform.openai.com/api-keys](https://platform.openai.com/api-keys).
*   The Smooth Operator Agent Tools library itself handles the server installation on first use or when `StartServerAsync` is called (placing it in `%APPDATA%\SmoothOperator\AgentToolsServer`).
*   You can optionally handle that installation manually by extracting the contents of the [zip file](https://github.com/fstandhartinger/smooth-operator-client-csharp/blob/master/SmoothOperator.AgentTools/smooth-operator-server.zip) there (e.g. as part of a setup routine)

## Setup

1.  **Clone the repository (if you haven't already):**
    ```bash
    # Navigate to the parent directory where you want the 'smooth-operator' folder
    git clone <repository-url> # Replace with the actual URL
    cd smooth-operator/client-libs/example-csharp
    ```

2.  **Restore NuGet packages:**
    Open the solution (`example-csharp.sln`) in Visual Studio (it should restore automatically) or use the dotnet CLI:
    ```bash
    cd example-csharp # Navigate into the project directory
    dotnet restore
    ```

3.  **Configure API keys:**
    *   Navigate to the `example-csharp` subfolder: `cd example-csharp`
    *   Rename the `.env.example` file to `.env`.
    *   Open the `.env` file and replace the placeholder values with your actual ScreenGrasp API key and (optionally) your OpenAI API key.
    *   **Important:** The `.env` file is listed in the `.gitignore` and should not be committed to version control.

## Running the Example

You can run the example from Visual Studio by setting `example-csharp` as the startup project and pressing F5, or use the dotnet CLI from the `client-libs/example-csharp/example-csharp` directory:

```bash
dotnet run
```

*(Alternatively, from the `client-libs/example-csharp` directory, you can run `dotnet run --project example-csharp`)*

Be aware: if you place breakpoints in the example you might affect functionality, because the code assumes the app it opened stays in focus.

The application will:

1.  Start the Smooth Operator server connection (installing/updating it if necessary).
2.  Open the Windows Calculator.
3.  Type "3+4".
4.  Click the "equals" button using ScreenGrasp.
5.  Retrieve the calculator's UI state (automation tree).
6.  If an OpenAI key was provided in `.env`, it will ask GPT-4o what result is displayed.
7.  Print the result from OpenAI (if applicable).
8.  Wait for you to press a key before exiting.

## Notes

*   This example uses a `.env` file to manage API keys. This is a common practice to keep sensitive credentials out of source control.
*   The OpenAI integration is optional. If you don't provide an API key in the `.env` file, that part of the example will be skipped.
*   The code includes a commented-out section demonstrating how to use screenshots instead of the automation tree for analysis. Screenshots are generally less reliable and more costly in terms of API credits than using the automation tree, but offer a great fallback alternative for situations where the contents aren't programatically accessible otherwise. 