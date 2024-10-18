# TextCraft
TextCraft is an add-in for Microsoft® Word® that seamlessly integrates essential AI tools, including text generation, proofreading, and more, directly into the user interface. Designed for offline use, TextCraft allows you to access AI-powered features without requiring an internet connection, making it a more privacy-friendly alternative to Microsoft® Copilot™️.


https://github.com/user-attachments/assets/0e37b253-3f95-4ff2-ab3f-eba480df4c61


# Prerequisities
To install this application, ensure your system meets the following requirements:
1. Windows 10 version 20H2 or later (32/64 bit)
2. Microsoft®️ Office®️ 2010 or later

# Install
To install TextCraft, the Office® Add-In with integrated AI tools, follow these steps:
1. **Install** [Ollama™️](https://ollama.com/download).
2. **Pull** the language model of your choice, for example:
   - `ollama pull qwen2.5:1.5b`
4. **Pull** an embedding model of your choice, for example:
   - `ollama pull all-minilm`
6. **Download the appropriate setup file:**
    - For a 32-bit system, download [`TextCraft_x32.zip`](https://github.com/suncloudsmoon/TextCraft/releases/download/v1.0.5/TextCraft_x32.zip).
    - For a 64-bit system, download [`TextCraft_x64.zip`](https://github.com/suncloudsmoon/TextCraft/releases/download/v1.0.5/TextCraft_x64.zip).
7. **Extract the contents** of the downloaded zip file to a folder of your choice.
8. **Run** `setup.exe`: This will install any required dependencies for TextCraft, including .NET Framework® 4.8.1 and Visual Studio® 2010 Tools for Office Runtime.
9. **Run** `OfficeAddInSetup.msi` to install TextCraft.
10. **Open Microsoft Word**® to confirm that TextCraft has been successfully installed with its integrated AI tools.

# Development
To customize and develop this Software, follow these steps:
1. Download and run the [Visual Studio® 2022 installer](https://visualstudio.microsoft.com/vs/).
2. When the installer opens, select both ".NET desktop development" and "Office/SharePoint development." When ".NET desktop development" is selected, under "Installation details" on the right, scroll to the "Optional" section and ensure the ".NET Framework 4.8.1 development tools" checkbox is selected.
3. Click "Install."
4. Clone this repository using Git.
5. Open the solution by double-clicking on the "TextForge.sln" file.
6. First, build the "TextCraft" project (in Release mode), then right-click on the "OfficeAddInSetup" project and select "Build."
7. Navigate to the project directory and find the "OfficeAddInSetup" folder, which contains the add-in setup files.
8. Be sure to clean the "TextCraft" project in both "Debug" and "Release" modes by right-clicking on the project in Visual Studio and selecting "Clean." This is necessary because building the project in Visual Studio, even without running the add-in installer, will automatically install the add-in in Microsoft Word. Cleaning is the only way to remove it.

> **Note**: For more information on VSTO (Visual Studio® Tools for Office®) development, you can refer to the [official Microsoft® documentation](https://learn.microsoft.com/en-us/visualstudio/vsto/walkthrough-creating-your-first-vsto-add-in-for-word?view=vs-2022&tabs=csharp).

# FAQ
1. **Can I change the OpenAI®️ endpoint the Office®️ add-in uses?**
    - Absolutely! You can easily update the OpenAI®️ endpoint by adjusting the user environment variable `TEXTCRAFT_OPENAI_ENDPOINT`. And if you need to set a new API key, just modify the `TEXTCRAFT_API_KEY` variable.
2. **How do I switch the embed model for the Office®️ add-in?**
    - It's simple! Just update the `TEXTCRAFT_EMBED_MODEL` environment variable, and you're good to go.
3. **What should I do if I run into any issues?**
    - Don't worry! If you hit a snag, just [open an issue in this repository](https://github.com/suncloudsmoon/TextCraft/issues/new), and we'll be happy to help.

# Credits
1. [Ollama™️](https://github.com/ollama/ollama)
2. [OpenAI®️ .NET API library](https://github.com/openai/openai-dotnet)
3. [HyperVectorDB](https://github.com/suncloudsmoon/HyperVectorDB)
4. [PdfPig](https://github.com/UglyToad/PdfPig)
 
