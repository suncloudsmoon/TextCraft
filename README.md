# TextCraft
TextCraft™️ is an add-in for Microsoft® Word® that seamlessly integrates essential AI tools, including text generation, proofreading, and more, directly into the user interface. Designed for offline use, TextCraft™️ allows you to access AI-powered features without requiring an internet connection, making it a more privacy-friendly alternative to Microsoft® Copilot™️.


https://github.com/user-attachments/assets/0e37b253-3f95-4ff2-ab3f-eba480df4c61


# Prerequisities
To install this application, ensure your system meets the following requirements:
1. Windows 10 version 22H2 or later (32/64 bit)

# Install
To install TextCraft™, the Office® Add-In with integrated AI tools, follow these steps:
1. **Install** [Ollama™️](https://ollama.com/download).
2. **Pull** the language model of your choice, for example:
   - `ollama pull qwen2:1.5b-instruct-q4_K_M`
4. **Pull** an embedding model of your choice, for example:
   - `ollama pull all-minilm`
6. **Download the appropriate setup file:**
    - For a 32-bit system, download [`TextCraft_x32.zip`](https://github.com/suncloudsmoon/TextCraft/releases/download/v1.0.0/TextCraft_x32.zip).
    - For a 64-bit system, download [`TextCraft_x64.zip`](https://github.com/suncloudsmoon/TextCraft/releases/download/v1.0.0/TextCraft_x64.zip).
7. **Extract the contents** of the downloaded zip file to a folder of your choice.
8. **Run** `setup.exe`: This will install any required dependencies for TextCraft™, including .NET Framework® 4.8.1 and Visual Studio® 2010 Tools for Office Runtime.
9. **Run** `OfficeAddInSetup.msi` to install TextCraft™.
10. **Open Microsoft Word**® to confirm that TextCraft™ has been successfully installed with its integrated AI tools.

# FAQ
1. **Can I change the OpenAI®️ endpoint the Office®️ add-in uses?**
    - Absolutely! You can easily update the OpenAI®️ endpoint by adjusting the user environment variable `TEXTFORGE_OPENAI_ENDPOINT`. And if you need to set a new API key, just modify the `TEXTFORGE_API_KEY` variable.
2. **How do I switch the embed model for the Office®️ add-in?**
    - It's simple! Just update the `TEXTFORGE_EMBED_MODEL` environment variable, and you're good to go.
3. **What should I do if I run into any issues?**
    - Don't worry! If you hit a snag, just [open an issue in this repository](https://github.com/suncloudsmoon/TextCraft/issues/new), and we'll be happy to help.

# Credits
1. [Ollama™️](https://github.com/ollama/ollama)
2. [OpenAI®️ .NET API library](https://github.com/openai/openai-dotnet)
3. [HyperVectorDB](https://github.com/suncloudsmoon/HyperVectorDB)
4. [PdfPig](https://github.com/UglyToad/PdfPig)
 
