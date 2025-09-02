# Ollama Word Add-in

An AI-powered Microsoft Word add-in that integrates with Ollama to provide writing assistance directly within Word documents.

## ✨ Features

### 🚀 **Quick Actions**
- **Improve Text**: Enhance clarity and writing quality
- **Summarize**: Create concise summaries of selected text
- **Fix Grammar**: Correct grammar, spelling, and punctuation errors
- **Translate**: Translate text to 15+ languages (Spanish, French, German, Italian, Portuguese, Chinese, Japanese, Korean, Russian, Arabic, Dutch, Swedish, Norwegian, Danish, Finnish)

### 🧠 **Smart AI Integration**
- **Model Pre-loading**: Automatically initializes your preferred model on startup for faster responses
- **Smart Default Model**: Intelligently selects the best available model (prioritizes gpt-oss:latest, llama2:latest, etc.)
- **Real-time Streaming**: Watch AI responses generate word-by-word with progress indicators
- **PMID Search**: Automatically detects PMIDs in selected text and fetches PubMed article details as context

### 📝 **Custom Prompts**
- **Flexible Prompting**: Execute any custom AI prompt with selected text
- **Context Integration**: Use selected text as context with real-time word count
- **PMID Enhancement**: Automatically fetch and include PubMed article abstracts when PMIDs are detected
- **Response Management**: Insert, copy, or clear AI responses

### ⚙️ **Advanced Configuration**
- **Model Parameters**: Fine-tune temperature, top-p, top-k, repeat penalty, and max tokens
- **System Prompts**: Customize AI behavior with system-level instructions
- **Translation Settings**: Set your preferred translation language
- **Persistent Settings**: All preferences saved between sessions

### 🎨 **Modern Interface**
- **Clean Design**: Optimized for Word task pane with minimal scrolling
- **Real-time Status**: Connection status indicator for Ollama server
- **Smart UI**: Buttons enable/disable based on context and response availability
- **Responsive Layout**: Adapts to different task pane sizes

## 🛠️ Prerequisites

1. **Ollama**: Install and run Ollama on your machine
   - Download from [https://ollama.ai](https://ollama.ai)
   - Pull at least one model (e.g., `ollama pull gpt-oss` or `ollama pull llama2`)
   - Ensure Ollama is running on `http://localhost:11434`

2. **Microsoft Word**: Word 2016 or later (Windows/Mac) or Word Online

## 📦 Installation

1. **Clone or download** this repository
2. **Install dependencies**:
   ```bash
   # Download Node.js at https://nodejs.org/en/download/
   # Now you can use npm to install the add-in
   npm install
   ```

3. **Start the development server**:
   ```bash
   npm start
   ```
   This will serve the add-in at `http://localhost:3000`

4. **Sideload the add-in** in Word:
   - Open Word
   - Go to **Insert** > **Add-ins** > **Upload My Add-in**
   - Select the `manifest.xml` file
   - The add-in will appear in the Home ribbon
   
   **For Mac users**: If "Upload My Add-in" is not available, copy the `manifest.xml` file to:
   `~/Library/Containers/com.microsoft.Word/Data/Documents/wef/`

## 🎯 Usage

### Quick Start
1. **Open the add-in** by clicking the "Show Ollama" button in the Home ribbon
2. **Check connection** - the status indicator should show "Connected to Ollama (model ready)"
3. **Select text** in your Word document
4. **Use Quick Actions**: Click "Improve Text", "Summarize", "Fix Grammar", or "Translate"
5. **Watch responses** stream in real-time as they're generated

### Custom Prompts
1. **Enter a custom prompt** in the text area
2. **Enable "Use selected text as context"** (default: enabled)
3. **Enable "PMID Search"** to automatically fetch PubMed articles when PMIDs are detected
4. **Click "Execute Prompt"** to run your custom prompt

### Advanced Features
- **PMID Integration**: When you select text containing PMIDs (like "PMID:12345678"), the add-in automatically fetches the article's title, authors, journal, and abstract to provide rich context
- **Model Pre-loading**: The add-in automatically initializes your preferred model on startup for faster first responses
- **Smart Defaults**: Automatically selects the best available model and sets it as default

## ⚙️ Configuration

### Model Settings
- **URL**: Ollama server URL (default: `http://localhost:11434`)
- **Temperature**: Controls randomness (0.0 to 1.0)
- **Max Tokens**: Maximum response length
- **Top P**: Nucleus sampling parameter (0.0 to 1.0)
- **Top K**: Top-K sampling parameter (1 to 100)
- **Repeat Penalty**: Prevents repetition (0.5 to 2.0)
- **System Prompt**: Custom instructions for AI behavior
- **Default Model**: Set preferred model for automatic selection

### Translation Settings
- **Translation Language**: Choose from 15+ languages for the Translate button
- **Persistent Selection**: Your language choice is remembered between sessions

### PMID Search Settings
- **Enable PMID Search**: Automatically fetch PubMed article details when PMIDs are detected
- **Smart Context**: Includes article title, authors, journal, publication date, DOI, and abstract

## 🏗️ Development

### File Structure
```
ollama-word-addin/
├── manifest.xml          # Add-in manifest
├── taskpane.html         # Main UI
├── taskpane.css          # Styling
├── taskpane.js           # Core functionality
├── commands.html         # Command functions
├── package.json          # Dependencies
├── assets/               # Icons and images
└── README.md            # This file
```

### Key Components
- **OllamaAPI Class**: Handles communication with Ollama server including streaming
- **PubMedAPI Class**: Fetches article details from PubMed for PMID integration
- **Word Integration**: Uses Office.js for document manipulation and text insertion
- **Real-time UI**: Progress indicators, streaming responses, and live updates
- **Smart Context**: Automatic text selection monitoring and word counting
- **Settings Management**: Comprehensive parameter control and persistence

### Available Scripts
- `npm start`: Start development server (port 3000)
- `npm run dev`: Start with CORS headers enabled
- `npm run validate`: Validate the add-in manifest

## 🔧 Troubleshooting

### Connection Issues
- Ensure Ollama is running: `ollama serve`
- Check if models are available: `ollama list`
- Verify server URL in settings
- Check console for connection status messages

### Model Loading Issues
- The add-in automatically initializes models on startup
- Check console for initialization messages
- Large models may take longer to initialize (up to 15 seconds)

### Add-in Not Loading
- Check that the development server is running on port 3000
- Ensure manifest.xml is valid
- Try refreshing Word or reloading the add-in
- For Mac: Ensure manifest.xml is in the correct directory

### CORS Issues
- Use `npm run dev` which includes CORS headers
- Ensure Ollama allows cross-origin requests

### PMID Search Issues
- Check internet connection for PubMed API access
- Verify PMID format (8-digit numbers, optionally prefixed with "PMID:")
- Check console for PubMed API error messages

## 🎨 Design Philosophy

### 🎯 **User-Centric Design**
- **Minimal Scrolling**: Most content fits in standard task pane height
- **Smart Defaults**: Sensible defaults that work for most users
- **Progressive Disclosure**: Settings hidden by default, expandable when needed
- **Task-focused Layout**: Essential features always visible

### 🔄 **Performance Optimized**
- **Model Pre-loading**: Faster first responses through automatic initialization
- **Smart Caching**: PubMed results cached for 24 hours
- **Real-time Feedback**: Immediate response to user actions
- **Efficient UI**: Minimal resource usage

### 🛡️ **Robust & Reliable**
- **Error Handling**: Clear messages and graceful fallbacks
- **Persistent Settings**: User preferences remembered between sessions
- **Connection Monitoring**: Real-time status indicators
- **Graceful Degradation**: Works even when some features fail

## 📄 License

MIT License - see LICENSE file for details

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## 📞 Support

For issues and questions:
1. Check the troubleshooting section above
2. Review Ollama documentation at [https://ollama.ai](https://ollama.ai)
3. Open an issue on GitHub

## 🆕 Recent Updates

- **PMID Search Integration**: Automatically fetch PubMed article details for enhanced context
- **Model Pre-loading**: Faster startup with automatic model initialization
- **Smart Default Model Selection**: Intelligently chooses the best available model
- **Translation Language Settings**: Persistent language selection in settings
- **Improved UI Organization**: Clean separation between Quick Actions and Settings
- **Enhanced Error Handling**: Better feedback and graceful fallbacks