# Ollama Word Add-in

An AI-powered Microsoft Word add-in that integrates with Ollama to provide writing assistance directly within Word documents.

## Features

- **Quick Actions**: Improve text, summarize, fix grammar, and translate
- **Custom Prompts**: Execute custom AI prompts with selected text
- **Model Selection**: Choose from available Ollama models with default model setting
- **Real-time Generation**: Stream responses as they're generated with progress indicator
- **Smart Text Integration**: Use selected text as context with real-time word count
- **Advanced Settings**: Fine-tune model parameters (temperature, top-p, top-k, repeat penalty)
- **System Prompts**: Customize AI behavior with system-level instructions
- **Document Integration**: Insert AI responses directly into Word documents
- **Modern UI**: Clean, compact design optimized for task pane usage
- **Real-time Status**: Connection status indicator for Ollama server

## Prerequisites

1. **Ollama**: Install and run Ollama on your machine
   - Download from [https://ollama.ai](https://ollama.ai)
   - Pull at least one model (e.g., `ollama pull llama2`)
   - Ensure Ollama is running on `http://localhost:11434`

2. **Microsoft Word**: Word 2016 or later (Windows/Mac) or Word Online

## Installation

1. **Clone or download** this repository
2. **Install dependencies**:
   ```bash
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

## Usage

1. **Open the add-in** by clicking the "Show Ollama" button in the Home ribbon
2. **Check connection** - the status indicator should show "Connected to Ollama"
3. **Select a model** from the dropdown (models are loaded automatically)
4. **Use Quick Actions**:
   - Select text and click "Improve Text", "Summarize", "Fix Grammar", or "Translate"
5. **Custom Prompts**:
   - Enter a custom prompt in the text area
   - Optionally check "Use selected text as context" (shows real-time word count)
   - Click "Execute Prompt"
6. **Real-time Generation**: Watch responses stream in as they're generated
7. **Response Management**:
   - **Insert into Document**: Add AI responses to your Word document
   - **Copy**: Copy responses to clipboard
   - **Clear**: Remove response output to start fresh
8. **Stop Generation**: Cancel ongoing AI generation at any time

## Configuration

### Ollama Server Settings
- **URL**: Default is `http://localhost:11434`
- **Temperature**: Controls randomness (0.0 to 1.0)
- **Max Tokens**: Maximum response length
- **Top P**: Nucleus sampling parameter (0.0 to 1.0)
- **Top K**: Top-K sampling parameter (1 to 100)
- **Repeat Penalty**: Prevents repetition (0.5 to 2.0)
- **System Prompt**: Custom instructions for AI behavior
- **Default Model**: Set preferred model for automatic selection

### Supported Actions
- **Improve Text**: Enhances clarity and writing quality
- **Summarize**: Creates concise summaries
- **Fix Grammar**: Corrects grammar and spelling
- **Translate**: Translates to specified languages
- **Custom Prompts**: Execute any custom AI prompt with context
- **Real-time Streaming**: Watch AI responses generate live
- **Smart Context**: Use selected text with automatic word counting

## Development

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
- **Word Integration**: Uses Office.js for document manipulation and text insertion
- **Real-time UI**: Progress indicators, streaming responses, and live updates
- **Smart Context**: Automatic text selection monitoring and word counting
- **Advanced Settings**: Comprehensive model parameter control and persistence
- **Modern UI**: Responsive design with glass morphism and smooth animations

## New Features & Improvements

### 🚀 **Real-time Generation**
- **Streaming Responses**: Watch AI responses generate word-by-word in real-time
- **Progress Indicators**: Visual feedback during generation with animated progress bars
- **Stop Generation**: Cancel ongoing AI generation at any time
- **Auto-scrolling**: Response area automatically scrolls during streaming

### 🎯 **Smart Text Integration**
- **Real-time Word Count**: See word count update automatically as you change text selection
- **Context Awareness**: Use selected text as context for AI prompts
- **Selection Monitoring**: Automatic detection of text selection changes
- **Smart Insertion**: Insert responses at cursor position or document end

### ⚙️ **Advanced Model Control**
- **Fine-tuned Parameters**: Control temperature, top-p, top-k, and repeat penalty
- **System Prompts**: Customize AI behavior with system-level instructions
- **Default Model**: Set preferred model for automatic selection on startup
- **Parameter Persistence**: All settings saved between sessions

### 🎨 **Modern User Interface**
- **Compact Design**: Optimized for Word task pane with minimal scrolling
- **Glass Morphism**: Modern visual effects with backdrop blur and transparency
- **Responsive Layout**: Adapts to different task pane sizes
- **Smooth Animations**: Hover effects and transitions for better user experience

### 📝 **Response Management**
- **Clear Button**: Remove response output to start fresh
- **Smart Button States**: Buttons enable/disable based on response availability
- **Multiple Insert Options**: Insert at cursor or document end
- **Copy to Clipboard**: Easy copying of AI responses

## Troubleshooting

### Connection Issues
- Ensure Ollama is running: `ollama serve`
- Check if models are available: `ollama list`
- Verify server URL in settings

### Add-in Not Loading
- Check that the development server is running on port 3000
- Ensure manifest.xml is valid
- Try refreshing Word or reloading the add-in

### CORS Issues
- Use `npm run dev` which includes CORS headers
- Ensure Ollama allows cross-origin requests

## License

MIT License - see LICENSE file for details

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## Design Philosophy

### 🎯 **Compact & Efficient**
- **Minimal Scrolling**: Most content fits in standard task pane height
- **Smart Spacing**: Optimized margins and padding for space efficiency
- **Progressive Disclosure**: Settings hidden by default, expandable when needed
- **Task-focused Layout**: Essential features always visible, advanced features accessible

### 🔄 **User Experience**
- **Real-time Feedback**: Immediate response to user actions
- **Contextual Help**: Word counts and status indicators provide guidance
- **Persistent Settings**: User preferences remembered between sessions
- **Error Handling**: Clear messages and graceful fallbacks

## Support

For issues and questions:
1. Check the troubleshooting section
2. Review Ollama documentation
3. Open an issue on GitHub
