/* global Office, console */

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        document.getElementById("improveBtn").onclick = () => improveText();
        document.getElementById("summarizeBtn").onclick = () => summarizeText();
        document.getElementById("grammarBtn").onclick = () => fixGrammar();
        document.getElementById("translateBtn").onclick = () => translateText();
        document.getElementById("executePrompt").onclick = () => executeCustomPrompt();
        document.getElementById("insertBtn").onclick = () => insertResponse();
        document.getElementById("copyBtn").onclick = () => copyResponse();
        document.getElementById("clearBtn").onclick = () => clearResponse();
        document.getElementById("refreshModels").onclick = () => loadModels();
        document.getElementById("stopGeneration").onclick = () => stopGeneration();
        
        // Settings event listeners
        document.getElementById("temperature").oninput = (e) => {
            document.getElementById("temperatureValue").textContent = e.target.value;
        };
        document.getElementById("topP").oninput = (e) => {
            document.getElementById("topPValue").textContent = e.target.value;
        };
        document.getElementById("repeatPenalty").oninput = (e) => {
            document.getElementById("repeatPenaltyValue").textContent = e.target.value;
        };
        
        // Selection change listener for word count
        document.getElementById("useSelection").onchange = () => updateWordCount();
        
        // Set up real-time selection monitoring
        setupSelectionMonitoring();
        
        // Initialize the add-in
        initializeAddin();
    }
});

// Global variables
let currentResponse = '';
let isGenerating = false;
let abortController = null;
let ollamaSettings = {
    url: 'http://localhost:11434',
    temperature: 0.7,
    maxTokens: 2048,
    topP: 0.9,
    topK: 40,
    repeatPenalty: 1.1,
    systemPrompt: ''
};

// Initialize the add-in
async function initializeAddin() {
    // Load settings from storage
    loadSettings();
    
    // Check Ollama connection
    await checkOllamaConnection();
    
    // Load available models
    await loadModels();
}

// Ollama API Functions
class OllamaAPI {
    constructor(baseUrl) {
        this.baseUrl = baseUrl;
    }

    async checkConnection() {
        try {
            const response = await fetch(`${this.baseUrl}/api/tags`);
            return response.ok;
        } catch (error) {
            console.error('Connection check failed:', error);
            return false;
        }
    }

    async getModels() {
        try {
            const response = await fetch(`${this.baseUrl}/api/tags`);
            if (!response.ok) throw new Error('Failed to fetch models');
            
            const data = await response.json();
            return data.models || [];
        } catch (error) {
            console.error('Failed to get models:', error);
            return [];
        }
    }

    async generateResponse(model, prompt, options = {}, onProgress = null) {
        try {
            abortController = new AbortController();
            
            const requestBody = {
                model: model,
                prompt: prompt,
                stream: !!onProgress,
                options: {
                    temperature: options.temperature || 0.7,
                    num_predict: options.maxTokens || 2048,
                    top_p: options.topP || 0.9,
                    top_k: options.topK || 40,
                    repeat_penalty: options.repeatPenalty || 1.1
                }
            };

            // Add system prompt if provided
            if (options.systemPrompt && options.systemPrompt.trim()) {
                requestBody.system = options.systemPrompt.trim();
            }

            const response = await fetch(`${this.baseUrl}/api/generate`, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify(requestBody),
                signal: abortController.signal
            });

            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }

            if (onProgress) {
                return await this.handleStreamResponse(response, onProgress);
            } else {
                const data = await response.json();
                return data.response;
            }
        } catch (error) {
            if (error.name === 'AbortError') {
                throw new Error('Generation was stopped by user');
            }
            console.error('Generate response failed:', error);
            throw error;
        }
    }

    async handleStreamResponse(response, onProgress) {
        const reader = response.body.getReader();
        const decoder = new TextDecoder();
        let fullResponse = '';

        try {
            while (true) {
                const { done, value } = await reader.read();
                
                if (done) break;
                
                const chunk = decoder.decode(value);
                const lines = chunk.split('\n').filter(line => line.trim());
                
                for (const line of lines) {
                    try {
                        const data = JSON.parse(line);
                        if (data.response) {
                            fullResponse += data.response;
                            onProgress(fullResponse, !data.done);
                        }
                        if (data.done) {
                            return fullResponse;
                        }
                    } catch (e) {
                        // Skip invalid JSON lines
                        console.warn('Invalid JSON in stream:', line);
                    }
                }
            }
        } finally {
            reader.releaseLock();
        }
        
        return fullResponse;
    }
}

// Connection and model management
async function checkOllamaConnection() {
    const statusIndicator = document.getElementById('statusIndicator');
    const statusText = document.getElementById('statusText');
    
    statusIndicator.className = 'status-indicator connecting';
    statusText.textContent = 'Connecting to Ollama...';
    
    const ollama = new OllamaAPI(ollamaSettings.url);
    const isConnected = await ollama.checkConnection();
    
    if (isConnected) {
        statusIndicator.className = 'status-indicator connected';
        statusText.textContent = 'Connected to Ollama';
    } else {
        statusIndicator.className = 'status-indicator disconnected';
        statusText.textContent = 'Failed to connect to Ollama';
    }
    
    return isConnected;
}

async function loadModels() {
    const modelSelect = document.getElementById('modelSelect');
    const defaultModelSelect = document.getElementById('defaultModel');
    
    modelSelect.innerHTML = '<option value="">Loading models...</option>';
    
    const ollama = new OllamaAPI(ollamaSettings.url);
    const models = await ollama.getModels();
    
    modelSelect.innerHTML = '';
    
    // Update default model dropdown
    defaultModelSelect.innerHTML = '<option value="">No default (manual selection)</option>';
    
    if (models.length === 0) {
        modelSelect.innerHTML = '<option value="">No models available</option>';
        return;
    }
    
    models.forEach(model => {
        // Main model selector
        const option = document.createElement('option');
        option.value = model.name;
        option.textContent = model.name;
        modelSelect.appendChild(option);
        
        // Default model selector
        const defaultOption = document.createElement('option');
        defaultOption.value = model.name;
        defaultOption.textContent = model.name;
        defaultModelSelect.appendChild(defaultOption);
    });
    
    // Apply default model selection if set
    const defaultModel = localStorage.getItem('ollama_default_model');
    if (defaultModel && models.some(m => m.name === defaultModel)) {
        modelSelect.value = defaultModel;
        defaultModelSelect.value = defaultModel;
    } else if (models.length > 0) {
        // Select the first model if no default is set
        modelSelect.value = models[0].name;
    }
}

// Word document interaction functions
async function getSelectedText() {
    return new Promise((resolve, reject) => {
        Word.run(async (context) => {
            const selection = context.document.getSelection();
            selection.load('text');
            
            await context.sync();
            resolve(selection.text);
        }).catch(reject);
    });
}

async function insertTextAtCursor(text) {
    return new Promise((resolve, reject) => {
        Word.run(async (context) => {
            try {
                const selection = context.document.getSelection();
                selection.load('text');
                await context.sync();
                
                // Insert text at current selection/cursor position
                selection.insertText(text, Word.InsertLocation.replace);
                await context.sync();
                resolve();
            } catch (error) {
                reject(error);
            }
        }).catch(reject);
    });
}

async function insertTextAtEnd(text) {
    return new Promise((resolve, reject) => {
        Word.run(async (context) => {
            try {
                const body = context.document.body;
                body.insertText('\n\n' + text, Word.InsertLocation.end);
                await context.sync();
                resolve();
            } catch (error) {
                reject(error);
            }
        }).catch(reject);
    });
}

// AI prompt functions
async function improveText() {
    const selectedText = await getSelectedText();
    if (!selectedText.trim()) {
        showError('Please select some text to improve.');
        return;
    }
    
    const prompt = `Please improve the following text by making it clearer, more concise, and better written while maintaining the original meaning:\n\n${selectedText}`;
    await processPrompt(prompt, true);
}

async function summarizeText() {
    const selectedText = await getSelectedText();
    if (!selectedText.trim()) {
        showError('Please select some text to summarize.');
        return;
    }
    
    const prompt = `Please provide a concise summary of the following text:\n\n${selectedText}`;
    await processPrompt(prompt);
}

async function fixGrammar() {
    const selectedText = await getSelectedText();
    if (!selectedText.trim()) {
        showError('Please select some text to check grammar.');
        return;
    }
    
    const prompt = `Please fix any grammar, spelling, and punctuation errors in the following text while keeping the original meaning:\n\n${selectedText}`;
    await processPrompt(prompt, true);
}

async function translateText() {
    const selectedText = await getSelectedText();
    if (!selectedText.trim()) {
        showError('Please select some text to translate.');
        return;
    }
    
    const targetLanguage = prompt('Enter target language (e.g., Spanish, French, German):');
    if (!targetLanguage) return;
    
    const translationPrompt = `Please translate the following text to ${targetLanguage}:\n\n${selectedText}`;
    await processPrompt(translationPrompt);
}

async function executeCustomPrompt() {
    const customPrompt = document.getElementById('customPrompt').value.trim();
    if (!customPrompt) {
        showError('Please enter a custom prompt.');
        return;
    }
    
    const useSelection = document.getElementById('useSelection').checked;
    let finalPrompt = customPrompt;
    
    if (useSelection) {
        const selectedText = await getSelectedText();
        if (selectedText.trim()) {
            finalPrompt += `\n\nText to work with:\n${selectedText}`;
        }
    }
    
    const replaceText = document.getElementById('replaceText').checked;
    await processPrompt(finalPrompt, replaceText);
}

// Core processing function
async function processPrompt(prompt, replaceSelection = false) {
    const modelSelect = document.getElementById('modelSelect');
    const selectedModel = modelSelect.value;
    
    if (!selectedModel) {
        showError('Please select a model first.');
        return;
    }
    
    if (isGenerating) {
        showError('Already generating a response. Please wait or stop the current generation.');
        return;
    }

    isGenerating = true;
    showGenerationProgress(true);
    
    try {
        const ollama = new OllamaAPI(ollamaSettings.url);
        
        // Use streaming for real-time updates
        const response = await ollama.generateResponse(selectedModel, prompt, {
            temperature: ollamaSettings.temperature,
            maxTokens: ollamaSettings.maxTokens,
            topP: ollamaSettings.topP,
            topK: ollamaSettings.topK,
            repeatPenalty: ollamaSettings.repeatPenalty,
            systemPrompt: ollamaSettings.systemPrompt
        }, (partialResponse, isStreaming) => {
            displayResponse(partialResponse, isStreaming);
        });
        
        currentResponse = response;
        displayResponse(response, false);
        
        if (replaceSelection) {
            await insertTextAtCursor(response);
        }
        
    } catch (error) {
        if (error.message.includes('stopped by user')) {
            showSuccess('Generation stopped by user.');
        } else {
            showError(`Failed to generate response: ${error.message}`);
        }
    } finally {
        isGenerating = false;
        showGenerationProgress(false);
    }
}

// Stop generation function
function stopGeneration() {
    if (abortController) {
        abortController.abort();
        abortController = null;
    }
    isGenerating = false;
    showGenerationProgress(false);
}

// Selection monitoring setup
function setupSelectionMonitoring() {
    let lastSelectedText = '';
    
    // Monitor selection changes every 500ms when checkbox is checked
    setInterval(async () => {
        const useSelection = document.getElementById('useSelection').checked;
        if (useSelection) {
            try {
                const currentSelectedText = await getSelectedText();
                if (currentSelectedText !== lastSelectedText) {
                    lastSelectedText = currentSelectedText;
                    updateWordCount();
                }
            } catch (error) {
                // Ignore errors during monitoring
            }
        }
    }, 500);
}

// Word count functionality
async function updateWordCount() {
    const useSelection = document.getElementById('useSelection').checked;
    const wordCountElement = document.getElementById('wordCount');
    
    if (useSelection) {
        try {
            const selectedText = await getSelectedText();
            const wordCount = selectedText.trim() ? selectedText.trim().split(/\s+/).length : 0;
            
            if (wordCount > 0) {
                wordCountElement.textContent = `(${wordCount} word${wordCount !== 1 ? 's' : ''})`;
                wordCountElement.style.display = 'inline';
            } else {
                wordCountElement.textContent = '(no text selected)';
                wordCountElement.style.display = 'inline';
            }
        } catch (error) {
            wordCountElement.textContent = '(selection unavailable)';
            wordCountElement.style.display = 'inline';
        }
    } else {
        wordCountElement.style.display = 'none';
    }
}

// UI helper functions
function displayResponse(response, isStreaming = false) {
    const responseArea = document.getElementById('responseArea');
    responseArea.textContent = response;
    
    // Auto-scroll to bottom during streaming
    if (isStreaming) {
        responseArea.scrollTop = responseArea.scrollHeight;
    }
    
    // Enable response action buttons only when generation is complete
    const hasResponse = response && response.trim().length > 0;
    document.getElementById('insertBtn').disabled = isStreaming || !hasResponse;
    document.getElementById('copyBtn').disabled = isStreaming || !hasResponse;
    document.getElementById('clearBtn').disabled = isStreaming || !hasResponse;
}

function clearResponse() {
    currentResponse = '';
    const responseArea = document.getElementById('responseArea');
    responseArea.innerHTML = '<p class="placeholder-text">AI responses will appear here...</p>';
    
    // Disable action buttons
    document.getElementById('insertBtn').disabled = true;
    document.getElementById('copyBtn').disabled = true;
    document.getElementById('clearBtn').disabled = true;
}

function showGenerationProgress(show) {
    const progressElement = document.getElementById('generationProgress');
    const stopButton = document.getElementById('stopGeneration');
    
    progressElement.style.display = show ? 'flex' : 'none';
    stopButton.style.display = show ? 'inline-block' : 'none';
    
    // Disable action buttons during generation
    const actionButtons = document.querySelectorAll('.action-btn, #executePrompt');
    actionButtons.forEach(btn => {
        btn.disabled = show;
    });
}

async function insertResponse() {
    if (!currentResponse) return;
    
    try {
        // Try to insert at cursor first, then at end if cursor fails
        await insertTextAtCursor(currentResponse);
        showSuccess('Response inserted at cursor position.');
    } catch (error) {
        try {
            await insertTextAtEnd(currentResponse);
            showSuccess('Response inserted at end of document.');
        } catch (endError) {
            showError(`Failed to insert response: ${endError.message}`);
        }
    }
}

function copyResponse() {
    if (!currentResponse) return;
    
    navigator.clipboard.writeText(currentResponse).then(() => {
        showSuccess('Response copied to clipboard.');
    }).catch(() => {
        // Fallback for older browsers
        const textarea = document.createElement('textarea');
        textarea.value = currentResponse;
        document.body.appendChild(textarea);
        textarea.select();
        document.execCommand('copy');
        document.body.removeChild(textarea);
        showSuccess('Response copied to clipboard.');
    });
}

function showLoading(show) {
    const loadingOverlay = document.getElementById('loadingOverlay');
    loadingOverlay.style.display = show ? 'flex' : 'none';
}

function showError(message) {
    // Simple error display - in a production app, you might want a more sophisticated notification system
    alert(`Error: ${message}`);
}

function showSuccess(message) {
    // Simple success display - in a production app, you might want a more sophisticated notification system
    console.log(`Success: ${message}`);
}

// Settings management
function loadSettings() {
    // Load all settings from localStorage
    const savedUrl = localStorage.getItem('ollama_url');
    const savedTemp = localStorage.getItem('ollama_temperature');
    const savedTokens = localStorage.getItem('ollama_max_tokens');
    const savedTopP = localStorage.getItem('ollama_top_p');
    const savedTopK = localStorage.getItem('ollama_top_k');
    const savedRepeatPenalty = localStorage.getItem('ollama_repeat_penalty');
    const savedSystemPrompt = localStorage.getItem('ollama_system_prompt');
    const savedDefaultModel = localStorage.getItem('ollama_default_model');
    
    if (savedUrl) {
        ollamaSettings.url = savedUrl;
        document.getElementById('ollamaUrl').value = savedUrl;
    }
    
    if (savedTemp) {
        ollamaSettings.temperature = parseFloat(savedTemp);
        document.getElementById('temperature').value = savedTemp;
        document.getElementById('temperatureValue').textContent = savedTemp;
    }
    
    if (savedTokens) {
        ollamaSettings.maxTokens = parseInt(savedTokens);
        document.getElementById('maxTokens').value = savedTokens;
    }
    
    if (savedTopP) {
        ollamaSettings.topP = parseFloat(savedTopP);
        document.getElementById('topP').value = savedTopP;
        document.getElementById('topPValue').textContent = savedTopP;
    }
    
    if (savedTopK) {
        ollamaSettings.topK = parseInt(savedTopK);
        document.getElementById('topK').value = savedTopK;
    }
    
    if (savedRepeatPenalty) {
        ollamaSettings.repeatPenalty = parseFloat(savedRepeatPenalty);
        document.getElementById('repeatPenalty').value = savedRepeatPenalty;
        document.getElementById('repeatPenaltyValue').textContent = savedRepeatPenalty;
    }
    
    if (savedSystemPrompt) {
        ollamaSettings.systemPrompt = savedSystemPrompt;
        document.getElementById('systemPrompt').value = savedSystemPrompt;
    }
    
    if (savedDefaultModel) {
        document.getElementById('defaultModel').value = savedDefaultModel;
    }
    
    // Add event listeners for settings changes
    document.getElementById('ollamaUrl').addEventListener('change', (e) => {
        ollamaSettings.url = e.target.value;
        localStorage.setItem('ollama_url', e.target.value);
        checkOllamaConnection();
        loadModels();
    });
    
    document.getElementById('temperature').addEventListener('input', (e) => {
        ollamaSettings.temperature = parseFloat(e.target.value);
        localStorage.setItem('ollama_temperature', e.target.value);
    });
    
    document.getElementById('maxTokens').addEventListener('change', (e) => {
        ollamaSettings.maxTokens = parseInt(e.target.value);
        localStorage.setItem('ollama_max_tokens', e.target.value);
    });
    
    document.getElementById('topP').addEventListener('input', (e) => {
        ollamaSettings.topP = parseFloat(e.target.value);
        localStorage.setItem('ollama_top_p', e.target.value);
    });
    
    document.getElementById('topK').addEventListener('change', (e) => {
        ollamaSettings.topK = parseInt(e.target.value);
        localStorage.setItem('ollama_top_k', e.target.value);
    });
    
    document.getElementById('repeatPenalty').addEventListener('input', (e) => {
        ollamaSettings.repeatPenalty = parseFloat(e.target.value);
        localStorage.setItem('ollama_repeat_penalty', e.target.value);
    });
    
    document.getElementById('systemPrompt').addEventListener('change', (e) => {
        ollamaSettings.systemPrompt = e.target.value;
        localStorage.setItem('ollama_system_prompt', e.target.value);
    });
    
    document.getElementById('defaultModel').addEventListener('change', (e) => {
        localStorage.setItem('ollama_default_model', e.target.value);
        // Update main model selector if a default is chosen
        if (e.target.value) {
            document.getElementById('modelSelect').value = e.target.value;
        }
    });
}
