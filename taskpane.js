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
        
        // PMID search event listener
        const pmidSearchElement = document.getElementById("enablePmidSearch");
        if (pmidSearchElement) {
            pmidSearchElement.onchange = (e) => {
                pmidSearchEnabled = e.target.checked;
            };
        }
        
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
let pmidSearchEnabled = true;
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
    
    // Initialize Ollama with test message to pre-load model
    await initializeOllamaModel();
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

// Initialize Ollama model with test message to pre-load it
async function initializeOllamaModel() {
    const defaultModelSelect = document.getElementById('defaultModel');
    const selectedModel = defaultModelSelect.value;
    
    if (!selectedModel) {
        console.log('No model selected for initialization');
        return;
    }
    
    try {
        console.log(`🚀 Initializing Ollama model: ${selectedModel} (this will make future requests faster)`);
        
        const ollama = new OllamaAPI(ollamaSettings.url);
        
        // Send a simple test message to pre-load the model
        const testPrompt = "Hello! Please respond with just 'Ready' to confirm the model is loaded.";
        
        // Use a shorter timeout for initialization
        const initController = new AbortController();
        const timeoutId = setTimeout(() => initController.abort(), 15000); // 15 second timeout for larger models
        
        const response = await ollama.generateResponse(selectedModel, testPrompt, {
            temperature: 0.1, // Low temperature for consistent response
            maxTokens: 10,    // Very short response
            topP: 0.9,
            topK: 40,
            repeatPenalty: 1.1,
            systemPrompt: 'Respond with only the word "Ready".'
        });
        
        clearTimeout(timeoutId);
        
        console.log(`✅ Model ${selectedModel} initialized successfully! Response: "${response.trim()}"`);
        
        // Update status to show model is ready
        const statusText = document.getElementById('statusText');
        if (statusText && statusText.textContent.includes('Connected to Ollama')) {
            statusText.textContent = `Connected to Ollama (${selectedModel} ready)`;
        }
        
        return true; // Success
        
    } catch (error) {
        if (error.name === 'AbortError') {
            console.warn(`⏰ Model ${selectedModel} initialization timed out (model may still work, just slower on first use)`);
        } else {
            console.warn(`❌ Failed to initialize model ${selectedModel}:`, error.message);
        }
        // Don't show error to user as this is a background initialization
        // The model will still work, just might take longer on first use
        return false; // Failed
    }
}

// PMID Search and PubMed API Functions
class PubMedAPI {
    constructor() {
        this.baseUrl = 'https://eutils.ncbi.nlm.nih.gov/entrez/eutils';
        this.cache = new Map();
        this.maxCacheSize = 100;
    }

    // Extract PMIDs from text using regex
    extractPMIDs(text) {
        if (!text) return [];
        
        // PMID pattern: 8 digits, optionally preceded by "PMID:", "pmid:", or "PMID"
        const pmidPattern = /(?:PMID:?\s*)?(\d{8})/gi;
        const matches = text.match(pmidPattern);
        
        if (!matches) return [];
        
        // Extract just the numbers and remove duplicates
        const pmids = matches.map(match => {
            const numberMatch = match.match(/\d{8}/);
            return numberMatch ? numberMatch[0] : null;
        }).filter(pmid => pmid !== null);
        
        return [...new Set(pmids)]; // Remove duplicates
    }

    // Fetch article details from PubMed
    async fetchArticleDetails(pmid) {
        // Check cache first
        if (this.cache.has(pmid)) {
            const cached = this.cache.get(pmid);
            if (Date.now() - cached.timestamp < 86400000) { // 24 hours
                return cached.data;
            }
        }

        try {
            console.log(`🔍 Fetching PubMed details for PMID: ${pmid}`);
            
            // First, get the article summary
            const summaryUrl = `${this.baseUrl}/esummary.fcgi?db=pubmed&id=${pmid}&retmode=json`;
            const summaryResponse = await fetch(summaryUrl);
            
            if (!summaryResponse.ok) {
                throw new Error(`HTTP error! status: ${summaryResponse.status}`);
            }
            
            const summaryData = await summaryResponse.json();
            const result = summaryData.result[pmid];
            
            if (!result || result.error) {
                console.warn(`No data found for PMID: ${pmid}`);
                return null;
            }

            // Extract relevant information
            const articleInfo = {
                pmid: pmid,
                title: result.title || 'No title available',
                authors: result.authors ? result.authors.map(a => a.name).join(', ') : 'No authors listed',
                journal: result.source || 'No journal listed',
                pubDate: result.pubdate || 'No date available',
                abstract: result.abstract || 'No abstract available',
                doi: result.elocationid || 'No DOI available'
            };

            // Cache the result
            this.cache.set(pmid, {
                data: articleInfo,
                timestamp: Date.now()
            });

            // Limit cache size
            if (this.cache.size > this.maxCacheSize) {
                const firstKey = this.cache.keys().next().value;
                this.cache.delete(firstKey);
            }

            console.log(`✅ Successfully fetched details for PMID: ${pmid}`);
            return articleInfo;

        } catch (error) {
            console.error(`Failed to fetch details for PMID ${pmid}:`, error);
            return null;
        }
    }

    // Fetch multiple articles and format as context
    async fetchMultipleArticles(pmids) {
        if (!pmids || pmids.length === 0) return '';

        console.log(`📚 Fetching details for ${pmids.length} PMIDs: ${pmids.join(', ')}`);
        
        const articles = [];
        for (const pmid of pmids) {
            const article = await this.fetchArticleDetails(pmid);
            if (article) {
                articles.push(article);
            }
        }

        if (articles.length === 0) {
            console.log('No articles found for the provided PMIDs');
            return '';
        }

        // Format articles as context
        let context = '\n\n--- PubMed Articles Context ---\n';
        articles.forEach((article, index) => {
            context += `\n[Article ${index + 1} - PMID: ${article.pmid}]\n`;
            context += `Title: ${article.title}\n`;
            context += `Authors: ${article.authors}\n`;
            context += `Journal: ${article.journal}\n`;
            context += `Publication Date: ${article.pubDate}\n`;
            context += `DOI: ${article.doi}\n`;
            context += `Abstract: ${article.abstract}\n`;
            context += '---\n';
        });

        console.log(`✅ Formatted context for ${articles.length} articles`);
        return context;
    }
}

// Global PubMed API instance
const pubmedAPI = new PubMedAPI();

// Select the best default model based on available models
function selectBestDefaultModel(models) {
    // Priority order for model selection (best to worst for general use)
    const preferredModels = [
        'gpt-oss:latest',          // Preferred general purpose model
        'llama2:latest',           // Good general purpose model
        'qwen3-coder:latest',      // Good for coding tasks
        'qwen3:30b',              // Good general purpose
        'gemma3:27b-it-qat',      // Good instruction following
    ];
    
    // First, try to find a preferred model
    for (const preferred of preferredModels) {
        const found = models.find(m => m.name === preferred);
        if (found) {
            console.log(`Found preferred model: ${preferred}`);
            return preferred;
        }
    }
    
    // If no preferred model found, look for models with good characteristics
    const goodModels = models.filter(m => {
        const name = m.name.toLowerCase();
        return (
            name.includes('llama') || 
            name.includes('gpt') || 
            name.includes('qwen') ||
            name.includes('gemma')
        );
    });
    
    if (goodModels.length > 0) {
        console.log(`Using good general purpose model: ${goodModels[0].name}`);
        return goodModels[0].name;
    }
    
    // Fallback to first model
    console.log(`Using first available model: ${models[0].name}`);
    return models[0].name;
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
    const defaultModelSelect = document.getElementById('defaultModel');
    
    defaultModelSelect.innerHTML = '<option value="">Loading models...</option>';
    
    const ollama = new OllamaAPI(ollamaSettings.url);
    const models = await ollama.getModels();
    
    // Update default model dropdown
    defaultModelSelect.innerHTML = '<option value="">No default (manual selection)</option>';
    
    if (models.length === 0) {
        defaultModelSelect.innerHTML = '<option value="">No models available</option>';
        return;
    }
    
    models.forEach(model => {
        // Default model selector
        const defaultOption = document.createElement('option');
        defaultOption.value = model.name;
        defaultOption.textContent = model.name;
        defaultModelSelect.appendChild(defaultOption);
    });
    
    // Apply default model selection if set
    const defaultModel = localStorage.getItem('ollama_default_model');
    if (defaultModel && models.some(m => m.name === defaultModel)) {
        defaultModelSelect.value = defaultModel;
        console.log(`Using saved default model: ${defaultModel}`);
    } else if (models.length > 0) {
        // Select the best default model if no default is set
        const bestModel = selectBestDefaultModel(models);
        console.log(`No default model set, using recommended model: ${bestModel}`);
        
        // Auto-set the best model as default for faster future startups
        localStorage.setItem('ollama_default_model', bestModel);
        defaultModelSelect.value = bestModel;
        console.log(`Auto-set ${bestModel} as default model for faster startup`);
    }
    
    // Add event listener for model selection changes to initialize new models
    defaultModelSelect.addEventListener('change', async () => {
        if (defaultModelSelect.value) {
            await initializeOllamaModel();
        }
    });
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
    
    const targetLanguageElement = document.getElementById('targetLanguage');
    if (!targetLanguageElement) {
        showError('Translation language setting not found.');
        return;
    }
    
    const targetLanguage = targetLanguageElement.value;
    if (!targetLanguage) {
        showError('Please select a target language.');
        return;
    }
    
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
    let selectedText = '';
    
    if (useSelection) {
        selectedText = await getSelectedText();
        if (selectedText.trim()) {
            finalPrompt += `\n\nText to work with:\n${selectedText}`;
            
            // Check for PMIDs if PMID search is enabled
            if (pmidSearchEnabled) {
                const pmids = pubmedAPI.extractPMIDs(selectedText);
                if (pmids.length > 0) {
                    console.log(`🔍 Found ${pmids.length} PMIDs in selected text: ${pmids.join(', ')}`);
                    
                    // Show loading message for PMID search
                    const responseArea = document.getElementById('responseArea');
                    responseArea.innerHTML = '<p class="placeholder-text">🔍 Fetching PubMed articles...</p>';
                    
                    try {
                        const pmidContext = await pubmedAPI.fetchMultipleArticles(pmids);
                        if (pmidContext) {
                            finalPrompt += pmidContext;
                            console.log('✅ Added PubMed context to prompt');
                        }
                    } catch (error) {
                        console.warn('Failed to fetch PMID context:', error);
                        // Continue without PMID context
                    }
                }
            }
        }
    }
    
    const replaceText = document.getElementById('replaceText').checked;
    await processPrompt(finalPrompt, replaceText);
}

// Core processing function
async function processPrompt(prompt, replaceSelection = false) {
    const defaultModelSelect = document.getElementById('defaultModel');
    const selectedModel = defaultModelSelect.value;
    
    if (!selectedModel) {
        showError('Please select a model first in Settings.');
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
    const savedPmidSearch = localStorage.getItem('ollama_pmid_search');
    const savedTargetLanguage = localStorage.getItem('ollama_target_language');
    
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
    
    if (savedPmidSearch !== null) {
        pmidSearchEnabled = savedPmidSearch === 'true';
        const pmidSearchElement = document.getElementById('enablePmidSearch');
        if (pmidSearchElement) {
            pmidSearchElement.checked = pmidSearchEnabled;
        }
    }
    
    if (savedTargetLanguage) {
        const targetLanguageElement = document.getElementById('targetLanguage');
        if (targetLanguageElement) {
            targetLanguageElement.value = savedTargetLanguage;
        }
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
    });
    
    const pmidSearchElement = document.getElementById('enablePmidSearch');
    if (pmidSearchElement) {
        pmidSearchElement.addEventListener('change', (e) => {
            pmidSearchEnabled = e.target.checked;
            localStorage.setItem('ollama_pmid_search', e.target.checked);
        });
    }
    
    const targetLanguageElement = document.getElementById('targetLanguage');
    if (targetLanguageElement) {
        targetLanguageElement.addEventListener('change', (e) => {
            localStorage.setItem('ollama_target_language', e.target.value);
        });
    }
}
