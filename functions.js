Office.onReady((info) => {
    if (info.host === Excel.HostType.Excel) {
        // Register the custom function
        Excel.Script.CustomFunctions.define("AZURE.ANALYZE", function(text, systemPrompt, temperature) {
            return new OfficeExtension.Promise(async function(resolve, reject) {
                try {
                    // Ensure we're in a valid Excel context
                    await Excel.run(async (context) => {
                        // Get Azure configuration from document settings
                        const settings = Office.context.document.settings;
                        await context.sync();
                        
                        const endpoint = settings.get('azureEndpoint');
                        const apiKey = settings.get('azureApiKey');

                        if (!endpoint || !apiKey) {
                            throw new Error('Azure configuration not found. Please configure the add-in first.');
                        }

                        // Input validation
                        if (!text || typeof text !== 'string') {
                            throw new Error('Invalid input: text must be a non-empty string');
                        }

                        // Set default values for optional parameters
                        systemPrompt = systemPrompt || '';
                        temperature = parseFloat(temperature) || 0.7;

                        if (temperature < 0 || temperature > 1) {
                            throw new Error('Temperature must be between 0 and 1');
                        }

                        // Prepare the request to Azure AI
                        const response = await fetch(`${endpoint}/openai/deployments/gpt-35-turbo/chat/completions?api-version=2023-05-15`, {
                            method: 'POST',
                            headers: {
                                'Content-Type': 'application/json',
                                'api-key': apiKey
                            },
                            body: JSON.stringify({
                                messages: [
                                    { role: 'system', content: systemPrompt },
                                    { role: 'user', content: text }
                                ],
                                temperature: temperature
                            })
                        });

                        if (!response.ok) {
                            const errorData = await response.json();
                            throw new Error(errorData.error?.message || 'Failed to get response from Azure AI');
                        }

                        const data = await response.json();
                        resolve(data.choices[0].message.content);
                    });
                } catch (error) {
                    reject(error.message || 'An unexpected error occurred');
                }
            });
        });
    }
});

// Metadata for the custom function
Excel.Script.CustomFunctions["AZURE.ANALYZE"].metadata = {
    parameters: [
        {
            name: "text",
            description: "The text content to analyze",
            type: "string",
            requiresAddress: false
        },
        {
            name: "systemPrompt",
            description: "Optional instructions for the AI",
            type: "string",
            optional: true,
            requiresAddress: false
        },
        {
            name: "temperature",
            description: "Controls response creativity (0-1)",
            type: "number",
            optional: true,
            requiresAddress: false
        }
    ],
    result: {
        type: "string",
        description: "The AI analysis result"
    },
    options: {
        stream: false,
        cancelable: true
    }
};