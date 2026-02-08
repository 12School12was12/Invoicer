// ============================================================
// KOOPLE — AI Service Integration
// Connects to OpenAI or Anthropic APIs for real AI mediator responses
// ============================================================

const AIService = {
    // ── Configuration ────────────────────────────────────────
    _provider: null, // 'openai' or 'anthropic'
    _apiKey: null,

    configure(provider, apiKey) {
        if (provider !== 'openai' && provider !== 'anthropic') {
            throw new Error('Provider must be "openai" or "anthropic"');
        }
        this._provider = provider;
        this._apiKey = apiKey ? apiKey.trim() : null;
    },

    isConfigured() {
        return !!this._apiKey;
    },

    getConfig() {
        return {
            provider: this._provider,
            hasKey: !!this._apiKey
        };
    },

    disconnect() {
        this._provider = null;
        this._apiKey = null;
        try {
            localStorage.removeItem('koople_ai_config');
        } catch (e) { /* storage unavailable */ }
    },

    saveConfig() {
        if (!this._provider || !this._apiKey) return;
        try {
            localStorage.setItem('koople_ai_config', JSON.stringify({
                provider: this._provider,
                apiKey: this._apiKey
            }));
        } catch (e) { /* storage unavailable */ }
    },

    loadConfig() {
        try {
            const raw = localStorage.getItem('koople_ai_config');
            if (!raw) return;
            const config = JSON.parse(raw);
            if (config && config.provider && config.apiKey) {
                this._provider = config.provider;
                this._apiKey = config.apiKey;
            }
        } catch (e) { /* corrupt or unavailable */ }
    },

    // ── System Prompt ────────────────────────────────────────

    MEDIATOR_PROMPT: `You are Koople, an AI mediator for couples. You help partners understand each other better, preventively resolve micro-conflicts, and gently encourage behavioral improvements.

Your role:
- Never take sides
- Never directly relay one partner's words to the other
- Transform complaints into constructive challenges
- Use positive reinforcement instead of criticism
- Understand that behind every complaint is an unmet need

Rules:
1. CONFIDENTIALITY: Each partner talks to you separately. Never reveal that one partner complained about the other.
2. Challenges should sound like general recommendations, not reactions to complaints.
3. Tone: warm but professional, no judgment, moderate emoji use, "ты" form (informal).
4. If you detect signs of abuse, violence, or serious mental health issues, recommend professional help.

IMPORTANT: Respond in the language specified in the user context. Current language: {language}

User context:
- User: {user_name}
- Partner: {partner_name}
- Together for: {duration}
- Living together: {living_together}
- Harmony score: {harmony_score}%
- Their concerns: {growth_areas}`,

    // ── Build system prompt with user context ────────────────

    _buildSystemPrompt(userContext) {
        const ctx = userContext || {};
        let prompt = this.MEDIATOR_PROMPT;

        prompt = prompt.replace('{language}', ctx.language || 'ru');
        prompt = prompt.replace('{user_name}', ctx.name || 'User');
        prompt = prompt.replace('{partner_name}', ctx.partnerName || 'Partner');
        prompt = prompt.replace('{duration}', ctx.duration || 'unknown');
        prompt = prompt.replace('{living_together}', ctx.livingTogether || 'unknown');
        prompt = prompt.replace('{harmony_score}', ctx.harmonyScore != null ? String(ctx.harmonyScore) : 'N/A');
        prompt = prompt.replace('{growth_areas}', ctx.growthAreas || 'not yet assessed');

        return prompt;
    },

    // ── Convert messages to API format ───────────────────────

    _convertMessages(messages) {
        if (!messages || !messages.length) return [];
        return messages.map(function (msg) {
            return {
                role: msg.role === 'assistant' ? 'assistant' : 'user',
                content: msg.text || ''
            };
        });
    },

    // ── Main chat method ─────────────────────────────────────

    async chat(messages, userContext) {
        if (!this.isConfigured()) {
            return {
                text: 'AI сервис не настроен. Укажите API ключ в настройках.',
                challenge: null
            };
        }

        var systemPrompt = this._buildSystemPrompt(userContext);
        var apiMessages = this._convertMessages(messages);

        try {
            var responseText;

            if (this._provider === 'openai') {
                responseText = await this._callOpenAI(systemPrompt, apiMessages);
            } else if (this._provider === 'anthropic') {
                responseText = await this._callAnthropic(systemPrompt, apiMessages);
            } else {
                return {
                    text: 'Неизвестный AI провайдер. Проверьте настройки.',
                    challenge: null
                };
            }

            return {
                text: responseText,
                challenge: null // TODO: parse AI response for challenge generation
            };
        } catch (error) {
            return {
                text: this._handleError(error),
                challenge: null
            };
        }
    },

    // ── OpenAI API call ──────────────────────────────────────

    async _callOpenAI(systemPrompt, messages) {
        var body = {
            model: 'gpt-4o-mini',
            messages: [{ role: 'system', content: systemPrompt }].concat(messages),
            temperature: 0.7,
            max_tokens: 800
        };

        var response = await fetch('https://api.openai.com/v1/chat/completions', {
            method: 'POST',
            headers: {
                'Authorization': 'Bearer ' + this._apiKey,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(body)
        });

        if (!response.ok) {
            var err = new Error('API error');
            err.status = response.status;
            throw err;
        }

        var data = await response.json();
        return data.choices[0].message.content;
    },

    // ── Anthropic API call ───────────────────────────────────

    async _callAnthropic(systemPrompt, messages) {
        var body = {
            model: 'claude-sonnet-4-20250514',
            system: systemPrompt,
            messages: messages,
            max_tokens: 800
        };

        var response = await fetch('https://api.anthropic.com/v1/messages', {
            method: 'POST',
            headers: {
                'x-api-key': this._apiKey,
                'anthropic-version': '2023-06-01',
                'Content-Type': 'application/json',
                'anthropic-dangerous-direct-browser-access': 'true'
            },
            body: JSON.stringify(body)
        });

        if (!response.ok) {
            var err = new Error('API error');
            err.status = response.status;
            throw err;
        }

        var data = await response.json();
        return data.content[0].text;
    },

    // ── Error handling ───────────────────────────────────────

    _handleError(error) {
        if (!error) {
            return 'Произошла неизвестная ошибка. Попробуйте ещё раз.';
        }

        // Network / fetch errors (no status code)
        if (error.name === 'TypeError' || error.message === 'Failed to fetch') {
            return 'Ошибка сети. Проверьте подключение к интернету.';
        }

        // HTTP status-based errors
        if (error.status === 401) {
            return 'Неверный API ключ. Проверьте настройки.';
        }

        if (error.status === 429) {
            return 'Слишком много запросов. Попробуйте через минуту.';
        }

        if (error.status === 403) {
            return 'Доступ запрещён. Проверьте разрешения API ключа.';
        }

        if (error.status === 500 || error.status === 502 || error.status === 503) {
            return 'Сервис временно недоступен. Попробуйте позже.';
        }

        return 'Произошла ошибка при обращении к AI. Попробуйте ещё раз.';
    }
};

// ── Auto-load saved config on script load ────────────────────
AIService.loadConfig();
