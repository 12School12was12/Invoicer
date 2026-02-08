// ============================================================
// KOOPLE — AI Service Integration v3.0
// Backend proxy mode: users don't need API keys
// Falls back to direct API (dev mode) or pattern matching
// ============================================================

const AIService = {

    // ── Configuration ────────────────────────────────────────
    // Backend proxy URL — set this to your server endpoint
    // The frontend sends requests here; the server holds the API key
    BACKEND_URL: '',  // e.g. 'https://your-worker.workers.dev/api/chat'

    _provider: null,   // 'openai' or 'anthropic' (direct mode only)
    _apiKey: null,     // only for direct mode (dev/advanced)

    MAX_HISTORY_MESSAGES: 20,

    // ── Is AI available? ────────────────────────────────────
    isConfigured() {
        return !!(this.BACKEND_URL || this._apiKey);
    },

    getConfig() {
        return {
            mode: this.BACKEND_URL ? 'backend' : (this._apiKey ? 'direct' : 'fallback'),
            provider: this._provider,
            hasKey: !!this._apiKey,
            hasBackend: !!this.BACKEND_URL
        };
    },

    // ── Direct API config (dev/advanced) ────────────────────
    configure(provider, apiKey) {
        if (provider !== 'openai' && provider !== 'anthropic') {
            throw new Error('Provider must be "openai" or "anthropic"');
        }
        this._provider = provider;
        this._apiKey = apiKey ? apiKey.trim() : null;
    },

    disconnect() {
        this._provider = null;
        this._apiKey = null;
        try { localStorage.removeItem('koople_ai_config'); } catch (e) {}
    },

    saveConfig() {
        if (!this._provider || !this._apiKey) return;
        try {
            localStorage.setItem('koople_ai_config', JSON.stringify({
                provider: this._provider,
                apiKey: this._apiKey
            }));
        } catch (e) {}
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
        } catch (e) {}
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
3. Tone: warm but professional, no judgment, moderate emoji use.
4. If you detect signs of abuse, violence, or serious mental health issues, recommend professional help.
5. ALWAYS respond in the language specified below.

IMPORTANT: Respond in {language}.

## User context
- User: {user_name}
- Partner: {partner_name}
- Together for: {duration}
- Living together: {living_together}
- Couple harmony score: {harmony_score}%
- Growth areas (low-scoring): {growth_areas}
- Strengths (high-scoring): {strengths}

## Questionnaire insights
{questionnaire_summary}

## Challenge generation
When you think a challenge would help, include a JSON block at the END of your response in this exact format:

\`\`\`challenge
{
  "title": "Short positive title (3-5 words)",
  "description": "What to do (2-3 sentences)",
  "category": "emotional|communication|household|intimacy|finances|quality_time|family|values|habits",
  "icon": "one emoji",
  "duration_days": 5,
  "difficulty": "easy|medium|hard",
  "assigned_to": "user|partner|both"
}
\`\`\`

Only include this block when a challenge is genuinely appropriate. Do NOT include it in every response.`,

    // ── Build system prompt with user context ────────────────
    _buildSystemPrompt(userContext) {
        const ctx = userContext || {};
        let prompt = this.MEDIATOR_PROMPT;

        const langNames = {
            ru: 'Russian', en: 'English', fr: 'French', de: 'German',
            es: 'Spanish', no: 'Norwegian', fi: 'Finnish', sv: 'Swedish', pt: 'Portuguese'
        };
        const langName = langNames[ctx.language] || ctx.language || 'English';

        prompt = prompt.replace('{language}', langName);
        prompt = prompt.replace('{user_name}', ctx.name || 'User');
        prompt = prompt.replace('{partner_name}', ctx.partnerName || 'Partner');
        prompt = prompt.replace('{duration}', ctx.duration || 'unknown');
        prompt = prompt.replace('{living_together}', ctx.livingTogether || 'unknown');
        prompt = prompt.replace('{harmony_score}', ctx.harmonyScore != null ? String(ctx.harmonyScore) : 'N/A');
        prompt = prompt.replace('{growth_areas}', ctx.growthAreas || 'not yet assessed');
        prompt = prompt.replace('{strengths}', ctx.strengths || 'not yet assessed');
        prompt = prompt.replace('{questionnaire_summary}', ctx.questionnaireSummary || 'No questionnaire data available yet.');

        return prompt;
    },

    // ── Build questionnaire summary from answers ─────────────
    buildQuestionnaireSummary(answers, catScores) {
        if (!answers || Object.keys(answers).length === 0) {
            return 'No questionnaire data available yet.';
        }
        const lines = [];
        if (catScores) {
            lines.push('Category scores (0-100):');
            const catNames = {
                emotional: 'Emotional connection', communication: 'Communication',
                household: 'Household/chores', intimacy: 'Intimacy',
                finances: 'Finances', quality_time: 'Quality time',
                family: 'Family & social', values: 'Values & goals', habits: 'Habits'
            };
            for (const [key, val] of Object.entries(catScores)) {
                lines.push('  ' + (catNames[key] || key) + ': ' + val + '%');
            }
            lines.push('');
        }
        const keyQ = {
            emo_love_language: 'Love language', emo_express: 'Expresses feelings (1-5)',
            comm_conflict_style: 'Conflict style', comm_unspoken: 'Unspoken topics',
            comm_after_fight: 'After arguments', house_chores: 'Chore satisfaction (1-5)',
            house_annoy: 'Household annoyances', intim_satisfaction: 'Intimacy satisfaction (1-5)',
            intim_frequency: 'Intimacy frequency match', fin_tension: 'Financial tension',
            time_together: 'Quality time (1-5)', hab_triggers: 'Triggers',
            exp_koople_goal: 'Goals for Koople'
        };
        lines.push('Key answers:');
        for (const [qId, label] of Object.entries(keyQ)) {
            if (answers[qId] != null) {
                lines.push('  ' + label + ': ' + (Array.isArray(answers[qId]) ? answers[qId].join(', ') : String(answers[qId])));
            }
        }
        return lines.join('\n');
    },

    // ── Convert messages to API format ───────────────────────
    _convertMessages(messages) {
        if (!messages || !messages.length) return [];
        return messages.slice(-this.MAX_HISTORY_MESSAGES).map(function (msg) {
            return {
                role: msg.role === 'assistant' ? 'assistant' : 'user',
                content: msg.text || ''
            };
        });
    },

    // ── Parse challenge from AI response ─────────────────────
    _parseChallengeFromResponse(text) {
        if (!text) return { cleanText: text, challenge: null };
        var regex = /```challenge\s*\n?([\s\S]*?)```/;
        var match = text.match(regex);
        if (!match) return { cleanText: text, challenge: null };

        var cleanText = text.replace(regex, '').trim();
        try {
            var p = JSON.parse(match[1].trim());
            if (!p.title || !p.description) return { cleanText: cleanText, challenge: null };
            return {
                cleanText: cleanText,
                challenge: {
                    id: 'ch_ai_' + Date.now(),
                    title: p.title,
                    description: p.description,
                    category: p.category || 'emotional',
                    icon: p.icon || '\uD83C\uDFAF',
                    duration: p.duration_days ? (p.duration_days === 1 ? 'One-time' : p.duration_days + ' days') : '7 days',
                    difficulty: ['easy', 'medium', 'hard'].includes(p.difficulty) ? p.difficulty : 'medium',
                    progress: 0,
                    total: p.duration_days || 7,
                    status: 'active',
                    assignedTo: p.assigned_to || 'both'
                }
            };
        } catch (e) {
            return { cleanText: cleanText, challenge: null };
        }
    },

    // ── Main chat method ─────────────────────────────────────
    async chat(messages, userContext) {
        var systemPrompt = this._buildSystemPrompt(userContext);
        var apiMessages = this._convertMessages(messages);

        try {
            var responseText;

            // Priority 1: Backend proxy (no API key needed)
            if (this.BACKEND_URL) {
                responseText = await this._callBackend(systemPrompt, apiMessages, userContext);
            }
            // Priority 2: Direct API call (dev mode)
            else if (this._apiKey && this._provider === 'openai') {
                responseText = await this._callOpenAI(systemPrompt, apiMessages);
            } else if (this._apiKey && this._provider === 'anthropic') {
                responseText = await this._callAnthropic(systemPrompt, apiMessages);
            } else {
                // No AI available — return null so app uses fallback
                return null;
            }

            var parsed = this._parseChallengeFromResponse(responseText);
            return { text: parsed.cleanText, challenge: parsed.challenge };

        } catch (error) {
            // On error, return error message but don't crash
            return { text: this._handleError(error), challenge: null };
        }
    },

    // ── Backend proxy call ───────────────────────────────────
    // Your backend receives the full context and returns AI response
    // Expected request: POST { system, messages, context }
    // Expected response: { text: "...", challenge?: {...} }
    async _callBackend(systemPrompt, messages, userContext) {
        var response = await fetch(this.BACKEND_URL, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                system: systemPrompt,
                messages: messages,
                context: userContext || {}
            })
        });

        if (!response.ok) {
            var err = new Error('Backend error');
            err.status = response.status;
            try { var d = await response.json(); err.detail = d.error || ''; } catch(e) {}
            throw err;
        }

        var data = await response.json();
        // Backend can return pre-parsed challenge or raw text
        if (data.challenge) {
            return data.text + '\n```challenge\n' + JSON.stringify(data.challenge) + '\n```';
        }
        return data.text;
    },

    // ── OpenAI API call (direct mode) ────────────────────────
    async _callOpenAI(systemPrompt, messages) {
        var response = await fetch('https://api.openai.com/v1/chat/completions', {
            method: 'POST',
            headers: {
                'Authorization': 'Bearer ' + this._apiKey,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                model: 'gpt-4o-mini',
                messages: [{ role: 'system', content: systemPrompt }].concat(messages),
                temperature: 0.7,
                max_tokens: 1000
            })
        });
        if (!response.ok) {
            var err = new Error('API error'); err.status = response.status;
            try { var d = await response.json(); err.detail = d.error ? d.error.message : ''; } catch(e) {}
            throw err;
        }
        var data = await response.json();
        return data.choices[0].message.content;
    },

    // ── Anthropic API call (direct mode) ─────────────────────
    async _callAnthropic(systemPrompt, messages) {
        var response = await fetch('https://api.anthropic.com/v1/messages', {
            method: 'POST',
            headers: {
                'x-api-key': this._apiKey,
                'anthropic-version': '2023-06-01',
                'Content-Type': 'application/json',
                'anthropic-dangerous-direct-browser-access': 'true'
            },
            body: JSON.stringify({
                model: 'claude-sonnet-4-20250514',
                system: systemPrompt,
                messages: messages,
                max_tokens: 1000
            })
        });
        if (!response.ok) {
            var err = new Error('API error'); err.status = response.status;
            try { var d = await response.json(); err.detail = d.error ? d.error.message : ''; } catch(e) {}
            throw err;
        }
        var data = await response.json();
        return data.content[0].text;
    },

    // ── Error handling ───────────────────────────────────────
    _handleError(error) {
        if (!error) return 'An unknown error occurred. Please try again.';
        if (error.name === 'TypeError' || error.message === 'Failed to fetch') {
            return 'Network error. Please check your internet connection.';
        }
        if (error.status === 401) return 'Invalid API key. Please check your settings.';
        if (error.status === 429) return 'Too many requests. Please try again in a minute.';
        if (error.status === 403) return 'Access denied. Check your API key permissions.';
        if (error.status === 400) return 'Bad request. ' + (error.detail || 'Please try again.');
        if (error.status >= 500) return 'AI service is temporarily unavailable. Please try later.';
        return 'An error occurred while contacting AI. Please try again.';
    }
};

// ── Auto-load saved config on script load ────────────────────
AIService.loadConfig();
