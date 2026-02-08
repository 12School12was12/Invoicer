// ============================================================
// KOOPLE — AI Service Integration v2.0
// Connects to OpenAI or Anthropic APIs for real AI mediator responses
// Enhanced: full context, challenge parsing, history limits
// ============================================================

const AIService = {

        // ── Configuration ────────────────────────────────────────
        _provider: null, // 'openai' or 'anthropic'
        _apiKey: null,

        MAX_HISTORY_MESSAGES: 20, // limit context window

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
                    return { provider: this._provider, hasKey: !!this._apiKey };
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

        // ── Build system prompt with full user context ────────────
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

            // Category scores
            if (catScores) {
                            lines.push('Category scores (0-100):');
                            const catNames = {
                                                emotional: 'Emotional connection',
                                                communication: 'Communication',
                                                household: 'Household/chores',
                                                intimacy: 'Intimacy',
                                                finances: 'Finances',
                                                quality_time: 'Quality time',
                                                family: 'Family & social',
                                                values: 'Values & goals',
                                                habits: 'Habits'
                            };
                            for (const [key, val] of Object.entries(catScores)) {
                                                const name = catNames[key] || key;
                                                lines.push('  ' + name + ': ' + val + '%');
                            }
                            lines.push('');
            }

            // Key answers (selected important ones)
            const keyQuestions = {
                            emo_love_language: 'Love language',
                            emo_express: 'Expresses feelings (1-5 scale)',
                            comm_conflict_style: 'Conflict style',
                            comm_unspoken: 'Unspoken topics',
                            comm_after_fight: 'After arguments',
                            house_chores: 'Chore satisfaction (1-5)',
                            house_annoy: 'Household annoyances',
                            intim_satisfaction: 'Intimacy satisfaction (1-5)',
                            intim_frequency: 'Intimacy frequency match',
                            fin_tension: 'Financial tension frequency',
                            time_together: 'Quality time satisfaction (1-5)',
                            time_activities: 'Missing shared activities',
                            fam_interference: 'Family interference',
                            val_future: 'Aligned on future (1-5)',
                            hab_triggers: 'Partner habit triggers',
                            expectations: 'What user wants from Koople'
            };

            lines.push('Key questionnaire answers:');
                    for (const [qId, label] of Object.entries(keyQuestions)) {
                                    if (answers[qId] !== undefined && answers[qId] !== null) {
                                                        const val = Array.isArray(answers[qId]) ? answers[qId].join(', ') : String(answers[qId]);
                                                        lines.push('  ' + label + ': ' + val);
                                    }
                    }

            return lines.join('\n');
        },

        // ── Convert messages to API format (with history limit) ──
        _convertMessages(messages) {
                    if (!messages || !messages.length) return [];
                    // Take only the last N messages to stay within context limits
            const limited = messages.slice(-this.MAX_HISTORY_MESSAGES);
                    return limited.map(function (msg) {
                                    return {
                                                        role: msg.role === 'assistant' ? 'assistant' : 'user',
                                                        content: msg.text || ''
                                    };
                    });
        },

        // ── Parse challenge from AI response ─────────────────────
        _parseChallengeFromResponse(text) {
                    if (!text) return { cleanText: text, challenge: null };

            const challengeRegex = /```challenge\s*\n?([\s\S]*?)```/;
                    const match = text.match(challengeRegex);

            if (!match) {
                            return { cleanText: text, challenge: null };
            }

            // Remove the challenge block from visible text
            const cleanText = text.replace(challengeRegex, '').trim();

            try {
                            const parsed = JSON.parse(match[1].trim());

                        // Validate required fields
                        if (!parsed.title || !parsed.description) {
                                            return { cleanText, challenge: null };
                        }

                        const challenge = {
                                            id: 'ch_ai_' + Date.now(),
                                            title: parsed.title,
                                            description: parsed.description,
                                            category: parsed.category || 'emotional',
                                            icon: parsed.icon || '\uD83C\uDFAF',
                                            duration: parsed.duration_days ? (parsed.duration_days === 1 ? 'One-time' : parsed.duration_days + ' days') : '7 days',
                                            difficulty: ['easy', 'medium', 'hard'].includes(parsed.difficulty) ? parsed.difficulty : 'medium',
                                            progress: 0,
                                            total: parsed.duration_days || 7,
                                            status: 'active',
                                            assignedTo: parsed.assigned_to || 'both'
                        };

                        return { cleanText, challenge };
            } catch (e) {
                            // JSON parse failed — return text without the broken block
                        return { cleanText, challenge: null };
            }
        },

        // ── Main chat method ─────────────────────────────────────
        async chat(messages, userContext) {
                    if (!this.isConfigured()) {
                                    return {
                                                        text: 'AI service is not configured. Please add your API key in Settings.',
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
                                                return { text: 'Unknown AI provider. Check your settings.', challenge: null };
                            }

                        // Parse challenge from response if present
                        var parsed = this._parseChallengeFromResponse(responseText);

                        return {
                                            text: parsed.cleanText,
                                            challenge: parsed.challenge
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
                                    max_tokens: 1000
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
                            try {
                                                var errData = await response.json();
                                                err.detail = errData.error ? errData.error.message : '';
                            } catch(e) {}
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
                                    max_tokens: 1000
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
                            try {
                                                var errData = await response.json();
                                                err.detail = errData.error ? errData.error.message : '';
                            } catch(e) {}
                            throw err;
            }

            var data = await response.json();
                    return data.content[0].text;
        },

        // ── Error handling ───────────────────────────────────────
        _handleError(error) {
                    if (!error) {
                                    return 'An unknown error occurred. Please try again.';
                    }
                    if (error.name === 'TypeError' || error.message === 'Failed to fetch') {
                                    return 'Network error. Please check your internet connection.';
                    }
                    if (error.status === 401) {
                                    return 'Invalid API key. Please check your settings.';
                    }
                    if (error.status === 429) {
                                    return 'Too many requests. Please try again in a minute.';
                    }
                    if (error.status === 403) {
                                    return 'Access denied. Check your API key permissions.';
                    }
                    if (error.status === 400) {
                                    return 'Bad request. ' + (error.detail || 'Please try again.');
                    }
                    if (error.status >= 500) {
                                    return 'AI service is temporarily unavailable. Please try later.';
                    }
                    return 'An error occurred while contacting AI. Please try again.';
        }
};

// ── Auto-load saved config on script load ────────────────────
AIService.loadConfig();
