// ============================================================
// KOOPLE — Main Application Logic
// AI-Mediator for Couples
// ============================================================

const app = {
    // ── State ──────────────────────────────────────────────
    state: {
        currentScreen: 'welcome',
        previousScreen: null,
        user: {
            name: '',
            partnerName: '',
            email: '',
            duration: '',
            livingTogether: 'yes'
        },
        onboarding: {
            currentIndex: 0,
            answers: {},
            totalQuestions: QUESTIONNAIRE.length
        },
        chat: {
            messages: [...CHAT_WELCOME_MESSAGES],
            isTyping: false
        },
        challengeTab: 'active'
    },

    // ── Initialization ────────────────────────────────────
    init() {
        // Check if user has completed onboarding
        const saved = localStorage.getItem('koople_state');
        if (saved) {
            try {
                const parsed = JSON.parse(saved);
                Object.assign(this.state, parsed);
                if (this.state.user.name) {
                    this.navigate('dashboard');
                    return;
                }
            } catch (e) {
                // ignore
            }
        }
        this.navigate('welcome');
    },

    save() {
        localStorage.setItem('koople_state', JSON.stringify(this.state));
    },

    // ── Navigation ────────────────────────────────────────
    navigate(screenId) {
        const current = document.querySelector('.screen.active');
        const next = document.getElementById(`screen-${screenId}`);
        if (!next || next === current) return;

        this.state.previousScreen = this.state.currentScreen;
        this.state.currentScreen = screenId;

        if (current) {
            current.classList.add('exit');
            current.classList.remove('active');
            setTimeout(() => current.classList.remove('exit'), 400);
        }

        next.classList.add('active');

        // Update bottom nav highlights
        document.querySelectorAll('.bottom-nav .nav-item').forEach(item => {
            item.classList.remove('active');
        });

        // Screen-specific setup
        switch (screenId) {
            case 'dashboard':
                this.renderDashboard();
                break;
            case 'challenges':
                this.renderChallenges();
                break;
            case 'mediator':
                this.renderChat();
                break;
            case 'insights':
                this.renderInsights();
                break;
            case 'wishes':
                this.renderWishes();
                break;
            case 'onboarding':
                this.renderQuestion();
                break;
            case 'profile':
                this.renderProfile();
                break;
        }

        // Update active nav item for screens with bottom nav
        const navScreens = next.querySelectorAll('.bottom-nav .nav-item');
        navScreens.forEach(item => {
            const onclick = item.getAttribute('onclick') || '';
            if (onclick.includes(`'${screenId}'`)) {
                item.classList.add('active');
            }
        });
    },

    // ── Welcome & Auth ────────────────────────────────────
    startOnboarding() {
        this.navigate('setup');
    },

    showLogin() {
        this.navigate('login');
    },

    login() {
        // Demo: just go to dashboard
        this.state.user.name = 'Аня';
        this.state.user.partnerName = 'Миша';
        this.save();
        this.navigate('dashboard');
    },

    saveProfile() {
        this.state.user.name = document.getElementById('setup-name').value;
        this.state.user.partnerName = document.getElementById('setup-partner').value;
        this.state.user.duration = document.getElementById('setup-duration').value;
        this.state.user.livingTogether = document.getElementById('setup-living').value;
        this.state.user.email = document.getElementById('setup-email').value;
        this.save();
        this.state.onboarding.currentIndex = 0;
        this.navigate('onboarding');
    },

    toggleSelect(btn, hiddenId) {
        const group = btn.parentElement;
        group.querySelectorAll('.toggle-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        document.getElementById(hiddenId).value = btn.dataset.value;
    },

    // ── Onboarding Questionnaire ──────────────────────────
    renderQuestion() {
        const idx = this.state.onboarding.currentIndex;
        const q = QUESTIONNAIRE[idx];
        if (!q) return;

        const total = QUESTIONNAIRE.length;
        const progress = ((idx + 1) / total) * 100;

        document.getElementById('onboarding-progress').style.width = progress + '%';
        document.getElementById('onboarding-progress-text').textContent = `${idx + 1} / ${total}`;
        document.getElementById('category-icon').textContent = q.categoryIcon;
        document.getElementById('category-name').textContent = q.categoryName;
        document.getElementById('question-text').textContent = q.text;
        document.getElementById('question-hint').textContent = q.hint || '';

        const area = document.getElementById('answer-area');
        area.innerHTML = '';

        const existingAnswer = this.state.onboarding.answers[q.id];

        switch (q.type) {
            case 'single':
                q.options.forEach(opt => {
                    const btn = document.createElement('button');
                    btn.className = 'option-btn';
                    if (existingAnswer === opt.value) btn.classList.add('selected');
                    btn.textContent = opt.label;
                    btn.onclick = () => {
                        area.querySelectorAll('.option-btn').forEach(b => b.classList.remove('selected'));
                        btn.classList.add('selected');
                        this.state.onboarding.answers[q.id] = opt.value;
                    };
                    area.appendChild(btn);
                });
                break;

            case 'multi':
                const selected = existingAnswer || [];
                q.options.forEach(opt => {
                    const btn = document.createElement('button');
                    btn.className = 'option-btn multi';
                    if (selected.includes(opt.value)) btn.classList.add('selected');
                    btn.innerHTML = `<span class="check-icon"></span>${opt.label}`;
                    btn.onclick = () => {
                        let current = this.state.onboarding.answers[q.id] || [];
                        if (current.includes(opt.value)) {
                            current = current.filter(v => v !== opt.value);
                            btn.classList.remove('selected');
                        } else {
                            if (q.maxSelect && current.length >= q.maxSelect) return;
                            current.push(opt.value);
                            btn.classList.add('selected');
                        }
                        this.state.onboarding.answers[q.id] = current;
                    };
                    area.appendChild(btn);
                });
                break;

            case 'scale':
                const scaleContainer = document.createElement('div');
                scaleContainer.className = 'scale-container';
                q.scaleLabels.forEach((label, i) => {
                    const val = i + 1;
                    const item = document.createElement('button');
                    item.className = 'scale-item';
                    if (existingAnswer === val) item.classList.add('selected');
                    item.innerHTML = `<span class="scale-number">${val}</span><span class="scale-label">${label}</span>`;
                    item.onclick = () => {
                        scaleContainer.querySelectorAll('.scale-item').forEach(s => s.classList.remove('selected'));
                        item.classList.add('selected');
                        this.state.onboarding.answers[q.id] = val;
                    };
                    scaleContainer.appendChild(item);
                });
                area.appendChild(scaleContainer);
                break;

            case 'textarea':
                const ta = document.createElement('textarea');
                ta.className = 'answer-textarea';
                ta.placeholder = q.placeholder || 'Ваш ответ...';
                ta.rows = 4;
                ta.value = existingAnswer || '';
                ta.oninput = () => {
                    this.state.onboarding.answers[q.id] = ta.value;
                };
                area.appendChild(ta);
                break;
        }

        // Animate in
        const container = document.getElementById('question-container');
        container.classList.remove('slide-in');
        void container.offsetWidth;
        container.classList.add('slide-in');
    },

    nextQuestion() {
        if (this.state.onboarding.currentIndex < QUESTIONNAIRE.length - 1) {
            this.state.onboarding.currentIndex++;
            this.renderQuestion();
            this.save();
        } else {
            this.completeOnboarding();
        }
    },

    prevQuestion() {
        if (this.state.onboarding.currentIndex > 0) {
            this.state.onboarding.currentIndex--;
            this.renderQuestion();
        } else {
            this.navigate('setup');
        }
    },

    skipQuestion() {
        this.nextQuestion();
    },

    completeOnboarding() {
        this.save();
        this.navigate('onboarding-complete');
    },

    copyInviteCode() {
        const code = document.getElementById('invite-code').textContent;
        navigator.clipboard.writeText(code).then(() => {
            const btn = document.querySelector('.btn-copy');
            btn.textContent = 'Скопировано!';
            setTimeout(() => btn.textContent = 'Копировать', 2000);
        }).catch(() => {
            // fallback
        });
    },

    shareInvite() {
        const code = document.getElementById('invite-code').textContent;
        const text = `Присоединяйся ко мне в Koople — AI-медиаторе для пар! Мой код: ${code}`;
        if (navigator.share) {
            navigator.share({ title: 'Koople', text });
        } else {
            this.copyInviteCode();
        }
    },

    goToDashboard() {
        this.navigate('dashboard');
    },

    // ── Dashboard ─────────────────────────────────────────
    renderDashboard() {
        const name = this.state.user.name || 'Аня';
        document.getElementById('greeting').textContent = `Привет, ${name}!`;
        document.getElementById('user-avatar').textContent = name.charAt(0).toUpperCase();

        // Render active challenges
        const challengesEl = document.getElementById('active-challenges');
        const activeChallenges = DEMO_CHALLENGES.filter(c => c.status === 'active');
        challengesEl.innerHTML = activeChallenges.map(ch => `
            <div class="challenge-card-mini" onclick="app.showChallengeDetail('${ch.id}')">
                <div class="challenge-mini-icon">${ch.icon}</div>
                <div class="challenge-mini-info">
                    <h4>${ch.title}</h4>
                    <div class="challenge-mini-progress">
                        <div class="mini-progress-bar">
                            <div class="mini-progress-fill" style="width:${(ch.progress/ch.total)*100}%"></div>
                        </div>
                        <span>${ch.progress}/${ch.total}</span>
                    </div>
                </div>
            </div>
        `).join('');

        // Render activity feed
        const feedEl = document.getElementById('activity-feed');
        feedEl.innerHTML = DEMO_ACTIVITY.map(a => `
            <div class="activity-item">
                <span class="activity-icon">${a.icon}</span>
                <div class="activity-content">
                    <p>${a.text}</p>
                    <span class="activity-time">${a.time}</span>
                </div>
            </div>
        `).join('');
    },

    // ── Challenges ────────────────────────────────────────
    renderChallenges() {
        this.switchChallengeTab(this.state.challengeTab);
    },

    switchChallengeTab(tab, clickedBtn) {
        this.state.challengeTab = tab;

        // Update tabs
        if (clickedBtn) {
            document.querySelectorAll('#screen-challenges .tab').forEach(t => t.classList.remove('active'));
            clickedBtn.classList.add('active');
        }

        const list = document.getElementById('challenges-list');
        const filtered = DEMO_CHALLENGES.filter(c => c.status === tab);

        if (filtered.length === 0) {
            list.innerHTML = `
                <div class="empty-state">
                    <span class="empty-icon">${tab === 'completed' ? '\uD83C\uDFC6' : '\uD83C\uDF31'}</span>
                    <p>${tab === 'completed' ? 'Пока нет завершённых челленджей' : 'Новые челленджи появятся после анализа'}</p>
                </div>
            `;
            return;
        }

        list.innerHTML = filtered.map(ch => {
            const progressPct = (ch.progress / ch.total) * 100;
            const difficultyLabel = { easy: '\u{1F7E2} Лёгкий', medium: '\u{1F7E1} Средний', hard: '\u{1F7E0} Сложный' }[ch.difficulty];
            const assignedLabel = { both: 'Для обоих', user: 'Для вас', partner: `Для ${this.state.user.partnerName || 'партнёра'}` }[ch.assignedTo];

            return `
                <div class="challenge-card" onclick="app.showChallengeDetail('${ch.id}')">
                    <div class="challenge-card-header">
                        <span class="challenge-icon-large">${ch.icon}</span>
                        <div>
                            <h3>${ch.title}</h3>
                            <div class="challenge-meta">
                                <span>${difficultyLabel}</span>
                                <span>\u00B7</span>
                                <span>${ch.duration}</span>
                                <span>\u00B7</span>
                                <span>${assignedLabel}</span>
                            </div>
                        </div>
                    </div>
                    <p class="challenge-desc">${ch.description}</p>
                    ${ch.status !== 'suggested' ? `
                        <div class="challenge-progress">
                            <div class="progress-bar">
                                <div class="progress-fill ${ch.status === 'completed' ? 'complete' : ''}" style="width:${progressPct}%"></div>
                            </div>
                            <span class="progress-label">${ch.progress} из ${ch.total}</span>
                        </div>
                    ` : `
                        <button class="btn btn-secondary btn-small" onclick="event.stopPropagation(); app.acceptChallenge('${ch.id}')">Принять челлендж</button>
                    `}
                </div>
            `;
        }).join('');
    },

    showChallengeDetail(id) {
        const ch = DEMO_CHALLENGES.find(c => c.id === id);
        if (!ch) return;

        const progressPct = (ch.progress / ch.total) * 100;
        this.showModal(`
            <div class="challenge-detail">
                <div class="challenge-detail-icon">${ch.icon}</div>
                <h2>${ch.title}</h2>
                <p>${ch.description}</p>
                <div class="challenge-detail-meta">
                    <div class="meta-item">
                        <span class="meta-label">Длительность</span>
                        <span class="meta-value">${ch.duration}</span>
                    </div>
                    <div class="meta-item">
                        <span class="meta-label">Прогресс</span>
                        <span class="meta-value">${ch.progress}/${ch.total}</span>
                    </div>
                </div>
                <div class="challenge-progress" style="margin-top:16px">
                    <div class="progress-bar">
                        <div class="progress-fill ${ch.status === 'completed' ? 'complete' : ''}" style="width:${progressPct}%"></div>
                    </div>
                </div>
                ${ch.status === 'active' ? `
                    <button class="btn btn-primary btn-large" style="margin-top:20px" onclick="app.markChallengeDay('${ch.id}')">
                        \u2705 Отметить выполнение
                    </button>
                ` : ''}
                ${ch.status === 'suggested' ? `
                    <button class="btn btn-primary btn-large" style="margin-top:20px" onclick="app.acceptChallenge('${ch.id}')">
                        Принять челлендж
                    </button>
                ` : ''}
            </div>
        `);
    },

    markChallengeDay(id) {
        const ch = DEMO_CHALLENGES.find(c => c.id === id);
        if (ch && ch.progress < ch.total) {
            ch.progress++;
            if (ch.progress >= ch.total) {
                ch.status = 'completed';
            }
        }
        this.closeModal();
        this.renderChallenges();
    },

    acceptChallenge(id) {
        const ch = DEMO_CHALLENGES.find(c => c.id === id);
        if (ch) {
            ch.status = 'active';
        }
        this.closeModal();
        this.renderChallenges();
    },

    // ── AI Mediator Chat ──────────────────────────────────
    renderChat() {
        const messagesEl = document.getElementById('chat-messages');
        messagesEl.innerHTML = this.state.chat.messages.map(msg => `
            <div class="chat-message ${msg.role}">
                ${msg.role === 'assistant' ? '<div class="msg-avatar"><svg viewBox="0 0 32 32" fill="none"><circle cx="16" cy="16" r="14" fill="var(--primary-light)"/><circle cx="11" cy="13" r="1.5" fill="var(--primary)"/><circle cx="21" cy="13" r="1.5" fill="var(--primary)"/><path d="M11 20 Q16 24 21 20" stroke="var(--primary)" stroke-width="1.5" fill="none" stroke-linecap="round"/></svg></div>' : ''}
                <div class="msg-bubble">
                    <p>${msg.text.replace(/\n/g, '<br>')}</p>
                    <span class="msg-time">${msg.time}</span>
                </div>
            </div>
        `).join('');

        if (this.state.chat.isTyping) {
            messagesEl.innerHTML += `
                <div class="chat-message assistant">
                    <div class="msg-avatar"><svg viewBox="0 0 32 32" fill="none"><circle cx="16" cy="16" r="14" fill="var(--primary-light)"/><circle cx="11" cy="13" r="1.5" fill="var(--primary)"/><circle cx="21" cy="13" r="1.5" fill="var(--primary)"/><path d="M11 20 Q16 24 21 20" stroke="var(--primary)" stroke-width="1.5" fill="none" stroke-linecap="round"/></svg></div>
                    <div class="msg-bubble typing">
                        <span class="typing-dot"></span>
                        <span class="typing-dot"></span>
                        <span class="typing-dot"></span>
                    </div>
                </div>
            `;
        }

        messagesEl.scrollTop = messagesEl.scrollHeight;
    },

    sendMessage() {
        const input = document.getElementById('chat-input');
        const text = input.value.trim();
        if (!text) return;

        this.state.chat.messages.push({
            role: 'user',
            text,
            time: this.getCurrentTime()
        });

        input.value = '';
        input.style.height = 'auto';
        this.renderChat();

        // Hide suggestions after first message
        document.getElementById('chat-suggestions').style.display = 'none';

        // Simulate AI response
        this.state.chat.isTyping = true;
        this.renderChat();

        setTimeout(() => {
            this.state.chat.isTyping = false;
            const response = this.generateAIResponse(text);
            this.state.chat.messages.push({
                role: 'assistant',
                text: response,
                time: this.getCurrentTime()
            });
            this.renderChat();
            this.save();
        }, 1500 + Math.random() * 1500);
    },

    sendSuggestion(btn) {
        document.getElementById('chat-input').value = btn.textContent;
        this.sendMessage();
    },

    generateAIResponse(userMessage) {
        const lower = userMessage.toLowerCase();
        const partnerName = this.state.user.partnerName || 'партнёр';

        if (lower.includes('беспокоит') || lower.includes('раздражает') || lower.includes('бесит') || lower.includes('привычк')) {
            return `Я понимаю, что это может вызывать раздражение. Спасибо, что поделились — это важный первый шаг.\n\nДавайте разберёмся: как давно это происходит и как часто? Это поможет мне подобрать правильный подход.\n\nВажно: я не буду передавать ${partnerName} ваши слова напрямую. Вместо этого я создам мягкий челлендж, который поможет улучшить ситуацию естественным образом.`;
        }

        if (lower.includes('отношени') || lower.includes('обсудить') || lower.includes('поговорить')) {
            return `Конечно, давайте поговорим. Я здесь, чтобы помочь вам разобраться в чувствах.\n\nМожете начать с того, что вас волнует больше всего прямо сейчас? Не обязательно формулировать идеально — просто расскажите как есть.\n\nЯ замечу паттерны и помогу увидеть ситуацию с разных сторон.`;
        }

        if (lower.includes('сложно сказать') || lower.includes('не могу сказать') || lower.includes('стесня')) {
            return `Это нормально — многие вещи сложно произнести вслух. Именно для этого я здесь.\n\nВы можете написать мне всё, что хотели бы сказать ${partnerName}, но не решаетесь. Я помогу:\n\n1. Разобраться в ваших чувствах\n2. Найти правильные слова\n3. Или передать пожелание через мягкий челлендж\n\nЧто бы вы хотели сказать?`;
        }

        if (lower.includes('ванн') || lower.includes('волос') || lower.includes('чистот') || lower.includes('убира')) {
            return `Бытовые вещи — одна из самых частых причин микроконфликтов в парах. Это нормально!\n\nЯ создам челлендж для ${partnerName}: «Неделя чистоты в ванной» — с конкретными простыми действиями после каждого использования.\n\nПри этом ${partnerName} получит это как общую рекомендацию для улучшения совместной жизни, а не как жалобу от вас. Так это воспринимается гораздо легче.\n\nХотите, чтобы я запустил этот челлендж?`;
        }

        if (lower.includes('спасибо') || lower.includes('да') || lower.includes('запус')) {
            return `Отлично! Я подготовлю челлендж и аккуратно предложу его ${partnerName}.\n\nВот что произойдёт:\n\u2022 ${partnerName} получит рекомендацию в мягкой форме\n\u2022 Челлендж будет рассчитан на 7 дней\n\u2022 Вы сможете отслеживать прогресс\n\u2022 В конце оба оцените результат\n\nМеждо тем, есть ли ещё что-то, о чём вы хотели бы поговорить?`;
        }

        return `Спасибо, что делитесь этим со мной. Я анализирую ваши ответы вместе с данными анкеты, чтобы найти лучший способ помочь.\n\nМогу предложить несколько вариантов:\n\n1. \uD83D\uDCAC Обсудить ситуацию подробнее\n2. \uD83C\uDFAF Создать мягкий челлендж для партнёра\n3. \uD83D\uDCA1 Дать рекомендацию для вас обоих\n\nЧто вам ближе?`;
    },

    autoGrow(el) {
        el.style.height = 'auto';
        el.style.height = Math.min(el.scrollHeight, 120) + 'px';
    },

    getCurrentTime() {
        const now = new Date();
        return now.getHours().toString().padStart(2, '0') + ':' + now.getMinutes().toString().padStart(2, '0');
    },

    // ── Insights ──────────────────────────────────────────
    renderInsights() {
        // Compatibility bars
        const barsEl = document.getElementById('compatibility-bars');
        barsEl.innerHTML = COMPATIBILITY_DATA.map(item => `
            <div class="compat-row">
                <span class="compat-name">${item.name}</span>
                <div class="compat-bar-track">
                    <div class="compat-bar-fill" style="width:${item.value}%; background:${item.color}"></div>
                </div>
                <span class="compat-value">${item.value}%</span>
            </div>
        `).join('');

        // Progress chart (simplified bar chart)
        const chartEl = document.getElementById('progress-chart');
        const weeks = ['Нед 1', 'Нед 2', 'Нед 3', 'Нед 4'];
        const values = [65, 70, 74, 78];
        chartEl.innerHTML = `
            <div class="bar-chart">
                ${weeks.map((w, i) => `
                    <div class="bar-col">
                        <div class="bar" style="height:${values[i]}%">
                            <span class="bar-value">${values[i]}</span>
                        </div>
                        <span class="bar-label">${w}</span>
                    </div>
                `).join('')}
            </div>
        `;

        // AI Recommendations
        const recsEl = document.getElementById('ai-recommendations');
        recsEl.innerHTML = AI_RECOMMENDATIONS.map(rec => `
            <div class="recommendation-card priority-${rec.priority}">
                <h4>${rec.title}</h4>
                <p>${rec.text}</p>
            </div>
        `).join('');
    },

    // ── Wishes ────────────────────────────────────────────
    renderWishes() {
        const listEl = document.getElementById('wishes-list');
        listEl.innerHTML = DEMO_WISHES.map(w => {
            const statusLabel = { active: 'Активно', in_progress: 'AI работает над этим', fulfilled: 'Исполнено' }[w.status];
            const statusClass = w.status;
            return `
                <div class="wish-card">
                    <p class="wish-text">${w.text}</p>
                    <div class="wish-meta">
                        <span class="wish-status ${statusClass}">${statusLabel}</span>
                        <span class="wish-date">${w.createdAt}</span>
                    </div>
                </div>
            `;
        }).join('');
    },

    showAddWishModal() {
        this.showModal(`
            <h2>Новое желание</h2>
            <p class="modal-subtitle">Расскажите, чего бы вы хотели от партнёра. AI аккуратно поработает над этим.</p>
            <textarea id="wish-input" class="answer-textarea" placeholder="Например: хочу, чтобы чаще обнимал/-а меня без повода..." rows="4"></textarea>
            <div class="modal-category-select">
                <label>Категория:</label>
                <select id="wish-category">
                    <option value="emotional">Эмоции</option>
                    <option value="household">Быт</option>
                    <option value="quality_time">Время вместе</option>
                    <option value="intimacy">Близость</option>
                    <option value="communication">Общение</option>
                </select>
            </div>
            <button class="btn btn-primary btn-large" onclick="app.addWish()">Отправить желание</button>
        `);
    },

    addWish() {
        const text = document.getElementById('wish-input').value.trim();
        const category = document.getElementById('wish-category').value;
        if (!text) return;

        DEMO_WISHES.unshift({
            id: 'w' + Date.now(),
            text,
            category,
            status: 'active',
            createdAt: 'Только что'
        });

        this.closeModal();
        this.renderWishes();
    },

    // ── Profile ───────────────────────────────────────────
    renderProfile() {
        const name = this.state.user.name || 'Аня';
        const partnerName = this.state.user.partnerName || 'Миша';
        document.getElementById('profile-name').textContent = name;
        document.getElementById('profile-partner-name').textContent = partnerName;
        document.querySelector('.profile-avatar-large').textContent = name.charAt(0).toUpperCase();
    },

    showInvitePartner() {
        this.showModal(`
            <h2>Пригласить партнёра</h2>
            <p>Отправьте этот код партнёру, чтобы он/она смог присоединиться к вашей паре</p>
            <div class="invite-code-box">
                <span class="invite-code">KOOPLE-A7X9</span>
                <button class="btn-copy" onclick="app.copyInviteCode()">Копировать</button>
            </div>
            <button class="btn btn-secondary btn-large" onclick="app.shareInvite()">Отправить приглашение</button>
        `);
    },

    showAbout() {
        this.showModal(`
            <div style="text-align:center">
                <h2>Koople</h2>
                <p>AI-медиатор для пар</p>
                <p style="color:var(--text-secondary);margin-top:12px">Версия 1.0 MVP</p>
                <p style="color:var(--text-secondary);margin-top:8px">Помогаем парам лучше понимать друг друга через AI-анализ и мягкие челленджи</p>
            </div>
        `);
    },

    // ── Complaint & Appreciation Modals ───────────────────
    showComplaintModal() {
        this.showModal(`
            <h2>\uD83D\uDE14 Что вас беспокоит?</h2>
            <p class="modal-subtitle">Опишите ситуацию — AI-медиатор деликатно поможет решить её</p>
            <textarea id="complaint-input" class="answer-textarea" placeholder="Например: партнёр оставляет грязную посуду в раковине..." rows="4"></textarea>
            <div class="complaint-urgency">
                <label>Насколько это важно?</label>
                <div class="urgency-options">
                    <button class="urgency-btn" onclick="app.selectUrgency(this, 'low')">Мелочь</button>
                    <button class="urgency-btn active" onclick="app.selectUrgency(this, 'medium')">Беспокоит</button>
                    <button class="urgency-btn" onclick="app.selectUrgency(this, 'high')">Серьёзно</button>
                </div>
            </div>
            <button class="btn btn-primary btn-large" onclick="app.submitComplaint()">Отправить медиатору</button>
        `);
    },

    selectUrgency(btn, level) {
        document.querySelectorAll('.urgency-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
    },

    submitComplaint() {
        const text = document.getElementById('complaint-input').value.trim();
        if (!text) return;
        this.closeModal();
        // Navigate to chat with the complaint
        this.state.chat.messages.push({
            role: 'user',
            text: `Меня беспокоит: ${text}`,
            time: this.getCurrentTime()
        });
        this.navigate('mediator');
        // Trigger AI response
        this.state.chat.isTyping = true;
        this.renderChat();
        setTimeout(() => {
            this.state.chat.isTyping = false;
            this.state.chat.messages.push({
                role: 'assistant',
                text: this.generateAIResponse(text),
                time: this.getCurrentTime()
            });
            this.renderChat();
        }, 2000);
    },

    showAppreciationModal() {
        const partnerName = this.state.user.partnerName || 'партнёра';
        this.showModal(`
            <h2>\uD83D\uDC95 Похвалить ${partnerName}</h2>
            <p class="modal-subtitle">Позитивная обратная связь укрепляет отношения</p>
            <textarea id="appreciation-input" class="answer-textarea" placeholder="За что вы хотите поблагодарить или похвалить партнёра?" rows="4"></textarea>
            <button class="btn btn-primary btn-large" onclick="app.submitAppreciation()">Отправить</button>
        `);
    },

    submitAppreciation() {
        const text = document.getElementById('appreciation-input').value.trim();
        if (!text) return;
        this.closeModal();
        // Show success notification
        this.showNotification('\uD83D\uDC95 Отправлено! Партнёр получит приятное уведомление');
    },

    // ── Modals ────────────────────────────────────────────
    showModal(content) {
        const overlay = document.getElementById('modal-overlay');
        const modal = document.getElementById('modal-content');
        modal.innerHTML = `<button class="modal-close" onclick="app.closeModal()">&times;</button>${content}`;
        overlay.classList.add('active');
        document.body.style.overflow = 'hidden';
    },

    closeModal(event) {
        if (event && event.target !== event.currentTarget) return;
        const overlay = document.getElementById('modal-overlay');
        overlay.classList.remove('active');
        document.body.style.overflow = '';
    },

    // ── Notifications ─────────────────────────────────────
    showNotification(text) {
        const notif = document.createElement('div');
        notif.className = 'notification';
        notif.textContent = text;
        document.body.appendChild(notif);
        setTimeout(() => notif.classList.add('show'), 10);
        setTimeout(() => {
            notif.classList.remove('show');
            setTimeout(() => notif.remove(), 300);
        }, 3000);
    }
};

// ── Initialize ──────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => app.init());

// Handle Enter key in chat
document.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && !e.shiftKey && document.activeElement.id === 'chat-input') {
        e.preventDefault();
        app.sendMessage();
    }
});
