// ============================================================
// KOOPLE — Main Application Logic v2.0
// AI-Mediator for Couples — Full MVP
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
            completed: false
        },
        chat: {
            messages: [...CHAT_WELCOME_MESSAGES],
            isTyping: false,
            context: [] // tracks conversation topics
        },
        challenges: [...DEMO_CHALLENGES],
        wishes: [...DEMO_WISHES],
        activity: [...DEMO_ACTIVITY],
        complaints: [],
        appreciations: [],
        challengeTab: 'active',
        inviteCode: null,
        analysis: null // filled after onboarding
    },

    // ── Initialization ────────────────────────────────────
    init() {
        const saved = localStorage.getItem('koople_state');
        if (saved) {
            try {
                const parsed = JSON.parse(saved);
                this.state = Object.assign(this.state, parsed);
                if (this.state.user.name && this.state.onboarding.completed) {
                    this.navigate('dashboard');
                    return;
                }
            } catch (e) { /* ignore corrupt data */ }
        }
        this.navigate('welcome');
    },

    save() {
        localStorage.setItem('koople_state', JSON.stringify(this.state));
    },

    resetData() {
        localStorage.removeItem('koople_state');
        location.reload();
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

        // Screen-specific setup
        switch (screenId) {
            case 'dashboard': this.renderDashboard(); break;
            case 'challenges': this.renderChallenges(); break;
            case 'mediator': this.renderChat(); break;
            case 'insights': this.renderInsights(); break;
            case 'wishes': this.renderWishes(); break;
            case 'onboarding': this.renderQuestion(); break;
            case 'profile': this.renderProfile(); break;
        }

        // Update active nav items in target screen
        const navScreens = next.querySelectorAll('.bottom-nav .nav-item');
        navScreens.forEach(item => {
            const onclick = item.getAttribute('onclick') || '';
            item.classList.toggle('active', onclick.includes(`'${screenId}'`));
        });
    },

    // ── Welcome & Auth ────────────────────────────────────
    startOnboarding() { this.navigate('setup'); },
    showLogin() { this.navigate('login'); },

    login() {
        this.state.user.name = 'Аня';
        this.state.user.partnerName = 'Миша';
        this.state.onboarding.completed = true;
        this.state.analysis = this.getDefaultAnalysis();
        this.save();
        this.navigate('dashboard');
    },

    saveProfile() {
        this.state.user.name = document.getElementById('setup-name').value;
        this.state.user.partnerName = document.getElementById('setup-partner').value;
        this.state.user.duration = document.getElementById('setup-duration').value;
        this.state.user.livingTogether = document.getElementById('setup-living').value;
        this.state.user.email = document.getElementById('setup-email').value;
        this.state.onboarding.currentIndex = 0;
        this.save();
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
        const existing = this.state.onboarding.answers[q.id];

        switch (q.type) {
            case 'single':
                q.options.forEach(opt => {
                    const btn = document.createElement('button');
                    btn.className = 'option-btn' + (existing === opt.value ? ' selected' : '');
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
                const selected = existing || [];
                q.options.forEach(opt => {
                    const btn = document.createElement('button');
                    btn.className = 'option-btn multi' + (selected.includes(opt.value) ? ' selected' : '');
                    btn.innerHTML = `<span class="check-icon"></span>${opt.label}`;
                    btn.onclick = () => {
                        let cur = this.state.onboarding.answers[q.id] || [];
                        if (cur.includes(opt.value)) {
                            cur = cur.filter(v => v !== opt.value);
                            btn.classList.remove('selected');
                        } else {
                            if (q.maxSelect && cur.length >= q.maxSelect) return;
                            cur.push(opt.value);
                            btn.classList.add('selected');
                        }
                        this.state.onboarding.answers[q.id] = cur;
                    };
                    area.appendChild(btn);
                });
                break;

            case 'scale':
                const sc = document.createElement('div');
                sc.className = 'scale-container';
                q.scaleLabels.forEach((label, i) => {
                    const val = i + 1;
                    const item = document.createElement('button');
                    item.className = 'scale-item' + (existing === val ? ' selected' : '');
                    item.innerHTML = `<span class="scale-number">${val}</span><span class="scale-label">${label}</span>`;
                    item.onclick = () => {
                        sc.querySelectorAll('.scale-item').forEach(s => s.classList.remove('selected'));
                        item.classList.add('selected');
                        this.state.onboarding.answers[q.id] = val;
                    };
                    sc.appendChild(item);
                });
                area.appendChild(sc);
                break;

            case 'textarea':
                const ta = document.createElement('textarea');
                ta.className = 'answer-textarea';
                ta.placeholder = q.placeholder || '';
                ta.rows = 4;
                ta.value = existing || '';
                ta.oninput = () => { this.state.onboarding.answers[q.id] = ta.value; };
                area.appendChild(ta);
                break;
        }

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

    skipQuestion() { this.nextQuestion(); },

    completeOnboarding() {
        this.state.onboarding.completed = true;
        this.state.inviteCode = this.generateInviteCode();
        this.state.analysis = this.analyzeAnswers();
        this.state.challenges = this.generateInitialChallenges();
        this.save();

        // Set invite code in DOM
        const codeEl = document.getElementById('invite-code');
        if (codeEl) codeEl.textContent = this.state.inviteCode;

        this.navigate('onboarding-complete');
    },

    generateInviteCode() {
        const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
        let code = 'KOOPLE-';
        for (let i = 0; i < 4; i++) code += chars.charAt(Math.floor(Math.random() * chars.length));
        return code;
    },

    // ============================================================
    // ANALYSIS ENGINE — converts answers into insights
    // ============================================================
    analyzeAnswers() {
        const a = this.state.onboarding.answers;

        const catScores = {
            emotional: this.scoreCategory(a, [
                { id: 'emo_express', type: 'scale', weight: 1.2 },
                { id: 'emo_need_closeness', type: 'positive_if', values: ['daily', 'often', 'sometimes'], weight: 1 },
                { id: 'emo_missing', type: 'fewer_is_better', noneValue: 'nothing', weight: 1.5 }
            ]),
            communication: this.scoreCategory(a, [
                { id: 'comm_conflict_style', type: 'positive_if', values: ['discuss', 'humor'], weight: 1.5 },
                { id: 'comm_criticism', type: 'positive_if', values: ['listen', 'depends'], weight: 1.2 },
                { id: 'comm_unspoken', type: 'fewer_is_better', noneValue: 'none', weight: 1.3 },
                { id: 'comm_after_fight', type: 'positive_if', values: ['resolve'], weight: 1.5 }
            ]),
            household: this.scoreCategory(a, [
                { id: 'house_chores', type: 'scale', weight: 1.5 },
                { id: 'house_annoy', type: 'fewer_is_better', noneValue: 'nothing', weight: 1.3 },
                { id: 'house_cleanliness', type: 'scale', weight: 0.8 }
            ]),
            intimacy: this.scoreCategory(a, [
                { id: 'intim_satisfaction', type: 'scale', weight: 1.5 },
                { id: 'intim_frequency', type: 'positive_if', values: ['match'], weight: 1.3 },
                { id: 'intim_affection', type: 'positive_if', values: ['plenty', 'enough'], weight: 1 }
            ]),
            finances: this.scoreCategory(a, [
                { id: 'fin_tension', type: 'positive_if', values: ['never', 'rarely'], weight: 1.5 },
                { id: 'fin_worry', type: 'fewer_is_better', noneValue: 'nothing', weight: 1.3 }
            ]),
            quality_time: this.scoreCategory(a, [
                { id: 'time_together', type: 'scale', weight: 1.5 },
                { id: 'time_activities', type: 'fewer_is_better', noneValue: 'enough', weight: 1 },
                { id: 'time_personal', type: 'positive_if', values: ['plenty', 'enough'], weight: 1 }
            ]),
            family: this.scoreCategory(a, [
                { id: 'fam_inlaws', type: 'scale', weight: 1.3 },
                { id: 'fam_interference', type: 'positive_if', values: ['never'], weight: 1.2 },
                { id: 'fam_friends', type: 'positive_if', values: ['no'], weight: 1 }
            ]),
            values: this.scoreCategory(a, [
                { id: 'val_future', type: 'scale', weight: 1.5 },
                { id: 'val_disagreements', type: 'fewer_is_better', noneValue: 'none', weight: 1.3 }
            ]),
            habits: this.scoreCategory(a, [
                { id: 'hab_triggers', type: 'fewer_is_better', noneValue: null, weight: 1 }
            ])
        };

        // Clamp all scores to 40-98 range for realism
        Object.keys(catScores).forEach(k => {
            catScores[k] = Math.max(40, Math.min(98, catScores[k]));
        });

        const total = Object.values(catScores);
        const harmony = Math.round(total.reduce((s, v) => s + v, 0) / total.length);

        // Identify strengths (top 3) and growth areas (bottom 3)
        const sorted = Object.entries(catScores).sort((a, b) => b[1] - a[1]);
        const categoryLabels = {
            emotional: 'Эмоциональная связь',
            communication: 'Коммуникация',
            household: 'Быт и дом',
            intimacy: 'Близость и интимность',
            finances: 'Финансы',
            quality_time: 'Время вместе',
            family: 'Семья и окружение',
            values: 'Ценности и цели',
            habits: 'Привычки'
        };

        const strengths = sorted.slice(0, 3).map(([k]) => categoryLabels[k]);
        const growth = sorted.slice(-3).reverse().map(([k]) => categoryLabels[k]);

        // Generate personalized recommendations
        const recommendations = this.generateRecommendations(a, catScores);

        return { catScores, harmony, strengths, growth, recommendations };
    },

    scoreCategory(answers, rules) {
        let totalScore = 0;
        let totalWeight = 0;

        for (const rule of rules) {
            const val = answers[rule.id];
            if (val === undefined || val === null) continue;

            let score = 0.5; // default neutral
            switch (rule.type) {
                case 'scale':
                    score = (val - 1) / 4; // 1-5 → 0-1
                    break;
                case 'positive_if':
                    score = rule.values.includes(val) ? 0.85 : 0.35;
                    break;
                case 'fewer_is_better':
                    if (Array.isArray(val)) {
                        if (val.includes(rule.noneValue)) score = 0.95;
                        else score = Math.max(0.15, 1 - val.length * 0.15);
                    } else {
                        score = val === rule.noneValue ? 0.95 : 0.5;
                    }
                    break;
            }
            totalScore += score * rule.weight;
            totalWeight += rule.weight;
        }

        if (totalWeight === 0) return 70; // default if no answers
        return Math.round((totalScore / totalWeight) * 100);
    },

    generateRecommendations(answers, scores) {
        const recs = [];
        const entries = Object.entries(scores).sort((a, b) => a[1] - b[1]);

        // Focus on lowest-scoring areas
        for (const [cat, score] of entries.slice(0, 3)) {
            const priority = score < 55 ? 'high' : score < 70 ? 'medium' : 'low';
            const rec = this.getRecommendationForCategory(cat, answers, priority);
            if (rec) recs.push(rec);
        }
        return recs;
    },

    getRecommendationForCategory(cat, answers, priority) {
        const recs = {
            household: {
                title: 'Наладьте бытовой баланс',
                text: 'Попробуйте составить совместный список обязанностей. Каждый выбирает то, что ему не в тягость — так распределение будет честнее.'
            },
            communication: {
                title: 'Практикуйте безопасные разговоры',
                text: 'Выделите 15 минут в день на разговор без телефонов. Правило: только слушать, не давая советов и не оценивая.'
            },
            intimacy: {
                title: 'Увеличьте физическую близость',
                text: 'Начните с несексуальных прикосновений: объятия, держаться за руки. Это восстанавливает связь постепенно.'
            },
            finances: {
                title: 'Проведите финансовый вечер',
                text: 'Раз в месяц обсуждайте бюджет без упрёков. Начните с одной общей финансовой цели — это сближает.'
            },
            quality_time: {
                title: 'Запланируйте «время вдвоём»',
                text: 'Выделите минимум один вечер в неделю только для вас двоих. Без телефонов, без планов — просто побудьте вместе.'
            },
            family: {
                title: 'Установите границы с семьями',
                text: 'Обсудите, какие темы не стоит обсуждать с родственниками. Единый фронт перед семьями — основа доверия.'
            },
            emotional: {
                title: 'Выучите язык любви партнёра',
                text: 'Попробуйте неделю показывать любовь на языке, который важен партнёру, а не на своём привычном.'
            },
            values: {
                title: 'Обсудите ваше будущее',
                text: 'Найдите тихий вечер и поговорите о том, где вы видите себя через 5 лет. Без давления — просто мечтайте вместе.'
            },
            habits: {
                title: 'Заведите ритуал благодарности',
                text: 'Каждый вечер делитесь одной вещью, за которую вы благодарны партнёру. Это переключает фокус с раздражителей на позитив.'
            }
        };
        const r = recs[cat];
        return r ? { ...r, priority } : null;
    },

    getDefaultAnalysis() {
        return {
            catScores: {
                emotional: 82, communication: 75, household: 62,
                intimacy: 85, finances: 70, quality_time: 78,
                family: 80, values: 88, habits: 58
            },
            harmony: 75,
            strengths: ['Ценности и цели', 'Близость и интимность', 'Эмоциональная связь'],
            growth: ['Привычки', 'Быт и дом', 'Финансы'],
            recommendations: AI_RECOMMENDATIONS
        };
    },

    generateInitialChallenges() {
        const a = this.state.onboarding.answers;
        const analysis = this.state.analysis;
        const partnerName = this.state.user.partnerName || 'партнёр';
        const challenges = [];
        let id = 1;

        // Challenge based on love language
        const loveLang = a.emo_love_language;
        const loveChallenges = {
            words: { title: 'Неделя комплиментов', desc: 'Каждый день говорите партнёру один искренний комплимент', icon: '\uD83D\uDCAC' },
            touch: { title: 'Неделя объятий', desc: 'Обнимайте партнёра минимум 3 раза в день — утром, днём и вечером', icon: '\uD83E\uDEC2' },
            time: { title: 'Вечера вдвоём', desc: 'Каждый вечер 20 минут только для вас — без телефонов и отвлечений', icon: '\u23F0' },
            gifts: { title: 'Неделя сюрпризов', desc: 'Каждый день маленький знак внимания: записка, кофе в постель, цветок', icon: '\uD83C\uDF81' },
            service: { title: 'Неделя заботы', desc: 'Каждый день делайте одну вещь за партнёра: готовка, уборка, поручение', icon: '\uD83D\uDCAA' }
        };
        if (loveLang && loveChallenges[loveLang]) {
            const lc = loveChallenges[loveLang];
            challenges.push({
                id: 'ch' + id++, title: lc.title, description: lc.desc,
                category: 'emotional', icon: lc.icon, duration: '7 дней',
                difficulty: 'easy', progress: 0, total: 7, status: 'active', assignedTo: 'both'
            });
        }

        // Challenge based on household annoyances
        const annoyances = a.house_annoy || [];
        if (annoyances.includes('bathroom') || annoyances.includes('mess')) {
            challenges.push({
                id: 'ch' + id++, title: 'Чистая зона', description: 'Следите за порядком в общих зонах: ванная, кухня. После себя — всегда чисто.',
                category: 'household', icon: '\u2728', duration: '7 дней',
                difficulty: 'easy', progress: 0, total: 7, status: 'active', assignedTo: 'partner'
            });
        }
        if (annoyances.includes('phone')) {
            challenges.push({
                id: 'ch' + id++, title: 'Вечер без телефонов', description: 'Проведите вечер вдвоём без телефонов. Поговорите, поиграйте или просто побудьте вместе.',
                category: 'quality_time', icon: '\uD83D\uDCF5', duration: '1 вечер',
                difficulty: 'medium', progress: 0, total: 1, status: 'active', assignedTo: 'both'
            });
        }
        if (annoyances.includes('dishes')) {
            challenges.push({
                id: 'ch' + id++, title: 'Правило чистой раковины', description: 'Мойте посуду сразу после еды — не оставляйте на потом.',
                category: 'household', icon: '\uD83E\uDDFD', duration: '5 дней',
                difficulty: 'easy', progress: 0, total: 5, status: 'suggested', assignedTo: 'partner'
            });
        }

        // Gratitude challenge is always helpful
        challenges.push({
            id: 'ch' + id++, title: 'Благодарность перед сном',
            description: 'Перед сном расскажите партнёру 3 вещи, за которые вы благодарны ему/ей сегодня.',
            category: 'appreciation', icon: '\uD83C\uDF19', duration: '5 дней',
            difficulty: 'easy', progress: 0, total: 5, status: 'suggested', assignedTo: 'both'
        });

        // Financial challenge if tensions exist
        if (['sometimes', 'often', 'constant'].includes(a.fin_tension)) {
            challenges.push({
                id: 'ch' + id++, title: 'Финансовый вечер',
                description: 'Устройте спокойный разговор о финансах: обсудите траты за месяц, планы и мечты.',
                category: 'finances', icon: '\uD83D\uDCB0', duration: 'Однократно',
                difficulty: 'hard', progress: 0, total: 1, status: 'suggested', assignedTo: 'both'
            });
        }

        // Quality time challenge
        if (a.time_together && a.time_together <= 3) {
            challenges.push({
                id: 'ch' + id++, title: 'Сюрприз-свидание',
                description: 'Организуйте неожиданное свидание для партнёра — от ужина дома до прогулки в новом месте.',
                category: 'quality_time', icon: '\uD83C\uDF39', duration: 'Однократно',
                difficulty: 'medium', progress: 0, total: 1, status: 'suggested', assignedTo: 'user'
            });
        }

        return challenges.length > 0 ? challenges : DEMO_CHALLENGES;
    },

    copyInviteCode() {
        const code = (this.state.inviteCode || document.getElementById('invite-code').textContent);
        navigator.clipboard.writeText(code).then(() => {
            const btns = document.querySelectorAll('.btn-copy');
            btns.forEach(btn => { btn.textContent = 'Скопировано!'; });
            setTimeout(() => btns.forEach(btn => { btn.textContent = 'Копировать'; }), 2000);
        }).catch(() => {});
    },

    shareInvite() {
        const code = this.state.inviteCode || 'KOOPLE-A7X9';
        const text = `Присоединяйся ко мне в Koople — AI-медиаторе для пар! Мой код: ${code}`;
        if (navigator.share) navigator.share({ title: 'Koople', text });
        else this.copyInviteCode();
    },

    goToDashboard() { this.navigate('dashboard'); },

    // ── Dashboard ─────────────────────────────────────────
    renderDashboard() {
        const name = this.state.user.name || 'Аня';
        const analysis = this.state.analysis || this.getDefaultAnalysis();
        const harmony = analysis.harmony;

        document.getElementById('greeting').textContent = `Привет, ${name}!`;
        document.getElementById('user-avatar').textContent = name.charAt(0).toUpperCase();

        // Dynamic health card
        const scoreValue = document.getElementById('health-score-value');
        if (scoreValue) scoreValue.textContent = harmony;

        const scoreCircle = document.querySelector('.score-circle');
        if (scoreCircle) {
            const circumference = 2 * Math.PI * 42;
            const offset = circumference * (1 - harmony / 100);
            scoreCircle.setAttribute('stroke-dasharray', circumference.toFixed(2));
            scoreCircle.setAttribute('stroke-dashoffset', offset.toFixed(2));
        }

        // Dynamic score details
        const detailsEl = document.getElementById('score-details');
        if (detailsEl) {
            const cats = analysis.catScores;
            const topCats = [
                { name: 'Коммуникация', val: cats.communication },
                { name: 'Быт', val: cats.household },
                { name: 'Близость', val: cats.intimacy },
                { name: 'Баланс', val: cats.quality_time }
            ];
            detailsEl.innerHTML = topCats.map(c => {
                const dotClass = c.val >= 80 ? 'dot-green' : c.val >= 65 ? 'dot-yellow' : 'dot-orange';
                return `<div class="score-detail"><span class="dot ${dotClass}"></span> ${c.name}: ${c.val}%</div>`;
            }).join('');
        }

        // Render active challenges
        const challengesEl = document.getElementById('active-challenges');
        const active = this.state.challenges.filter(c => c.status === 'active');
        if (active.length === 0) {
            challengesEl.innerHTML = '<p style="color:var(--text-tertiary);font-size:14px;padding:12px">Нет активных челленджей. Примите предложенные!</p>';
        } else {
            challengesEl.innerHTML = active.map(ch => `
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
        }

        // Render activity feed
        const feedEl = document.getElementById('activity-feed');
        feedEl.innerHTML = this.state.activity.slice(0, 6).map(a => `
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
        if (clickedBtn) {
            document.querySelectorAll('#screen-challenges .tab').forEach(t => t.classList.remove('active'));
            clickedBtn.classList.add('active');
        }

        const list = document.getElementById('challenges-list');
        const filtered = this.state.challenges.filter(c => c.status === tab);
        const partnerName = this.state.user.partnerName || 'партнёра';

        if (filtered.length === 0) {
            const msgs = {
                active: 'Нет активных челленджей. Примите предложенные!',
                completed: 'Пока нет завершённых челленджей',
                suggested: 'AI готовит новые челленджи на основе ваших ответов...'
            };
            list.innerHTML = `<div class="empty-state"><span class="empty-icon">${tab === 'completed' ? '\uD83C\uDFC6' : '\uD83C\uDF31'}</span><p>${msgs[tab]}</p></div>`;
            return;
        }

        list.innerHTML = filtered.map(ch => {
            const pct = (ch.progress / ch.total) * 100;
            const diff = { easy: '\uD83D\uDFE2 Лёгкий', medium: '\uD83D\uDFE1 Средний', hard: '\uD83D\uDFE0 Сложный' }[ch.difficulty];
            const assigned = { both: 'Для обоих', user: 'Для вас', partner: `Для ${partnerName}` }[ch.assignedTo];

            return `
                <div class="challenge-card" onclick="app.showChallengeDetail('${ch.id}')">
                    <div class="challenge-card-header">
                        <span class="challenge-icon-large">${ch.icon}</span>
                        <div>
                            <h3>${ch.title}</h3>
                            <div class="challenge-meta"><span>${diff}</span><span>\u00B7</span><span>${ch.duration}</span><span>\u00B7</span><span>${assigned}</span></div>
                        </div>
                    </div>
                    <p class="challenge-desc">${ch.description}</p>
                    ${ch.status !== 'suggested' ? `
                        <div class="challenge-progress">
                            <div class="progress-bar"><div class="progress-fill ${ch.status === 'completed' ? 'complete' : ''}" style="width:${pct}%"></div></div>
                            <span class="progress-label">${ch.progress} из ${ch.total}</span>
                        </div>
                    ` : `
                        <button class="btn btn-secondary btn-small" onclick="event.stopPropagation(); app.acceptChallenge('${ch.id}')">Принять челлендж</button>
                    `}
                </div>`;
        }).join('');
    },

    showChallengeDetail(id) {
        const ch = this.state.challenges.find(c => c.id === id);
        if (!ch) return;
        const pct = (ch.progress / ch.total) * 100;
        this.showModal(`
            <div class="challenge-detail">
                <div class="challenge-detail-icon">${ch.icon}</div>
                <h2>${ch.title}</h2>
                <p>${ch.description}</p>
                <div class="challenge-detail-meta">
                    <div class="meta-item"><span class="meta-label">Длительность</span><span class="meta-value">${ch.duration}</span></div>
                    <div class="meta-item"><span class="meta-label">Прогресс</span><span class="meta-value">${ch.progress}/${ch.total}</span></div>
                </div>
                <div class="challenge-progress" style="margin-top:16px">
                    <div class="progress-bar"><div class="progress-fill ${ch.status === 'completed' ? 'complete' : ''}" style="width:${pct}%"></div></div>
                </div>
                ${ch.status === 'active' ? `<button class="btn btn-primary btn-large" style="margin-top:20px" onclick="app.markChallengeDay('${ch.id}')">\u2705 Отметить выполнение</button>` : ''}
                ${ch.status === 'suggested' ? `<button class="btn btn-primary btn-large" style="margin-top:20px" onclick="app.acceptChallenge('${ch.id}')">Принять челлендж</button>` : ''}
                ${ch.status === 'completed' ? `<div style="margin-top:20px;color:var(--secondary);font-weight:600">\uD83C\uDFC6 Челлендж завершён!</div>` : ''}
            </div>
        `);
    },

    markChallengeDay(id) {
        const ch = this.state.challenges.find(c => c.id === id);
        if (ch && ch.progress < ch.total) {
            ch.progress++;
            if (ch.progress >= ch.total) ch.status = 'completed';
            this.addActivity(
                ch.progress >= ch.total ? 'challenge_complete' : 'challenge_progress',
                ch.progress >= ch.total ? '\uD83C\uDFC6' : '\u2705',
                ch.progress >= ch.total
                    ? `Челлендж «${ch.title}» завершён!`
                    : `Выполнен день ${ch.progress} в челлендже «${ch.title}»`
            );
            this.save();
        }
        this.closeModal();
        if (this.state.currentScreen === 'challenges') this.renderChallenges();
        else if (this.state.currentScreen === 'dashboard') this.renderDashboard();
    },

    acceptChallenge(id) {
        const ch = this.state.challenges.find(c => c.id === id);
        if (ch) {
            ch.status = 'active';
            this.addActivity('new_challenge', '\uD83C\uDF1F', `Принят челлендж: «${ch.title}»`);
            this.save();
        }
        this.closeModal();
        if (this.state.currentScreen === 'challenges') this.renderChallenges();
    },

    addActivity(type, icon, text) {
        this.state.activity.unshift({ type, icon, text, time: 'Только что' });
        if (this.state.activity.length > 20) this.state.activity.pop();
    },

    // ── AI Mediator Chat ──────────────────────────────────
    renderChat() {
        const messagesEl = document.getElementById('chat-messages');
        const avatarSvg = '<div class="msg-avatar"><svg viewBox="0 0 32 32" fill="none"><circle cx="16" cy="16" r="14" fill="var(--primary-light)"/><circle cx="11" cy="13" r="1.5" fill="var(--primary)"/><circle cx="21" cy="13" r="1.5" fill="var(--primary)"/><path d="M11 20 Q16 24 21 20" stroke="var(--primary)" stroke-width="1.5" fill="none" stroke-linecap="round"/></svg></div>';

        messagesEl.innerHTML = this.state.chat.messages.map(msg => `
            <div class="chat-message ${msg.role}">
                ${msg.role === 'assistant' ? avatarSvg : ''}
                <div class="msg-bubble">
                    <p>${msg.text.replace(/\n/g, '<br>')}</p>
                    <span class="msg-time">${msg.time}</span>
                </div>
            </div>
        `).join('');

        if (this.state.chat.isTyping) {
            messagesEl.innerHTML += `
                <div class="chat-message assistant">
                    ${avatarSvg}
                    <div class="msg-bubble typing"><span class="typing-dot"></span><span class="typing-dot"></span><span class="typing-dot"></span></div>
                </div>`;
        }
        messagesEl.scrollTop = messagesEl.scrollHeight;

        // Show/hide suggestions
        const sugEl = document.getElementById('chat-suggestions');
        if (sugEl) sugEl.style.display = this.state.chat.messages.length <= 1 ? 'flex' : 'none';
    },

    sendMessage() {
        const input = document.getElementById('chat-input');
        const text = input.value.trim();
        if (!text || this.state.chat.isTyping) return;

        this.state.chat.messages.push({ role: 'user', text, time: this.getCurrentTime() });
        input.value = '';
        input.style.height = 'auto';
        this.renderChat();

        this.state.chat.isTyping = true;
        this.renderChat();

        const delay = 1200 + Math.random() * 1500;
        setTimeout(() => {
            this.state.chat.isTyping = false;
            const response = this.generateAIResponse(text);
            this.state.chat.messages.push({ role: 'assistant', text: response.text, time: this.getCurrentTime() });

            // If AI generated a challenge, add it
            if (response.challenge) {
                this.state.challenges.push(response.challenge);
                this.addActivity('new_challenge', '\uD83C\uDF1F', `AI создал челлендж: «${response.challenge.title}»`);
            }

            this.renderChat();
            this.save();
        }, delay);
    },

    sendSuggestion(btn) {
        document.getElementById('chat-input').value = btn.textContent;
        this.sendMessage();
    },

    generateAIResponse(userMessage) {
        const lower = userMessage.toLowerCase();
        const pn = this.state.user.partnerName || 'партнёр';
        const answers = this.state.onboarding.answers;
        const ctx = this.state.chat.context;
        let challenge = null;

        // Track conversation context
        const topics = [];
        if (/ванн|волос|чистот|убира|порядо|грязн/.test(lower)) topics.push('household');
        if (/посуд|кухн|готов/.test(lower)) topics.push('dishes');
        if (/телефон|экран|гаджет/.test(lower)) topics.push('phone');
        if (/денег|деньги|трат|финанс|купи/.test(lower)) topics.push('finances');
        if (/секс|интим|близост|прикоснов|обним/.test(lower)) topics.push('intimacy');
        if (/вним|любо|чувств|эмоц|скуч/.test(lower)) topics.push('emotional');
        if (/родител|мам|пап|свекр|тёщ|семь/.test(lower)) topics.push('family');
        if (/друз|компани|ревну/.test(lower)) topics.push('friends');
        if (/врем|вместе|свидан|гуля|прогул/.test(lower)) topics.push('quality_time');
        if (/работ|карьер|устал|стресс/.test(lower)) topics.push('work');
        if (/обижа|молч|ссор|конфликт|крич/.test(lower)) topics.push('conflict');
        this.state.chat.context = [...ctx, ...topics].slice(-10);

        // ── Pattern matching with rich responses ──

        // Household complaints
        if (/ванн|волос/.test(lower)) {
            challenge = this.createChallengeFromChat('Чистая ванная', 'После каждого использования ванной: убрать волосы, протереть раковину, повесить полотенце.', 'household', '\u2728', 7, 'partner');
            return { text: `Понимаю — волосы в ванной это классика микроконфликтов в парах. Мелочь, но накапливается.\n\nЯ уже создал челлендж «Чистая ванная» для ${pn}. Он/она получит его как рекомендацию для улучшения комфорта в доме — без упоминания жалобы.\n\nЧеллендж рассчитан на 7 дней. Загляните в раздел «Челленджи», чтобы следить за прогрессом.`, challenge };
        }

        if (/посуд|грязн.*тарелк/.test(lower)) {
            challenge = this.createChallengeFromChat('Правило чистой раковины', 'Мойте посуду сразу после еды. Если нет сил — хотя бы замочите. Цель: ни одной грязной тарелки перед сном.', 'household', '\uD83E\uDDFD', 5, 'partner');
            return { text: `Грязная посуда — один из топ-3 бытовых раздражителей в парах. Это не мелочь, когда происходит каждый день.\n\nЯ подготовил челлендж «Правило чистой раковины» — ${pn} получит конкретные простые правила на 5 дней.\n\nВажно: челлендж подан как забота о совместном пространстве, а не как критика. Так он воспринимается намного лучше.`, challenge };
        }

        if (/телефон|экран|гаджет|сидит.*телефон/.test(lower)) {
            challenge = this.createChallengeFromChat('Вечера без экранов', 'Каждый вечер с 20:00 до 21:00 — телефоны в другую комнату. Поговорите, поиграйте или просто побудьте рядом.', 'quality_time', '\uD83D\uDCF5', 5, 'both');
            return { text: `«Ты всё время в телефоне» — одна из самых частых фраз в современных парах. И за ней стоит потребность во внимании.\n\nЯ создал челлендж «Вечера без экранов» — для вас обоих. Это будет час без телефонов каждый вечер. Начать проще, если это правило для двоих, а не претензия к одному.\n\nПосмотрите в «Челленджах» — там все детали.`, challenge };
        }

        // Emotional needs
        if (/не.*заме[чт]а|игнорир|не.*вним|невидим/.test(lower)) {
            challenge = this.createChallengeFromChat('Ежедневный чек-ин', 'Каждый день спрашивайте партнёра: «Как ты сегодня? Что чувствуешь?» — и слушайте ответ без советов.', 'emotional', '\uD83D\uDC96', 7, 'partner');
            return { text: `Чувствовать себя незамеченным — это больно. Ваши чувства абсолютно валидны.\n\nЧасто партнёры не замечают не потому что им всё равно, а потому что по-другому считывают сигналы. ${pn} может искренне не видеть то, что для вас очевидно.\n\nЯ создал челлендж «Ежедневный чек-ин» — ${pn} будет спрашивать каждый день, как вы себя чувствуете. Это простое действие, но оно строит привычку внимательности.\n\nХотите обсудить это подробнее?`, challenge };
        }

        if (/скуча|однообраз|рутин|скучн|нет.*романтик/.test(lower)) {
            challenge = this.createChallengeFromChat('Сюрприз-неделя', 'Каждые 2 дня организуйте маленький сюрприз: необычный ужин, записка, неожиданная прогулка.', 'quality_time', '\uD83C\uDF39', 3, 'both');
            return { text: `Рутина — враг отношений, но и их естественная часть. Хорошая новость: небольшие сюрпризы ломают шаблон лучше, чем грандиозные жесты.\n\nЯ подготовил «Сюрприз-неделю» для вас обоих — 3 маленьких неожиданности за неделю. Это вернёт элемент непредсказуемости.\n\nА пока расскажите: что вам обоим нравилось делать в начале отношений?`, challenge };
        }

        if (/люб|не.*говор.*люблю|не.*слыш.*люблю/.test(lower) && /не|мало|редко|хоч/.test(lower)) {
            return { text: `Потребность слышать «я тебя люблю» — это нормально и важно. Некоторые люди выражают любовь через действия, а не слова, и могут искренне не понимать, что вам нужно это слышать.\n\nМогу помочь двумя способами:\n\n1. Создать мягкий челлендж для ${pn} на словесные выражения чувств\n2. Помочь вам самим сказать ${pn}, что вам это важно\n\nЧто ближе?`, challenge: null };
        }

        // Financial concerns
        if (/денег|деньги|трат|финанс|бюджет/.test(lower)) {
            challenge = this.createChallengeFromChat('Финансовый вечер', 'Устройте спокойный разговор о финансах. Правила: без упрёков, фокус на общих целях, обсуждение одного конкретного вопроса.', 'finances', '\uD83D\uDCB0', 1, 'both');
            return { text: `Финансы — одна из самых сложных тем для пар. Здесь часто сталкиваются разные ценности и привычки из родительских семей.\n\nГлавное правило: обсуждать деньги нужно в спокойной обстановке, а не в момент конфликта.\n\nЯ создал «Финансовый вечер» — структурированный разговор с правилами, которые помогают обсуждать деньги без ссор. Загляните в раздел челленджей.\n\nЧто именно вас беспокоит — траты, накопления или что-то другое?`, challenge };
        }

        // Family issues
        if (/родител|мам|пап|свекр|тёщ|семь.*партнёр/.test(lower)) {
            return { text: `Отношения с семьями партнёров — это всегда тонкая территория. Здесь важен единый фронт: вы — команда, и ваши решения принимаете вы.\n\nНесколько мыслей:\n\n\u2022 Границы с семьями должны устанавливать сами партнёры (каждый со своей)\n\u2022 Никогда не критикуйте семью партнёра — даже если он/она сам жалуется\n\u2022 Обсудите заранее: в каких вопросах мнение родителей допустимо, а в каких — нет\n\nХотите, чтобы я создал челлендж на установление границ?`, challenge: null };
        }

        // Conflict resolution
        if (/ссор|крич|конфликт|ругаемся|скандал/.test(lower)) {
            return { text: `Конфликты в парах неизбежны, но разрушительным может быть не сам конфликт, а то, КАК вы его ведёте.\n\nИсследования Готтмана показывают «четырёх всадников» разрушения отношений:\n\n1. Критика (нападение на личность, а не на поведение)\n2. Защита (контратака вместо слушания)\n3. Презрение (сарказм, закатывание глаз)\n4. Отстранение (молчание, уход)\n\nРасскажите подробнее: как обычно начинается ваш конфликт? Кто первый начинает, и чем обычно заканчивается?`, challenge: null };
        }

        // Intimacy
        if (/секс|интим|близост|не.*хоч.*меня/.test(lower)) {
            return { text: `Спасибо за доверие — это действительно интимная тема, и многим сложно об этом говорить.\n\nВажно помнить: различия в потребностях — это нормально, и они меняются со временем. Стресс, усталость, гормоны — всё влияет.\n\nЯ могу помочь:\n\n1. Создать мягкий челлендж на увеличение несексуальной близости (объятия, прикосновения) — это часто восстанавливает и сексуальное желание\n2. Помочь найти слова, чтобы обсудить это с ${pn}\n\nВсё, что вы скажете, полностью конфиденциально. Что вас волнует больше всего?`, challenge: null };
        }

        // Appreciation / positive
        if (/спасибо|благодар|молодец|хорош|нрави|ценю/.test(lower)) {
            return { text: `Как приятно это слышать! Благодарность — мощнейший инструмент в отношениях.\n\nИсследования показывают, что пары, которые регулярно выражают благодарность, на 50% реже разводятся.\n\nХотите отправить ${pn} тёплое уведомление? Или я могу создать челлендж «Благодарность перед сном» для вас обоих — это потрясающе меняет атмосферу.`, challenge: null };
        }

        // Agreement / confirmation
        if (/^(да|давай|хорошо|ок|конечно|хочу|запус|создай|нужн)/.test(lower) && this.state.chat.context.length > 0) {
            const lastTopic = this.state.chat.context[this.state.chat.context.length - 1];
            return { text: `Отлично! Я подготовлю рекомендации на основе нашего разговора.\n\nВот что произойдёт:\n\u2022 ${pn} получит рекомендацию в мягкой форме\n\u2022 Вы сможете отслеживать прогресс в разделе «Челленджи»\n\u2022 В конце оба оцените результат\n\nМежду тем, есть ли ещё что-то, о чём вы хотели бы поговорить?`, challenge: null };
        }

        // Work/stress
        if (/работ|устал|стресс|выгора|перегруз/.test(lower)) {
            return { text: `Рабочий стресс — один из главных врагов отношений. Когда вы устали, снижается терпение и эмпатия.\n\nНесколько идей:\n\n\u2022 «Правило 20 минут» — после работы дайте себе 20 минут тишины перед общением\n\u2022 Договоритесь о «сигналах усталости» — чтобы партнёр понимал, когда не время для важных разговоров\n\u2022 Один вечер в неделю — только отдых, никаких рабочих тем\n\nЧто из этого откликается? Могу создать челлендж.`, challenge: null };
        }

        // Discussion / general talk
        if (/отношени|обсудить|поговорить|разобрать/.test(lower)) {
            return { text: `Конечно, давайте поговорим. Я здесь, чтобы помочь вам разобраться в чувствах.\n\nМожете начать с того, что вас волнует больше всего прямо сейчас? Не обязательно формулировать идеально — просто расскажите как есть.\n\nЯ замечу паттерны и помогу увидеть ситуацию с разных сторон.`, challenge: null };
        }

        // Hard to say
        if (/сложно.*сказать|не могу.*сказать|стесня|бою.*сказать/.test(lower)) {
            return { text: `Это нормально — многие вещи сложно произнести вслух. Именно для этого я здесь.\n\nВы можете написать мне всё, что хотели бы сказать ${pn}, но не решаетесь. Я помогу:\n\n1. Разобраться в ваших чувствах\n2. Найти правильные слова\n3. Или передать пожелание через мягкий челлендж\n\nЧто бы вы хотели сказать?`, challenge: null };
        }

        // Generic complaint
        if (/беспокоит|раздражает|бесит|напрягает|достал|надоел/.test(lower)) {
            return { text: `Я слышу, что это вас беспокоит. Спасибо, что доверяете мне.\n\nЧтобы помочь максимально эффективно, расскажите подробнее:\n\n\u2022 Как давно это происходит?\n\u2022 Как часто?\n\u2022 Говорили ли вы об этом с ${pn}?\n\nЯ не буду передавать ваши слова напрямую. Вместо этого я создам мягкий челлендж, который поможет улучшить ситуацию естественным образом.`, challenge: null };
        }

        // Default response with personalization
        const analysis = this.state.analysis;
        let extra = '';
        if (analysis && analysis.growth.length > 0) {
            extra = `\n\nКстати, по вашей анкете я вижу зоны роста: ${analysis.growth.join(', ')}. Хотите обсудить что-то из этого?`;
        }

        return { text: `Спасибо, что делитесь этим. Я анализирую ваши слова вместе с данными анкеты.\n\nМогу предложить:\n\n1. \uD83D\uDCAC Обсудить ситуацию подробнее\n2. \uD83C\uDFAF Создать мягкий челлендж для ${pn}\n3. \uD83D\uDCA1 Дать рекомендацию для вас обоих\n4. \u2764\uFE0F Отправить ${pn} что-то приятное\n\nЧто вам ближе?${extra}`, challenge: null };
    },

    createChallengeFromChat(title, desc, category, icon, total, assignedTo) {
        return {
            id: 'ch_' + Date.now(),
            title, description: desc, category, icon,
            duration: total === 1 ? 'Однократно' : `${total} дней`,
            difficulty: total <= 3 ? 'easy' : 'medium',
            progress: 0, total, status: 'active', assignedTo
        };
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
        const analysis = this.state.analysis || this.getDefaultAnalysis();
        const cats = analysis.catScores;

        // Dynamic strengths / growth
        const strengthsEl = document.getElementById('strengths-list');
        const growthEl = document.getElementById('growth-list');
        if (strengthsEl) strengthsEl.innerHTML = analysis.strengths.map(s => `<li class="positive">${s}</li>`).join('');
        if (growthEl) growthEl.innerHTML = analysis.growth.map(g => `<li class="growth">${g}</li>`).join('');

        // Compatibility bars
        const labels = {
            emotional: 'Эмоциональная связь', communication: 'Коммуникация',
            household: 'Быт', intimacy: 'Близость', finances: 'Финансы',
            quality_time: 'Время вместе', family: 'Семья', values: 'Ценности', habits: 'Привычки'
        };
        const barsEl = document.getElementById('compatibility-bars');
        barsEl.innerHTML = Object.entries(cats).map(([key, val]) => {
            const color = val >= 80 ? '#4ade80' : val >= 65 ? '#facc15' : '#fb923c';
            return `<div class="compat-row">
                <span class="compat-name">${labels[key] || key}</span>
                <div class="compat-bar-track"><div class="compat-bar-fill" style="width:${val}%;background:${color}"></div></div>
                <span class="compat-value">${val}%</span>
            </div>`;
        }).join('');

        // Progress chart
        const chartEl = document.getElementById('progress-chart');
        const base = Math.max(50, analysis.harmony - 15);
        const weeks = ['Нед 1', 'Нед 2', 'Нед 3', 'Сейчас'];
        const vals = [base, base + 5, base + 9, analysis.harmony];
        chartEl.innerHTML = `<div class="bar-chart">${weeks.map((w, i) => `
            <div class="bar-col"><div class="bar" style="height:${vals[i]}%"><span class="bar-value">${vals[i]}</span></div><span class="bar-label">${w}</span></div>
        `).join('')}</div>`;

        // Recommendations
        const recs = analysis.recommendations || AI_RECOMMENDATIONS;
        const recsEl = document.getElementById('ai-recommendations');
        recsEl.innerHTML = recs.map(rec => `
            <div class="recommendation-card priority-${rec.priority}"><h4>${rec.title}</h4><p>${rec.text}</p></div>
        `).join('');
    },

    // ── Wishes ────────────────────────────────────────────
    renderWishes() {
        const listEl = document.getElementById('wishes-list');
        if (this.state.wishes.length === 0) {
            listEl.innerHTML = '<div class="empty-state"><span class="empty-icon">\u2B50</span><p>У вас пока нет желаний. Нажмите + чтобы добавить первое.</p></div>';
            return;
        }
        listEl.innerHTML = this.state.wishes.map(w => {
            const label = { active: 'Активно', in_progress: 'AI работает над этим', fulfilled: 'Исполнено' }[w.status] || 'Активно';
            return `<div class="wish-card">
                <p class="wish-text">${w.text}</p>
                <div class="wish-meta">
                    <span class="wish-status ${w.status}">${label}</span>
                    <span class="wish-date">${w.createdAt}</span>
                </div>
            </div>`;
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
        this.state.wishes.unshift({ id: 'w' + Date.now(), text, category, status: 'active', createdAt: 'Только что' });
        this.addActivity('wish', '\u2B50', 'Добавлено новое желание');
        this.save();
        this.closeModal();
        this.renderWishes();
        this.showNotification('\u2B50 Желание добавлено! AI начнёт работать над ним');
    },

    // ── Profile ───────────────────────────────────────────
    renderProfile() {
        const name = this.state.user.name || 'Аня';
        const partnerName = this.state.user.partnerName || 'Миша';
        const analysis = this.state.analysis || this.getDefaultAnalysis();

        document.getElementById('profile-name').textContent = name;
        document.getElementById('profile-partner-name').textContent = partnerName;
        document.querySelector('.profile-avatar-large').textContent = name.charAt(0).toUpperCase();

        // Dynamic stats
        const completed = this.state.challenges.filter(c => c.status === 'completed').length;
        const statsEl = document.getElementById('profile-stats');
        if (statsEl) {
            statsEl.innerHTML = `
                <div class="stat"><span class="stat-value">${this.state.challenges.length}</span><span class="stat-label">челленджей</span></div>
                <div class="stat"><span class="stat-value">${completed}</span><span class="stat-label">завершено</span></div>
                <div class="stat"><span class="stat-value">${analysis.harmony}%</span><span class="stat-label">гармония</span></div>
            `;
        }
    },

    showInvitePartner() {
        const code = this.state.inviteCode || 'KOOPLE-A7X9';
        this.showModal(`
            <h2>Пригласить партнёра</h2>
            <p>Отправьте этот код партнёру, чтобы он/она смог присоединиться</p>
            <div class="invite-code-box">
                <span class="invite-code">${code}</span>
                <button class="btn-copy" onclick="app.copyInviteCode()">Копировать</button>
            </div>
            <button class="btn btn-secondary btn-large" onclick="app.shareInvite()">Отправить приглашение</button>
        `);
    },

    showNotifications() {
        this.showModal(`
            <h2>Уведомления</h2>
            <p class="modal-subtitle">Настройте, когда получать напоминания</p>
            <div class="settings-toggles">
                <div class="setting-row"><span>Напоминания о челленджах</span><span class="toggle-pill active" onclick="this.classList.toggle('active')"></span></div>
                <div class="setting-row"><span>Новые рекомендации AI</span><span class="toggle-pill active" onclick="this.classList.toggle('active')"></span></div>
                <div class="setting-row"><span>Активность партнёра</span><span class="toggle-pill active" onclick="this.classList.toggle('active')"></span></div>
                <div class="setting-row"><span>Еженедельный отчёт</span><span class="toggle-pill" onclick="this.classList.toggle('active')"></span></div>
            </div>
        `);
    },

    showPrivacy() {
        this.showModal(`
            <h2>Приватность</h2>
            <p class="modal-subtitle">Ваши данные под защитой</p>
            <div class="privacy-info">
                <div class="privacy-item"><strong>Конфиденциальность чата</strong><p>Ваши сообщения медиатору никогда не передаются партнёру напрямую</p></div>
                <div class="privacy-item"><strong>Анкета</strong><p>Ответы анализируются AI, но конкретные ответы не показываются партнёру</p></div>
                <div class="privacy-item"><strong>Жалобы</strong><p>Ваши жалобы трансформируются в мягкие челленджи без указания источника</p></div>
            </div>
            <button class="btn btn-secondary btn-large" onclick="app.resetData()" style="margin-top:20px;color:#f87171">Удалить все данные</button>
        `);
    },

    showHelp() {
        this.showModal(`
            <h2>Как работает Koople?</h2>
            <div class="help-steps">
                <div class="help-step"><span class="help-num">1</span><div><strong>Пройдите анкету</strong><p>40 вопросов помогут AI понять вашу пару</p></div></div>
                <div class="help-step"><span class="help-num">2</span><div><strong>Пригласите партнёра</strong><p>Он/она тоже проходит анкету отдельно</p></div></div>
                <div class="help-step"><span class="help-num">3</span><div><strong>Получайте челленджи</strong><p>AI создаёт персонализированные мини-задания</p></div></div>
                <div class="help-step"><span class="help-num">4</span><div><strong>Общайтесь с медиатором</strong><p>Жалуйтесь, просите совет, хвалите партнёра</p></div></div>
                <div class="help-step"><span class="help-num">5</span><div><strong>Растите вместе</strong><p>Отслеживайте прогресс и укрепляйте отношения</p></div></div>
            </div>
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

    // ── Complaint & Appreciation ──────────────────────────
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
            <input type="hidden" id="complaint-urgency" value="medium">
            <button class="btn btn-primary btn-large" onclick="app.submitComplaint()">Отправить медиатору</button>
        `);
    },

    selectUrgency(btn, level) {
        document.querySelectorAll('.urgency-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        const hidden = document.getElementById('complaint-urgency');
        if (hidden) hidden.value = level;
    },

    submitComplaint() {
        const text = document.getElementById('complaint-input').value.trim();
        if (!text) return;
        const urgencyEl = document.getElementById('complaint-urgency');
        const urgency = urgencyEl ? urgencyEl.value : 'medium';

        this.state.complaints.push({ text, urgency, time: new Date().toISOString() });
        this.addActivity('complaint', '\uD83D\uDE14', 'Вы поделились беспокойством с AI-медиатором');
        this.closeModal();

        // Navigate to chat with the complaint
        const prefix = { low: '', medium: 'Меня беспокоит: ', high: 'Мне серьёзно мешает: ' }[urgency] || 'Меня беспокоит: ';
        this.state.chat.messages.push({ role: 'user', text: `${prefix}${text}`, time: this.getCurrentTime() });
        this.navigate('mediator');

        this.state.chat.isTyping = true;
        this.renderChat();
        setTimeout(() => {
            this.state.chat.isTyping = false;
            const response = this.generateAIResponse(text);
            this.state.chat.messages.push({ role: 'assistant', text: response.text, time: this.getCurrentTime() });
            if (response.challenge) {
                this.state.challenges.push(response.challenge);
                this.addActivity('new_challenge', '\uD83C\uDF1F', `AI создал челлендж: «${response.challenge.title}»`);
            }
            this.renderChat();
            this.save();
        }, 2000);
    },

    showAppreciationModal() {
        const pn = this.state.user.partnerName || 'партнёра';
        this.showModal(`
            <h2>\uD83D\uDC95 Похвалить ${pn}</h2>
            <p class="modal-subtitle">Позитивная обратная связь укрепляет отношения</p>
            <textarea id="appreciation-input" class="answer-textarea" placeholder="За что вы хотите поблагодарить или похвалить партнёра?" rows="4"></textarea>
            <button class="btn btn-primary btn-large" onclick="app.submitAppreciation()">Отправить</button>
        `);
    },

    submitAppreciation() {
        const text = document.getElementById('appreciation-input').value.trim();
        if (!text) return;
        this.state.appreciations.push({ text, time: new Date().toISOString() });
        this.addActivity('appreciation', '\uD83D\uDC95', `Вы похвалили ${this.state.user.partnerName || 'партнёра'}`);
        this.save();
        this.closeModal();
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
        document.getElementById('modal-overlay').classList.remove('active');
        document.body.style.overflow = '';
    },

    // ── Notifications ─────────────────────────────────────
    showNotification(text) {
        const notif = document.createElement('div');
        notif.className = 'notification';
        notif.textContent = text;
        document.body.appendChild(notif);
        requestAnimationFrame(() => notif.classList.add('show'));
        setTimeout(() => { notif.classList.remove('show'); setTimeout(() => notif.remove(), 300); }, 3000);
    }
};

// ── Initialize & Keyboard ──────────────────────────────
document.addEventListener('DOMContentLoaded', () => app.init());
document.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && !e.shiftKey && document.activeElement.id === 'chat-input') {
        e.preventDefault();
        app.sendMessage();
    }
});
