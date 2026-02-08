// ============================================================
// KOOPLE — Main Application Logic v3.0
// AI-Mediator for Couples — Full MVP with i18n + AI integration
// ============================================================

const app = {
    // ── State ──────────────────────────────────────────────
    state: {
        currentScreen: 'intro',
        previousScreen: null
        introSlide: 0,
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
            messages: [],
            isTyping: false,
            context: []
        },
        challenges: [...DEMO_CHALLENGES],
        wishes: [...DEMO_WISHES],
        activity: [...DEMO_ACTIVITY],
        complaints: [],
        appreciations: [],
        challengeTab: 'active',
        inviteCode: null,
        analysis: null
    },

    // ── Initialization ────────────────────────────────────
    init() {
        // Restore saved language before anything else
        try {
            const savedLang = localStorage.getItem('koople_lang');
            if (savedLang && typeof setLang === 'function') setLang(savedLang);
        } catch (e) { /* storage unavailable */ }

        const saved = localStorage.getItem('koople_state');
        if (saved) {
            try {
                const parsed = JSON.parse(saved);
                this.state = Object.assign(this.state, parsed);
                // Migrate: clear old-format data (hardcoded Russian text, fake activity)
                if (this.state.activity && this.state.activity.length > 0 &&
                    this.state.activity.some(a => a.text && /[а-яА-ЯёЁ]/.test(a.text) && /Миша|выполнил|день \d/.test(a.text))) {
                    this.state.activity = [];
                }
                // Migrate old challenges that have title but no titleKey (old Russian-only format)
                if (this.state.challenges && this.state.challenges.length > 0 &&
                    this.state.challenges[0].title && !this.state.challenges[0].titleKey) {
                    this.state.challenges = [...DEMO_CHALLENGES];
                }
                if (this.state.user.name && this.state.onboarding.completed) {
                    this.navigate('dashboard');
                    return;
                }
            } catch (e) { /* ignore corrupt data */ }
        }
        this.initIntroSlides();
        this.applyTranslations();
        // navigate('intro') would return early since intro is already active from HTML,
        // so we just set the state directly
        this.state.currentScreen = 'intro';
    },

    save() {
        localStorage.setItem('koople_state', JSON.stringify(this.state));
    },

    resetData() {
        localStorage.removeItem('koople_state');
        localStorage.removeItem('koople_lang');
        localStorage.removeItem('koople_ai_config');
        location.reload();
    },

    // ── i18n helpers ─────────────────────────────────────
    // Resolve: if object has a key-based field, translate it; else use raw text
    ct(obj, field) {
        if (obj[field + 'Key']) return this.t(obj[field + 'Key']);
        return obj[field] || '';
    },

    t(key) {
        return (typeof t === 'function') ? t(key) : key;
    },

    applyTranslations() {
        // Nav labels
        document.querySelectorAll('[data-nav]').forEach(el => {
            const key = 'nav_' + el.dataset.nav;
            el.textContent = this.t(key);
        });
        // Static text elements by ID mapping
        const map = {
            'welcome-tagline': 'app_tagline',
            'welcome-feature1': 'intro_1_title',
            'welcome-feature2': 'intro_2_title',
            'welcome-feature3': 'intro_3_title',
            'btn-start': 'btn_start',
            'btn-login': 'btn_have_account',
            'login-title': 'login_title',
            'login-email-label': 'login_email_label',
            'login-password-label': 'login_password_label',
            'login-submit-btn': 'btn_login',
            'setup-title': 'setup_title',
            'setup-subtitle': 'setup_subtitle',
            'setup-name-label': 'setup_name_label',
            'setup-partner-label': 'setup_partner_label',
            'setup-duration-label': 'setup_duration_label',
            'setup-living-label': 'setup_living_label',
            'setup-email-label': 'setup_email_label',
            'setup-next-btn': 'btn_next',
            'setup-yes-btn': 'yes',
            'setup-no-btn': 'no',
            'setup-dur-default': 'select_placeholder',
            'setup-dur-6m': 'duration_opt_1',
            'setup-dur-6m1y': 'duration_opt_2',
            'setup-dur-1-3y': 'duration_opt_3',
            'setup-dur-3-5y': 'duration_opt_4',
            'setup-dur-5-10y': 'duration_opt_5',
            'setup-dur-10y': 'duration_opt_6',
            'btn-skip-question': 'btn_skip',
            'btn-next-question': 'btn_next',
            'complete-title': 'onboarding_complete_title',
            'complete-subtitle': 'onboarding_complete_text',
            'complete-invite-title': 'invite_title',
            'complete-invite-desc': 'invite_text',
            'btn-copy': 'btn_copy',
            'btn-share': 'btn_share_invite',
            'btn-go-dashboard': 'btn_go_dashboard',
            'health-label': 'harmony_label',
            'dash-challenges-title': 'active_challenges_title',
            'dash-all-btn': 'btn_all',
            'action-mediator': 'action_talk',
            'action-complaint': 'action_concern',
            'action-praise': 'action_appreciate',
            'action-wishes': 'action_wishes',
            'dash-recent-title': 'section_recent',
            'challenges-page-title': 'challenges_title',
            'challenges-page-subtitle': 'challenges_subtitle',
            'tab-active': 'tab_active',
            'tab-completed': 'tab_completed',
            'tab-suggested': 'tab_suggested',
            'chat-status': 'mediator_status',
            'insights-page-title': 'insights_title',
            'insights-page-subtitle': 'insights_subtitle',
            'strengths-title': 'strengths_title',
            'growth-title': 'growth_title',
            'compat-title': 'compat_title',
            'progress-title': 'progress_month_title',
            'recs-title': 'ai_rec_title',
            'wishes-page-title': 'wishes_title',
            'wishes-page-subtitle': 'wishes_subtitle',
            'profile-couple-label': 'profile_couple_with',
            'settings-retake': 'settings_retake',
            'settings-invite': 'settings_invite',
            'settings-language': 'settings_language',
            'settings-ai': 'settings_ai',
            'settings-notif': 'settings_notifications',
            'settings-privacy': 'settings_privacy',
            'settings-help': 'settings_help',
            'settings-about': 'settings_about',
            'btn-intro-skip': 'btn_skip',
            'btn-intro-next': 'btn_next',
            'btn-milestone-continue': 'milestone_btn',
            'intro-lang-label': 'setup_lang_label'
        };
        for (const [id, key] of Object.entries(map)) {
            const el = document.getElementById(id);
            if (el) el.textContent = this.t(key);
        }
        // Intro slides
        for (let i = 0; i < 3; i++) {
            const tEl = document.getElementById('intro-title-' + i);
            const dEl = document.getElementById('intro-desc-' + i);
            if (tEl) tEl.textContent = this.t('intro_' + (i+1) + '_title');
            if (dEl) dEl.textContent = this.t('intro_' + (i+1) + '_text');
        }
        // Chat suggestions
        const s1 = document.getElementById('suggestion-1');
        const s2 = document.getElementById('suggestion-2');
        const s3 = document.getElementById('suggestion-3');
        if (s1) s1.textContent = this.t('suggestion_1');
        if (s2) s2.textContent = this.t('suggestion_2');
        if (s3) s3.textContent = this.t('suggestion_3');
        // Chat input placeholder
        const ci = document.getElementById('chat-input');
        if (ci) ci.placeholder = this.t('chat_placeholder');
        // Setup input placeholders
        const sn = document.getElementById('setup-name');
        if (sn) sn.placeholder = this.t('setup_name_placeholder');
        const sp = document.getElementById('setup-partner');
        if (sp) sp.placeholder = this.t('setup_partner_placeholder');
    },

    // ── Intro Slides ─────────────────────────────────────
    initIntroSlides() {
        const grid = document.getElementById('lang-grid');
        if (grid && typeof LANGUAGES !== 'undefined') {
            const currentLangCode = (typeof getLang === 'function') ? getLang() : 'en';
            grid.innerHTML = LANGUAGES.map(l => `
                <button class="lang-option ${l.code === currentLangCode ? 'active' : ''}" onclick="app.selectLanguage('${l.code}')">
                    <span class="lang-flag">${l.flag}</span>
                    <span class="lang-name">${l.name}</span>
                </button>
            `).join('');
        }
        this.updateIntroSlide();
    },

    selectLanguage(code) {
        if (typeof setLang === 'function') setLang(code);
        try { localStorage.setItem('koople_lang', code); } catch(e) {}
        // Update lang grid active state
        document.querySelectorAll('.lang-option').forEach(btn => {
            btn.classList.toggle('active', btn.querySelector('.lang-name').textContent ===
                (LANGUAGES.find(l => l.code === code) || {}).name);
        });
        this.applyTranslations();
        this.updateIntroSlide();
    },

    updateIntroSlide() {
        const slides = document.querySelectorAll('.intro-slide');
        const dots = document.querySelectorAll('.intro-dot');
        const idx = this.state.introSlide;
        slides.forEach((s, i) => s.classList.toggle('active', i === idx));
        dots.forEach((d, i) => d.classList.toggle('active', i === idx));
        // Show language selector only on first slide
        const langSel = document.getElementById('lang-selector');
        if (langSel) langSel.style.display = idx === 0 ? 'block' : 'none';
        // Update button text
        const btn = document.getElementById('btn-intro-next');
        if (btn) btn.textContent = idx === 2 ? this.t('btn_start') : this.t('btn_next');
    },

    nextIntroSlide() {
        if (this.state.introSlide < 2) {
            this.state.introSlide++;
            this.updateIntroSlide();
        } else {
            this.navigate('welcome');
        }
    },

    skipIntro() {
        this.navigate('welcome');
    },

    // ── Navigation ────────────────────────────────────────
    navigate(screenId) {
        const current = document.querySelector('.screen.active');
        const next = document.getElementById('screen-' + screenId);
        if (!next || next === current) return;

        this.state.previousScreen = this.state.currentScreen;
        this.state.currentScreen = screenId;

        if (current) {
            current.classList.add('exit');
            current.classList.remove('active');
            setTimeout(() => current.classList.remove('exit'), 400);
        }
        next.classList.add('active');

        switch (screenId) {
            case 'dashboard': this.renderDashboard(); break;
            case 'challenges': this.renderChallenges(); break;
            case 'mediator': this.renderChat(); break;
            case 'insights': this.renderInsights(); break;
            case 'wishes': this.renderWishes(); break;
            case 'onboarding': this.renderQuestion(); break;
            case 'profile': this.renderProfile(); break;
            case 'welcome': this.applyTranslations(); break;
            case 'setup': this.applyTranslations(); break;
            case 'intro': this.initIntroSlides(); this.applyTranslations(); break;
        }

        const navScreens = next.querySelectorAll('.bottom-nav .nav-item');
        navScreens.forEach(item => {
            const onclick = item.getAttribute('onclick') || '';
            item.classList.toggle('active', onclick.includes("'" + screenId + "'"));
        });
    },

    // ── Welcome & Auth ────────────────────────────────────
    startOnboarding() { this.navigate('setup'); },
    showLogin() { this.navigate('login'); },

    login() {
        this.state.user.name = 'User';
        this.state.user.partnerName = 'Partner';
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

    retakeQuestionnaire() {
        this.state.onboarding.currentIndex = 0;
        this.state.onboarding.answers = {};
        this.state.onboarding.completed = false;
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
        document.getElementById('onboarding-progress-text').textContent = (idx + 1) + ' / ' + total;
        document.getElementById('category-icon').textContent = q.categoryIcon;

        // i18n for category name
        const catKey = 'cat_' + q.category;
        document.getElementById('category-name').textContent = this.t(catKey) !== catKey ? this.t(catKey) : q.categoryName;

        // i18n for question text / hint
        const qText = (typeof qt === 'function' && qt(q.id, 'text')) || q.text;
        const qHint = (typeof qt === 'function' && qt(q.id, 'hint')) || q.hint || '';
        document.getElementById('question-text').textContent = qText;
        document.getElementById('question-hint').textContent = qHint;

        const area = document.getElementById('answer-area');
        area.innerHTML = '';
        const existing = this.state.onboarding.answers[q.id];

        // Get translated options/scaleLabels if available
        const qOpts = (typeof qt === 'function' && qt(q.id, 'options')) || null;
        const qScale = (typeof qt === 'function' && qt(q.id, 'scaleLabels')) || null;

        switch (q.type) {
            case 'single':
                q.options.forEach((opt, oi) => {
                    const btn = document.createElement('button');
                    btn.className = 'option-btn' + (existing === opt.value ? ' selected' : '');
                    btn.textContent = (qOpts && qOpts[oi]) || opt.label;
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
                q.options.forEach((opt, oi) => {
                    const btn = document.createElement('button');
                    btn.className = 'option-btn multi' + (selected.includes(opt.value) ? ' selected' : '');
                    btn.innerHTML = '<span class="check-icon"></span>' + ((qOpts && qOpts[oi]) || opt.label);
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
                const labels = qScale || q.scaleLabels;
                labels.forEach((label, i) => {
                    const val = i + 1;
                    const item = document.createElement('button');
                    item.className = 'scale-item' + (existing === val ? ' selected' : '');
                    item.innerHTML = '<span class="scale-number">' + val + '</span><span class="scale-label">' + label + '</span>';
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
                const qPh = (typeof qt === 'function' && qt(q.id, 'placeholder')) || q.placeholder || '';
                ta.placeholder = qPh;
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
        const idx = this.state.onboarding.currentIndex;
        if (idx < QUESTIONNAIRE.length - 1) {
            this.state.onboarding.currentIndex++;
            // Check for milestones at Q10, Q20, Q30
            const nextIdx = this.state.onboarding.currentIndex;
            if (nextIdx === 10 || nextIdx === 20 || nextIdx === 30) {
                this.showMilestone(nextIdx);
            } else {
                this.renderQuestion();
            }
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

    // ── Milestones (motivational screens during quiz) ─────
    showMilestone(questionIndex) {
        const overlay = document.getElementById('milestone-overlay');
        if (!overlay) { this.renderQuestion(); return; }

        const emojis = { 10: '\uD83C\uDF89', 20: '\uD83D\uDE80', 30: '\uD83C\uDFC6' };
        const pct = Math.round((questionIndex / QUESTIONNAIRE.length) * 100);

        document.getElementById('milestone-emoji').textContent = emojis[questionIndex] || '\uD83C\uDF1F';
        document.getElementById('milestone-pct').textContent = pct + '%';
        document.getElementById('milestone-title').textContent = this.t('milestone_' + questionIndex + '_title');
        document.getElementById('milestone-desc').textContent = this.t('milestone_' + questionIndex + '_text');
        document.getElementById('btn-milestone-continue').textContent = this.t('milestone_btn');

        // Animate ring
        const ring = document.getElementById('milestone-ring-fill');
        if (ring) {
            const circumference = 2 * Math.PI * 42;
            const offset = circumference * (1 - pct / 100);
            ring.setAttribute('stroke-dasharray', circumference.toFixed(2));
            ring.setAttribute('stroke-dashoffset', offset.toFixed(2));
        }

        overlay.classList.add('active');
    },

    closeMilestone() {
        const overlay = document.getElementById('milestone-overlay');
        if (overlay) overlay.classList.remove('active');
        this.renderQuestion();
    },

    completeOnboarding() {
        this.state.onboarding.completed = true;
        this.state.inviteCode = this.generateInviteCode();
        this.state.analysis = this.analyzeAnswers();
        this.state.challenges = this.generateInitialChallenges();
        this.save();

        const codeEl = document.getElementById('invite-code');
        if (codeEl) codeEl.textContent = this.state.inviteCode;

        this.navigate('onboarding-complete');
        this.applyTranslations();
    },

    generateInviteCode() {
        const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
        let code = 'KOOPLE-';
        for (let i = 0; i < 4; i++) code += chars.charAt(Math.floor(Math.random() * chars.length));
        return code;
    },

    // ============================================================
    // ANALYSIS ENGINE
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

        Object.keys(catScores).forEach(k => {
            catScores[k] = Math.max(40, Math.min(98, catScores[k]));
        });

        const total = Object.values(catScores);
        const harmony = Math.round(total.reduce((s, v) => s + v, 0) / total.length);

        const sorted = Object.entries(catScores).sort((a, b) => b[1] - a[1]);
        const strengths = sorted.slice(0, 3).map(([k]) => this.t('cat_' + k));
        const growth = sorted.slice(-3).reverse().map(([k]) => this.t('cat_' + k));
        const recommendations = this.generateRecommendations(a, catScores);

        return { catScores, harmony, strengths, growth, recommendations };
    },

    scoreCategory(answers, rules) {
        let totalScore = 0, totalWeight = 0;
        for (const rule of rules) {
            const val = answers[rule.id];
            if (val === undefined || val === null) continue;
            let score = 0.5;
            switch (rule.type) {
                case 'scale': score = (val - 1) / 4; break;
                case 'positive_if': score = rule.values.includes(val) ? 0.85 : 0.35; break;
                case 'fewer_is_better':
                    if (Array.isArray(val)) {
                        if (val.includes(rule.noneValue)) score = 0.95;
                        else score = Math.max(0.15, 1 - val.length * 0.15);
                    } else { score = val === rule.noneValue ? 0.95 : 0.5; }
                    break;
            }
            totalScore += score * rule.weight;
            totalWeight += rule.weight;
        }
        if (totalWeight === 0) return 70;
        return Math.round((totalScore / totalWeight) * 100);
    },

    generateRecommendations(answers, scores) {
        const recs = [];
        const entries = Object.entries(scores).sort((a, b) => a[1] - b[1]);
        for (const [cat, score] of entries.slice(0, 3)) {
            const priority = score < 55 ? 'high' : score < 70 ? 'medium' : 'low';
            const rec = this.getRecommendationForCategory(cat, answers, priority);
            if (rec) recs.push(rec);
        }
        return recs;
    },

    getRecommendationForCategory(cat, answers, priority) {
        const recs = {
            household: { title: this.t('cat_household'), text: this.t('empty_suggested') },
            communication: { title: this.t('cat_communication'), text: '' },
            intimacy: { title: this.t('cat_intimacy'), text: '' },
            finances: { title: this.t('cat_finances'), text: '' },
            quality_time: { title: this.t('cat_quality_time'), text: '' },
            family: { title: this.t('cat_family'), text: '' },
            emotional: { title: this.t('cat_emotional'), text: '' },
            values: { title: this.t('cat_values'), text: '' },
            habits: { title: this.t('cat_habits'), text: '' }
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
            strengths: [this.t('cat_values'), this.t('cat_intimacy'), this.t('cat_emotional')],
            growth: [this.t('cat_habits'), this.t('cat_household'), this.t('cat_finances')],
            recommendations: AI_RECOMMENDATIONS
        };
    },

    generateInitialChallenges() {
        const a = this.state.onboarding.answers;
        const challenges = [];
        let id = 1;

        const loveLang = a.emo_love_language;
        const loveChallenges = {
            words: { titleKey: 'ch_compliment_week', descKey: 'ch_compliment_week_desc', icon: '\uD83D\uDCAC' },
            touch: { titleKey: 'ch_hug_week', descKey: 'ch_hug_week_desc', icon: '\uD83E\uDEC2' },
            time: { titleKey: 'ch_evenings_together', descKey: 'ch_evenings_together_desc', icon: '\u23F0' },
            gifts: { titleKey: 'ch_surprise_week', descKey: 'ch_surprise_week_desc', icon: '\uD83C\uDF81' },
            service: { titleKey: 'ch_care_week', descKey: 'ch_care_week_desc', icon: '\uD83D\uDCAA' }
        };
        if (loveLang && loveChallenges[loveLang]) {
            const lc = loveChallenges[loveLang];
            challenges.push({
                id: 'ch' + id++, titleKey: lc.titleKey, descKey: lc.descKey,
                category: 'emotional', icon: lc.icon, durationKey: 'ch_duration_7days',
                difficulty: 'easy', progress: 0, total: 7, status: 'active', assignedTo: 'both'
            });
        }

        const annoyances = a.house_annoy || [];
        if (annoyances.includes('bathroom') || annoyances.includes('mess')) {
            challenges.push({
                id: 'ch' + id++, titleKey: 'ch_clean_zone', descKey: 'ch_clean_zone_desc',
                category: 'household', icon: '\u2728', durationKey: 'ch_duration_7days',
                difficulty: 'easy', progress: 0, total: 7, status: 'active', assignedTo: 'partner'
            });
        }
        if (annoyances.includes('phone')) {
            challenges.push({
                id: 'ch' + id++, titleKey: 'ch_screen_free', descKey: 'ch_screen_free_desc',
                category: 'quality_time', icon: '\uD83D\uDCF5', durationKey: 'ch_duration_1evening',
                difficulty: 'medium', progress: 0, total: 1, status: 'active', assignedTo: 'both'
            });
        }

        challenges.push({
            id: 'ch' + id++, titleKey: 'ch_bedtime_gratitude', descKey: 'ch_bedtime_gratitude_desc',
            category: 'appreciation', icon: '\uD83C\uDF19', durationKey: 'ch_duration_5days',
            difficulty: 'easy', progress: 0, total: 5, status: 'suggested', assignedTo: 'both'
        });

        if (['sometimes', 'often', 'constant'].includes(a.fin_tension)) {
            challenges.push({
                id: 'ch' + id++, titleKey: 'ch_finance_evening', descKey: 'ch_finance_evening_desc',
                category: 'finances', icon: '\uD83D\uDCB0', durationKey: 'ch_duration_onetime',
                difficulty: 'hard', progress: 0, total: 1, status: 'suggested', assignedTo: 'both'
            });
        }

        return challenges.length > 0 ? challenges : DEMO_CHALLENGES;
    },

    copyInviteCode() {
        const code = this.state.inviteCode || document.getElementById('invite-code').textContent;
        navigator.clipboard.writeText(code).then(() => {
            const btns = document.querySelectorAll('.btn-copy');
            btns.forEach(btn => { btn.textContent = this.t('btn_copied'); });
            setTimeout(() => btns.forEach(btn => { btn.textContent = this.t('btn_copy'); }), 2000);
        }).catch(() => {});
    },

    shareInvite() {
        const code = this.state.inviteCode || 'KOOPLE-A7X9';
        const text = 'Join me on Koople — AI mediator for couples! My code: ' + code;
        if (navigator.share) navigator.share({ title: 'Koople', text });
        else this.copyInviteCode();
    },

    goToDashboard() { this.navigate('dashboard'); },

    // ── Dashboard ─────────────────────────────────────────
    renderDashboard() {
        this.applyTranslations();
        const name = this.state.user.name || 'User';
        const analysis = this.state.analysis || this.getDefaultAnalysis();
        const harmony = analysis.harmony;

        document.getElementById('greeting').textContent = this.t('greeting_prefix') + ' ' + name + '!';
        document.getElementById('user-avatar').textContent = name.charAt(0).toUpperCase();

        // Weekly focus — based on lowest scoring area
        const greetingSub = document.getElementById('greeting-sub');
        if (greetingSub) {
            const cats = analysis.catScores;
            const focusAreas = ['communication', 'household', 'emotional', 'quality_time', 'intimacy', 'finances'];
            const lowestCat = focusAreas.sort((a, b) => (cats[a] || 70) - (cats[b] || 70))[0];
            greetingSub.textContent = this.t('dashboard_subtitle').split(':')[0] + ': ' + this.t('cat_' + lowestCat);
        }

        // Health score
        const scoreValue = document.getElementById('health-score-value');
        if (scoreValue) scoreValue.textContent = harmony;
        const scoreCircle = document.querySelector('.score-circle');
        if (scoreCircle) {
            const circumference = 2 * Math.PI * 42;
            const offset = circumference * (1 - harmony / 100);
            scoreCircle.setAttribute('stroke-dasharray', circumference.toFixed(2));
            scoreCircle.setAttribute('stroke-dashoffset', offset.toFixed(2));
        }

        // Score details
        const detailsEl = document.getElementById('score-details');
        if (detailsEl) {
            const cats = analysis.catScores;
            const topCats = [
                { name: this.t('cat_communication'), val: cats.communication },
                { name: this.t('cat_household'), val: cats.household },
                { name: this.t('cat_intimacy'), val: cats.intimacy },
                { name: this.t('cat_quality_time'), val: cats.quality_time }
            ];
            detailsEl.innerHTML = topCats.map(c => {
                const dotClass = c.val >= 80 ? 'dot-green' : c.val >= 65 ? 'dot-yellow' : 'dot-orange';
                return '<div class="score-detail"><span class="dot ' + dotClass + '"></span> ' + c.name + ': ' + c.val + '%</div>';
            }).join('');
        }

        // Active challenges
        const challengesEl = document.getElementById('active-challenges');
        const active = this.state.challenges.filter(c => c.status === 'active');
        if (active.length === 0) {
            challengesEl.innerHTML = '<p style="color:var(--text-tertiary);font-size:14px;padding:12px">' + this.t('no_active_challenges') + '</p>';
        } else {
            challengesEl.innerHTML = active.map(ch => `
                <div class="challenge-card-mini" onclick="app.showChallengeDetail('${ch.id}')">
                    <div class="challenge-mini-icon">${ch.icon}</div>
                    <div class="challenge-mini-info">
                        <h4>${this.ct(ch, 'title')}</h4>
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

        // Activity feed — only real partner achievements
        const feedEl = document.getElementById('activity-feed');
        if (this.state.activity.length === 0) {
            feedEl.innerHTML = '<p style="color:var(--text-tertiary);font-size:14px;padding:12px;text-align:center">' + this.t('empty_completed') + '</p>';
        } else {
            feedEl.innerHTML = this.state.activity.slice(0, 6).map(a => `
                <div class="activity-item">
                    <span class="activity-icon">${a.icon}</span>
                    <div class="activity-content">
                        <p>${a.text}</p>
                        <span class="activity-time">${a.time}</span>
                    </div>
                </div>
            `).join('');
        }
    },

    // ── Challenges ────────────────────────────────────────
    renderChallenges() {
        this.applyTranslations();
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
        const partnerName = this.state.user.partnerName || 'Partner';

        if (filtered.length === 0) {
            const msgs = {
                active: this.t('empty_active'),
                completed: this.t('empty_completed'),
                suggested: this.t('empty_suggested')
            };
            list.innerHTML = '<div class="empty-state"><span class="empty-icon">' + (tab === 'completed' ? '\uD83C\uDFC6' : '\uD83C\uDF31') + '</span><p>' + msgs[tab] + '</p></div>';
            return;
        }

        list.innerHTML = filtered.map(ch => {
            const pct = (ch.progress / ch.total) * 100;
            const diff = { easy: '\uD83D\uDFE2 ' + this.t('diff_easy'), medium: '\uD83D\uDFE1 ' + this.t('diff_medium'), hard: '\uD83D\uDFE0 ' + this.t('diff_hard') }[ch.difficulty];
            const assigned = { both: this.t('assigned_both'), user: this.t('assigned_user'), partner: this.t('assigned_partner_prefix') + ' ' + partnerName }[ch.assignedTo];
            const title = this.ct(ch, 'title');
            const desc = this.ct(ch, 'desc') || this.ct(ch, 'description');
            const dur = this.ct(ch, 'duration');

            return `
                <div class="challenge-card" onclick="app.showChallengeDetail('${ch.id}')">
                    <div class="challenge-card-header">
                        <span class="challenge-icon-large">${ch.icon}</span>
                        <div>
                            <h3>${title}</h3>
                            <div class="challenge-meta"><span>${diff}</span><span>\u00B7</span><span>${dur}</span><span>\u00B7</span><span>${assigned}</span></div>
                        </div>
                    </div>
                    <p class="challenge-desc">${desc}</p>
                    ${ch.status !== 'suggested' ? `
                        <div class="challenge-progress">
                            <div class="progress-bar"><div class="progress-fill ${ch.status === 'completed' ? 'complete' : ''}" style="width:${pct}%"></div></div>
                            <span class="progress-label">${ch.progress} / ${ch.total}</span>
                        </div>
                    ` : `
                        <button class="btn btn-secondary btn-small" onclick="event.stopPropagation(); app.acceptChallenge('${ch.id}')">${this.t('btn_accept')}</button>
                    `}
                </div>`;
        }).join('');
    },

    showChallengeDetail(id) {
        const ch = this.state.challenges.find(c => c.id === id);
        if (!ch) return;
        const pct = (ch.progress / ch.total) * 100;
        const title = this.ct(ch, 'title');
        const desc = this.ct(ch, 'desc') || this.ct(ch, 'description');
        const dur = this.ct(ch, 'duration');
        this.showModal(`
            <div class="challenge-detail">
                <div class="challenge-detail-icon">${ch.icon}</div>
                <h2>${title}</h2>
                <p>${desc}</p>
                <div class="challenge-detail-meta">
                    <div class="meta-item"><span class="meta-label">${this.t('duration_label')}</span><span class="meta-value">${dur}</span></div>
                    <div class="meta-item"><span class="meta-label">${this.t('progress_label')}</span><span class="meta-value">${ch.progress}/${ch.total}</span></div>
                </div>
                <div class="challenge-progress" style="margin-top:16px">
                    <div class="progress-bar"><div class="progress-fill ${ch.status === 'completed' ? 'complete' : ''}" style="width:${pct}%"></div></div>
                </div>
                ${ch.status === 'active' ? '<button class="btn btn-primary btn-large" style="margin-top:20px" onclick="app.markChallengeDay(\'' + ch.id + '\')">\u2705 ' + this.t('btn_mark_done') + '</button>' : ''}
                ${ch.status === 'suggested' ? '<button class="btn btn-primary btn-large" style="margin-top:20px" onclick="app.acceptChallenge(\'' + ch.id + '\')">' + this.t('btn_accept') + '</button>' : ''}
                ${ch.status === 'completed' ? '<div style="margin-top:20px;color:var(--secondary);font-weight:600">\uD83C\uDFC6 ' + this.t('challenge_done') + '</div>' : ''}
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
                    ? '\uD83C\uDFC6 ' + this.ct(ch, 'title') + ' — ' + this.t('challenge_done')
                    : '\u2705 ' + this.ct(ch, 'title') + ' (' + ch.progress + '/' + ch.total + ')'
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
            this.addActivity('new_challenge', '\uD83C\uDF1F', '\uD83C\uDF1F ' + this.t('btn_accept') + ': ' + this.ct(ch, 'title'));
            this.save();
        }
        this.closeModal();
        if (this.state.currentScreen === 'challenges') this.renderChallenges();
    },

    addActivity(type, icon, text) {
        this.state.activity.unshift({ type, icon, text, time: this.t('just_now') });
        if (this.state.activity.length > 20) this.state.activity.pop();
    },

    // ── AI Mediator Chat ──────────────────────────────────
    renderChat() {
        this.applyTranslations();
        // Ensure welcome message exists
        if (this.state.chat.messages.length === 0) {
            this.state.chat.messages.push({
                role: 'assistant',
                text: this.t('chat_welcome'),
                time: this.t('chat_time_now')
            });
        }

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

        const sugEl = document.getElementById('chat-suggestions');
        if (sugEl) sugEl.style.display = this.state.chat.messages.length <= 1 ? 'flex' : 'none';
    },

    async sendMessage() {
        const input = document.getElementById('chat-input');
        const text = input.value.trim();
        if (!text || this.state.chat.isTyping) return;

        this.state.chat.messages.push({ role: 'user', text, time: this.getCurrentTime() });
        input.value = '';
        input.style.height = 'auto';
        this.state.chat.isTyping = true;
        this.renderChat();

        // Try real AI first, fall back to pattern matching
        if (typeof AIService !== 'undefined' && AIService.isConfigured()) {
            const analysis = this.state.analysis || this.getDefaultAnalysis();
                const userContext = {
                    language: (typeof getLang === 'function') ? getLang() : 'en',
                    name: this.state.user.name,
                    partnerName: this.state.user.partnerName,
                    duration: this.state.user.duration,
                    livingTogether: this.state.user.livingTogether,
                    harmonyScore: analysis.harmony,
                    growthAreas: analysis.growth ? analysis.growth.join(', ') : '',
                    strengths: analysis.strengths ? analysis.strengths.join(', ') : '',
                    questionnaireSummary: (typeof AIService !== 'undefined' && AIService.buildQuestionnaireSummary)
                        ? AIService.buildQuestionnaireSummary(this.state.onboarding.answers, analysis.catScores)
                        : ''
                };
            const result = await AIService.chat(this.state.chat.messages.slice(0, -1).concat([{role:'user', text}]), userContext);
            this.state.chat.isTyping = false;
            this.state.chat.messages.push({ role: 'assistant', text: result.text, time: this.getCurrentTime() });
            if (result.challenge) {
                this.state.challenges.push(result.challenge);
                this.addActivity('new_challenge', '\uD83C\uDF1F', 'AI: ' + result.challenge.title);
                    this.showNotification('\uD83C\uDFAF ' + result.challenge.title);
            }
            this.renderChat();
            this.save();
        } else {
            // Fallback: pattern-based responses
            const delay = 1200 + Math.random() * 1500;
            setTimeout(() => {
                this.state.chat.isTyping = false;
                const response = this.generateAIResponse(text);
                this.state.chat.messages.push({ role: 'assistant', text: response.text, time: this.getCurrentTime() });
                if (response.challenge) {
                    this.state.challenges.push(response.challenge);
                    this.addActivity('new_challenge', '\uD83C\uDF1F', 'AI: ' + response.challenge.title);
                }
                this.renderChat();
                this.save();
            }, delay);
        }
    },

    sendSuggestion(btn) {
        document.getElementById('chat-input').value = btn.textContent;
        this.sendMessage();
    },

    generateAIResponse(userMessage) {
        const lower = userMessage.toLowerCase();
        const pn = this.state.user.partnerName || 'partner';
        let challenge = null;

        if (/bathroom|hair|clean|ванн|волос|чистот/i.test(lower)) {
            challenge = this.createChallengeFromChat('Clean Bathroom', 'After each use: remove hair, wipe sink, hang towel.', 'household', '\u2728', 7, 'partner');
            return { text: this.t('chat_welcome').includes('Привет') ?
                'Понимаю — волосы в ванной это классика микроконфликтов. Я создал челлендж «Clean Bathroom» для ' + pn + '.' :
                'I understand — bathroom cleanliness is a classic micro-conflict. I\'ve created a "Clean Bathroom" challenge for ' + pn + '.', challenge };
        }

        if (/dish|dirty|plate|посуд|грязн/i.test(lower)) {
            challenge = this.createChallengeFromChat('Clean Sink Rule', 'Wash dishes right after meals.', 'household', '\uD83E\uDDFD', 5, 'partner');
            return { text: 'I\'ve prepared a "Clean Sink Rule" challenge for ' + pn + '. Small habits, big impact!', challenge };
        }

        if (/phone|screen|gadget|телефон|экран|гаджет/i.test(lower)) {
            challenge = this.createChallengeFromChat('Screen-Free Evenings', 'Phones away from 8-9pm. Talk, play, or just be together.', 'quality_time', '\uD83D\uDCF5', 5, 'both');
            return { text: 'I\'ve created a "Screen-Free Evenings" challenge for both of you. One hour without phones each evening.', challenge };
        }

        if (/ignor|notice|attention|не.*заме|игнор|невидим/i.test(lower)) {
            challenge = this.createChallengeFromChat('Daily Check-in', 'Ask your partner "How are you feeling today?" and listen without advice.', 'emotional', '\uD83D\uDC96', 7, 'partner');
            return { text: 'Feeling unseen is painful. I\'ve created a "Daily Check-in" challenge for ' + pn + ' to build a habit of attentiveness.', challenge };
        }

        if (/boring|routine|monoton|скуч|однообраз|рутин/i.test(lower)) {
            challenge = this.createChallengeFromChat('Surprise Week', 'Every 2 days — a little surprise: unusual dinner, a note, unexpected walk.', 'quality_time', '\uD83C\uDF39', 3, 'both');
            return { text: 'I\'ve prepared a "Surprise Week" for both of you — 3 little surprises this week to break the routine.', challenge };
        }

        if (/money|financ|budget|spend|денег|деньги|финанс|бюджет/i.test(lower)) {
            challenge = this.createChallengeFromChat('Finance Evening', 'A calm conversation about finances. Rules: no blame, focus on shared goals.', 'finances', '\uD83D\uDCB0', 1, 'both');
            return { text: 'Finances are one of the hardest topics for couples. I\'ve created a structured "Finance Evening" challenge.', challenge };
        }

        if (/fight|conflict|argue|yell|ссор|конфликт|ругаемся|крич/i.test(lower)) {
            return { text: 'Conflicts are inevitable in relationships, but it\'s HOW you handle them that matters.\n\nTell me more: how do your conflicts usually start? Who initiates, and how do they end?', challenge: null };
        }

        if (/thank|grateful|appreciate|спасибо|благодар|ценю/i.test(lower)) {
            return { text: 'That\'s wonderful! Gratitude is a powerful relationship tool. Would you like to send ' + pn + ' a warm notification?', challenge: null };
        }

        const analysis = this.state.analysis;
        let extra = '';
        if (analysis && analysis.growth.length > 0) {
            extra = '\n\nBy the way, based on your questionnaire I see growth areas: ' + analysis.growth.join(', ') + '. Want to discuss any of these?';
        }
        return { text: 'Thank you for sharing. I can:\n\n1. \uD83D\uDCAC Discuss the situation in detail\n2. \uD83C\uDFAF Create a gentle challenge for ' + pn + '\n3. \uD83D\uDCA1 Give a recommendation for both of you\n4. \u2764\uFE0F Send ' + pn + ' something nice\n\nWhat would you prefer?' + extra, challenge: null };
    },

    createChallengeFromChat(title, desc, category, icon, total, assignedTo) {
        return {
            id: 'ch_' + Date.now(),
            title, description: desc, category, icon,
            duration: total === 1 ? 'One-time' : total + ' days',
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
        this.applyTranslations();
        const analysis = this.state.analysis || this.getDefaultAnalysis();
        const cats = analysis.catScores;

        const strengthsEl = document.getElementById('strengths-list');
        const growthEl = document.getElementById('growth-list');
        if (strengthsEl) strengthsEl.innerHTML = analysis.strengths.map(s => '<li class="positive">' + s + '</li>').join('');
        if (growthEl) growthEl.innerHTML = analysis.growth.map(g => '<li class="growth">' + g + '</li>').join('');

        const barsEl = document.getElementById('compatibility-bars');
        barsEl.innerHTML = Object.entries(cats).map(([key, val]) => {
            const color = val >= 80 ? '#6BBF8A' : val >= 65 ? '#E5C76B' : '#D4A574';
            return '<div class="compat-row"><span class="compat-name">' + this.t('cat_' + key) + '</span><div class="compat-bar-track"><div class="compat-bar-fill" style="width:' + val + '%;background:' + color + '"></div></div><span class="compat-value">' + val + '%</span></div>';
        }).join('');

        const chartEl = document.getElementById('progress-chart');
        const base = Math.max(50, analysis.harmony - 15);
        const weeks = [this.t('week_1'), this.t('week_2'), this.t('week_3'), this.t('week_now')];
        const vals = [base, base + 5, base + 9, analysis.harmony];
        chartEl.innerHTML = '<div class="bar-chart">' + weeks.map((w, i) =>
            '<div class="bar-col"><div class="bar" style="height:' + vals[i] + '%"><span class="bar-value">' + vals[i] + '</span></div><span class="bar-label">' + w + '</span></div>'
        ).join('') + '</div>';

        const recs = analysis.recommendations || AI_RECOMMENDATIONS;
        const recsEl = document.getElementById('ai-recommendations');
        recsEl.innerHTML = recs.map(rec =>
            '<div class="recommendation-card priority-' + rec.priority + '"><h4>' + this.ct(rec, 'title') + '</h4><p>' + this.ct(rec, 'text') + '</p></div>'
        ).join('');
    },

    // ── Wishes ────────────────────────────────────────────
    renderWishes() {
        this.applyTranslations();
        const listEl = document.getElementById('wishes-list');
        if (this.state.wishes.length === 0) {
            listEl.innerHTML = '<div class="empty-state"><span class="empty-icon">\u2B50</span><p>' + this.t('wishes_empty') + '</p></div>';
            return;
        }
        listEl.innerHTML = this.state.wishes.map(w => {
            const label = { active: this.t('wish_active'), in_progress: this.t('wish_in_progress'), fulfilled: this.t('wish_fulfilled') }[w.status] || this.t('wish_active');
            const wText = this.ct(w, 'text');
            const wDate = this.ct(w, 'createdAt');
            return '<div class="wish-card"><p class="wish-text">' + wText + '</p><div class="wish-meta"><span class="wish-status ' + w.status + '">' + label + '</span><span class="wish-date">' + wDate + '</span></div></div>';
        }).join('');
    },

    showAddWishModal() {
        this.showModal(`
            <h2>${this.t('add_wish_title')}</h2>
            <p class="modal-subtitle">${this.t('add_wish_subtitle')}</p>
            <textarea id="wish-input" class="answer-textarea" placeholder="${this.t('wish_placeholder')}" rows="4"></textarea>
            <div class="modal-category-select">
                <label>${this.t('wish_cat_label')}</label>
                <select id="wish-category">
                    <option value="emotional">${this.t('wishcat_emotional')}</option>
                    <option value="household">${this.t('wishcat_household')}</option>
                    <option value="quality_time">${this.t('wishcat_quality_time')}</option>
                    <option value="intimacy">${this.t('wishcat_intimacy')}</option>
                    <option value="communication">${this.t('wishcat_communication')}</option>
                </select>
            </div>
            <button class="btn btn-primary btn-large" onclick="app.addWish()">${this.t('btn_send_wish')}</button>
        `);
    },

    addWish() {
        const text = document.getElementById('wish-input').value.trim();
        const category = document.getElementById('wish-category').value;
        if (!text) return;
        this.state.wishes.unshift({ id: 'w' + Date.now(), text, category, status: 'active', createdAt: this.t('just_now') });
        this.addActivity('wish', '\u2B50', this.t('activity_wish_added'));
        this.save();
        this.closeModal();
        this.renderWishes();
        this.showNotification('\u2B50 ' + this.t('wish_added_notif'));
    },

    // ── Profile ───────────────────────────────────────────
    renderProfile() {
        this.applyTranslations();
        const name = this.state.user.name || 'User';
        const partnerName = this.state.user.partnerName || 'Partner';
        const analysis = this.state.analysis || this.getDefaultAnalysis();

        document.getElementById('profile-name').textContent = name;
        document.getElementById('profile-partner-name').textContent = partnerName;
        document.querySelector('.profile-avatar-large').textContent = name.charAt(0).toUpperCase();

        const completed = this.state.challenges.filter(c => c.status === 'completed').length;
        const statsEl = document.getElementById('profile-stats');
        if (statsEl) {
            statsEl.innerHTML = `
                <div class="stat"><span class="stat-value">${this.state.challenges.length}</span><span class="stat-label">${this.t('profile_challenges')}</span></div>
                <div class="stat"><span class="stat-value">${completed}</span><span class="stat-label">${this.t('profile_completed')}</span></div>
                <div class="stat"><span class="stat-value">${analysis.harmony}%</span><span class="stat-label">${this.t('profile_harmony')}</span></div>
            `;
        }
    },

    showInvitePartner() {
        const code = this.state.inviteCode || 'KOOPLE-A7X9';
        this.showModal(`
            <h2>${this.t('invite_title')}</h2>
            <p>${this.t('invite_text')}</p>
            <div class="invite-code-box">
                <span class="invite-code">${code}</span>
                <button class="btn-copy" onclick="app.copyInviteCode()">${this.t('btn_copy')}</button>
            </div>
            <button class="btn btn-secondary btn-large" onclick="app.shareInvite()">${this.t('btn_share_invite')}</button>
        `);
    },

    showLanguageModal() {
        const currentCode = (typeof getLang === 'function') ? getLang() : 'en';
        const langs = (typeof LANGUAGES !== 'undefined') ? LANGUAGES : [{ code: 'en', name: 'English', flag: '' }];
        const langRows = langs.map(l =>
            '<button class="lang-row ' + (l.code === currentCode ? 'active' : '') + '" onclick="app.changeLanguage(\'' + l.code + '\')">' +
            '<span class="lang-flag">' + l.flag + '</span><span>' + l.name + '</span>' +
            (l.code === currentCode ? '<span class="lang-check">\u2713</span>' : '') +
            '</button>'
        ).join('');
        this.showModal(`
            <h2>${this.t('settings_language')}</h2>
            <div class="lang-grid-modal">${langRows}</div>
        `);
    },

    changeLanguage(code) {
        if (typeof setLang === 'function') setLang(code);
        try { localStorage.setItem('koople_lang', code); } catch(e) {}
        this.closeModal();
        this.applyTranslations();
        this.renderProfile();
    },

    showAISettings() {
        const configured = (typeof AIService !== 'undefined') && AIService.isConfigured();
        const config = configured ? AIService.getConfig() : { provider: 'openai', hasKey: false };
        this.showModal(`
            <h2>${this.t('ai_settings_title')}</h2>
            <p class="modal-subtitle">${this.t('ai_settings_desc')}</p>
            ${configured ? '<div class="ai-status connected">\u2705 ' + this.t('ai_connected') + ' (' + config.provider + ')</div>' : '<div class="ai-status">\uD83D\uDCA4 ' + this.t('ai_demo_mode') + '</div>'}
            <div class="ai-config-form">
                <div class="input-group">
                    <label>${this.t('ai_provider_label')}</label>
                    <select id="ai-provider">
                        <option value="openai" ${config.provider === 'openai' ? 'selected' : ''}>OpenAI (GPT-4o-mini)</option>
                        <option value="anthropic" ${config.provider === 'anthropic' ? 'selected' : ''}>Anthropic (Claude)</option>
                    </select>
                </div>
                <div class="input-group">
                    <label>${this.t('ai_api_key_label')}</label>
                    <input type="password" id="ai-api-key" placeholder="${this.t('ai_api_key_placeholder')}" value="">
                </div>
                <button class="btn btn-primary btn-large" onclick="app.connectAI()">${this.t('btn_connect_ai')}</button>
                ${configured ? '<button class="btn btn-ghost" style="margin-top:8px;color:#D48A8A" onclick="app.disconnectAI()">' + this.t('ai_disconnect') + '</button>' : ''}
            </div>
        `);
    },

    connectAI() {
        const provider = document.getElementById('ai-provider').value;
        const apiKey = document.getElementById('ai-api-key').value.trim();
        if (!apiKey) return;
        if (typeof AIService !== 'undefined') {
            AIService.configure(provider, apiKey);
            AIService.saveConfig();
        }
        this.closeModal();
        this.showNotification('\u2705 ' + this.t('ai_connected'));
    },

    disconnectAI() {
        if (typeof AIService !== 'undefined') AIService.disconnect();
        this.closeModal();
        this.showNotification(this.t('ai_demo_mode'));
    },

    showNotifications() {
        this.showModal(`
            <h2>${this.t('notif_title')}</h2>
            <p class="modal-subtitle">${this.t('notif_subtitle')}</p>
            <div class="settings-toggles">
                <div class="setting-row"><span>${this.t('notif_challenges')}</span><span class="toggle-pill active" onclick="this.classList.toggle('active')"></span></div>
                <div class="setting-row"><span>${this.t('notif_recommendations')}</span><span class="toggle-pill active" onclick="this.classList.toggle('active')"></span></div>
                <div class="setting-row"><span>${this.t('notif_partner')}</span><span class="toggle-pill active" onclick="this.classList.toggle('active')"></span></div>
                <div class="setting-row"><span>${this.t('notif_weekly')}</span><span class="toggle-pill" onclick="this.classList.toggle('active')"></span></div>
            </div>
        `);
    },

    showPrivacy() {
        this.showModal(`
            <h2>${this.t('privacy_title')}</h2>
            <p class="modal-subtitle">${this.t('privacy_subtitle')}</p>
            <div class="privacy-info">
                <div class="privacy-item"><strong>${this.t('privacy_chat_title')}</strong><p>${this.t('privacy_chat_text')}</p></div>
                <div class="privacy-item"><strong>${this.t('privacy_survey_title')}</strong><p>${this.t('privacy_survey_text')}</p></div>
                <div class="privacy-item"><strong>${this.t('privacy_complaints_title')}</strong><p>${this.t('privacy_complaints_text')}</p></div>
            </div>
            <button class="btn btn-secondary btn-large" onclick="app.resetData()" style="margin-top:20px;color:#D48A8A">${this.t('btn_delete_data')}</button>
        `);
    },

    showHelp() {
        this.showModal(`
            <h2>${this.t('help_title')}</h2>
            <div class="help-steps">
                <div class="help-step"><span class="help-num">1</span><div><strong>${this.t('help_1_title')}</strong><p>${this.t('help_1_text')}</p></div></div>
                <div class="help-step"><span class="help-num">2</span><div><strong>${this.t('help_2_title')}</strong><p>${this.t('help_2_text')}</p></div></div>
                <div class="help-step"><span class="help-num">3</span><div><strong>${this.t('help_3_title')}</strong><p>${this.t('help_3_text')}</p></div></div>
                <div class="help-step"><span class="help-num">4</span><div><strong>${this.t('help_4_title')}</strong><p>${this.t('help_4_text')}</p></div></div>
                <div class="help-step"><span class="help-num">5</span><div><strong>${this.t('help_5_title')}</strong><p>${this.t('help_5_text')}</p></div></div>
            </div>
        `);
    },

    showAbout() {
        this.showModal(`
            <div style="text-align:center">
                <h2>Koople</h2>
                <p>${this.t('app_tagline')}</p>
                <p style="color:var(--text-secondary);margin-top:12px">${this.t('about_version')}</p>
                <p style="color:var(--text-secondary);margin-top:8px">${this.t('about_desc')}</p>
            </div>
        `);
    },

    // ── Complaint & Appreciation ──────────────────────────
    showComplaintModal() {
        this.showModal(`
            <h2>\uD83D\uDE14 ${this.t('complaint_title')}</h2>
            <p class="modal-subtitle">${this.t('complaint_subtitle')}</p>
            <textarea id="complaint-input" class="answer-textarea" placeholder="${this.t('complaint_placeholder')}" rows="4"></textarea>
            <div class="complaint-urgency">
                <label>${this.t('urgency_label')}</label>
                <div class="urgency-options">
                    <button class="urgency-btn" onclick="app.selectUrgency(this, 'low')">${this.t('urgency_low')}</button>
                    <button class="urgency-btn active" onclick="app.selectUrgency(this, 'medium')">${this.t('urgency_medium')}</button>
                    <button class="urgency-btn" onclick="app.selectUrgency(this, 'high')">${this.t('urgency_high')}</button>
                </div>
            </div>
            <input type="hidden" id="complaint-urgency" value="medium">
            <button class="btn btn-primary btn-large" onclick="app.submitComplaint()">${this.t('btn_send_mediator')}</button>
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
        this.addActivity('complaint', '\uD83D\uDE14', this.t('activity_concern_shared'));
        this.closeModal();

        this.state.chat.messages.push({ role: 'user', text: text, time: this.getCurrentTime() });
        this.navigate('mediator');

        this.state.chat.isTyping = true;
        this.renderChat();
        setTimeout(() => {
            this.state.chat.isTyping = false;
            const response = this.generateAIResponse(text);
            this.state.chat.messages.push({ role: 'assistant', text: response.text, time: this.getCurrentTime() });
            if (response.challenge) {
                this.state.challenges.push(response.challenge);
                this.addActivity('new_challenge', '\uD83C\uDF1F', 'AI: ' + response.challenge.title);
            }
            this.renderChat();
            this.save();
        }, 2000);
    },

    showAppreciationModal() {
        const pn = this.state.user.partnerName || 'partner';
        this.showModal(`
            <h2>\uD83D\uDC95 ${this.t('appreciation_title_prefix')} ${pn}</h2>
            <p class="modal-subtitle">${this.t('appreciation_subtitle')}</p>
            <textarea id="appreciation-input" class="answer-textarea" placeholder="${this.t('appreciation_placeholder')}" rows="4"></textarea>
            <button class="btn btn-primary btn-large" onclick="app.submitAppreciation()">${this.t('btn_send')}</button>
        `);
    },

    submitAppreciation() {
        const text = document.getElementById('appreciation-input').value.trim();
        if (!text) return;
        this.state.appreciations.push({ text, time: new Date().toISOString() });
        this.addActivity('appreciation', '\uD83D\uDC95', this.t('activity_praised') + ' ' + (this.state.user.partnerName || 'partner'));
        this.save();
        this.closeModal();
        this.showNotification('\uD83D\uDC95 ' + this.t('appreciation_sent'));
    },

    // ── Modals ────────────────────────────────────────────
    showModal(content) {
        const overlay = document.getElementById('modal-overlay');
        const modal = document.getElementById('modal-content');
        modal.innerHTML = '<button class="modal-close" onclick="app.closeModal()">&times;</button>' + content;
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
