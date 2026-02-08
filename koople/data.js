// ============================================================
// KOOPLE — Questionnaire Data & App Models
// Based on family psychotherapy research:
// Gottman Method, EFT (Emotionally Focused Therapy),
// Imago Relationship Therapy, Nonviolent Communication
// ============================================================

const QUESTIONNAIRE = [
    // ─────────────────────────────────────────────────
    // Category 1: Эмоциональная связь (Emotional Bond)
    // ─────────────────────────────────────────────────
    {
        id: 'emo_love_language',
        category: 'emotional',
        categoryName: 'Эмоциональная связь',
        categoryIcon: '\u2764\uFE0F',
        text: 'Какой «язык любви» вам ближе всего?',
        hint: 'Как вы чувствуете любовь сильнее всего?',
        type: 'single',
        options: [
            { value: 'words', label: 'Слова поддержки и комплименты' },
            { value: 'touch', label: 'Прикосновения и объятия' },
            { value: 'time', label: 'Качественное время вместе' },
            { value: 'gifts', label: 'Подарки и знаки внимания' },
            { value: 'service', label: 'Помощь и забота в делах' }
        ]
    },
    {
        id: 'emo_express',
        category: 'emotional',
        categoryName: 'Эмоциональная связь',
        categoryIcon: '\u2764\uFE0F',
        text: 'Насколько легко вам говорить партнёру о своих чувствах?',
        hint: 'Оцените от 1 до 5',
        type: 'scale',
        scaleLabels: ['Очень сложно', 'Сложно', 'Средне', 'Легко', 'Очень легко']
    },
    {
        id: 'emo_need_closeness',
        category: 'emotional',
        categoryName: 'Эмоциональная связь',
        categoryIcon: '\u2764\uFE0F',
        text: 'Как часто вам нужно слышать «я тебя люблю» или получать подтверждение чувств?',
        hint: '',
        type: 'single',
        options: [
            { value: 'daily', label: 'Каждый день, это важно' },
            { value: 'often', label: 'Несколько раз в неделю' },
            { value: 'sometimes', label: 'Иногда, когда к месту' },
            { value: 'rarely', label: 'Редко, я и так знаю' },
            { value: 'actions', label: 'Мне важнее действия, не слова' }
        ]
    },
    {
        id: 'emo_missing',
        category: 'emotional',
        categoryName: 'Эмоциональная связь',
        categoryIcon: '\u2764\uFE0F',
        text: 'Чего вам больше всего не хватает в эмоциональной близости с партнёром?',
        hint: 'Выберите все подходящие варианты',
        type: 'multi',
        options: [
            { value: 'attention', label: 'Внимания к моим переживаниям' },
            { value: 'support', label: 'Поддержки в трудные моменты' },
            { value: 'romance', label: 'Романтики и сюрпризов' },
            { value: 'deep_talk', label: 'Глубоких разговоров' },
            { value: 'humor', label: 'Совместного смеха и лёгкости' },
            { value: 'nothing', label: 'Всего хватает' }
        ]
    },

    // ─────────────────────────────────────────────────
    // Category 2: Коммуникация (Communication)
    // ─────────────────────────────────────────────────
    {
        id: 'comm_conflict_style',
        category: 'communication',
        categoryName: 'Коммуникация',
        categoryIcon: '\uD83D\uDCAC',
        text: 'Как вы обычно ведёте себя во время конфликта?',
        hint: 'Выберите наиболее типичную реакцию',
        type: 'single',
        options: [
            { value: 'discuss', label: 'Стараюсь обсудить спокойно' },
            { value: 'emotional', label: 'Становлюсь эмоциональным/-ой' },
            { value: 'withdraw', label: 'Замыкаюсь и молчу' },
            { value: 'avoid', label: 'Стараюсь избежать конфликта' },
            { value: 'humor', label: 'Пытаюсь разрядить юмором' }
        ]
    },
    {
        id: 'comm_criticism',
        category: 'communication',
        categoryName: 'Коммуникация',
        categoryIcon: '\uD83D\uDCAC',
        text: 'Как вы реагируете на критику от партнёра?',
        hint: '',
        type: 'single',
        options: [
            { value: 'listen', label: 'Слушаю и стараюсь понять' },
            { value: 'defend', label: 'Начинаю защищаться' },
            { value: 'hurt', label: 'Обижаюсь и отдаляюсь' },
            { value: 'counter', label: 'Критикую в ответ' },
            { value: 'depends', label: 'Зависит от формулировки' }
        ]
    },
    {
        id: 'comm_unspoken',
        category: 'communication',
        categoryName: 'Коммуникация',
        categoryIcon: '\uD83D\uDCAC',
        text: 'Есть ли темы, которые вы избегаете обсуждать с партнёром?',
        hint: 'Выберите все подходящие',
        type: 'multi',
        options: [
            { value: 'money', label: 'Финансы и траты' },
            { value: 'sex', label: 'Интимная жизнь' },
            { value: 'family', label: 'Отношения с родственниками' },
            { value: 'future', label: 'Планы на будущее' },
            { value: 'habits', label: 'Раздражающие привычки' },
            { value: 'past', label: 'Прошлые отношения' },
            { value: 'feelings', label: 'Глубокие чувства и страхи' },
            { value: 'none', label: 'Мы обсуждаем всё' }
        ]
    },
    {
        id: 'comm_after_fight',
        category: 'communication',
        categoryName: 'Коммуникация',
        categoryIcon: '\uD83D\uDCAC',
        text: 'Как обычно завершаются ваши конфликты?',
        hint: '',
        type: 'single',
        options: [
            { value: 'resolve', label: 'Обсуждаем и находим решение' },
            { value: 'one_gives', label: 'Кто-то уступает ради мира' },
            { value: 'forget', label: 'Просто забываем и идём дальше' },
            { value: 'linger', label: 'Напряжение остаётся надолго' },
            { value: 'escalate', label: 'Часто перерастают в большие ссоры' }
        ]
    },

    // ─────────────────────────────────────────────────
    // Category 3: Быт и дом (Household & Daily Life)
    // ─────────────────────────────────────────────────
    {
        id: 'house_chores',
        category: 'household',
        categoryName: 'Быт и дом',
        categoryIcon: '\uD83C\uDFE0',
        text: 'Как вы оцениваете распределение домашних обязанностей?',
        hint: '',
        type: 'scale',
        scaleLabels: ['Совсем нечестно', 'Скорее нечестно', 'Более-менее', 'Справедливо', 'Идеально']
    },
    {
        id: 'house_annoy',
        category: 'household',
        categoryName: 'Быт и дом',
        categoryIcon: '\uD83C\uDFE0',
        text: 'Что из бытовых привычек партнёра вас раздражает?',
        hint: 'Выберите все подходящие',
        type: 'multi',
        options: [
            { value: 'mess', label: 'Оставляет беспорядок' },
            { value: 'dishes', label: 'Не моет посуду вовремя' },
            { value: 'bathroom', label: 'Не следит за чистотой в ванной' },
            { value: 'noise', label: 'Шумит, когда я отдыхаю' },
            { value: 'food', label: 'Привычки в еде' },
            { value: 'schedule', label: 'Разный режим дня' },
            { value: 'phone', label: 'Постоянно в телефоне' },
            { value: 'nothing', label: 'Ничего не раздражает' }
        ]
    },
    {
        id: 'house_cleanliness',
        category: 'household',
        categoryName: 'Быт и дом',
        categoryIcon: '\uD83C\uDFE0',
        text: 'Насколько важен для вас порядок в доме?',
        hint: '',
        type: 'scale',
        scaleLabels: ['Совсем не важен', 'Не очень', 'Умеренно', 'Важен', 'Очень важен']
    },
    {
        id: 'house_cooking',
        category: 'household',
        categoryName: 'Быт и дом',
        categoryIcon: '\uD83C\uDFE0',
        text: 'Как обстоят дела с готовкой?',
        hint: '',
        type: 'single',
        options: [
            { value: 'i_cook', label: 'Готовлю в основном я' },
            { value: 'partner_cooks', label: 'Готовит в основном партнёр' },
            { value: 'equal', label: 'Готовим поровну' },
            { value: 'together', label: 'Готовим вместе' },
            { value: 'order', label: 'Чаще заказываем еду' }
        ]
    },

    // ─────────────────────────────────────────────────
    // Category 4: Интимная жизнь (Intimacy)
    // ─────────────────────────────────────────────────
    {
        id: 'intim_satisfaction',
        category: 'intimacy',
        categoryName: 'Близость и интимность',
        categoryIcon: '\uD83D\uDD25',
        text: 'Насколько вы удовлетворены вашей интимной жизнью?',
        hint: 'Ваши ответы полностью конфиденциальны',
        type: 'scale',
        scaleLabels: ['Неудовлетворён/-а', 'Скорее нет', 'Средне', 'Скорее да', 'Полностью']
    },
    {
        id: 'intim_frequency',
        category: 'intimacy',
        categoryName: 'Близость и интимность',
        categoryIcon: '\uD83D\uDD25',
        text: 'Совпадают ли ваши потребности в близости по частоте?',
        hint: '',
        type: 'single',
        options: [
            { value: 'match', label: 'Да, мы на одной волне' },
            { value: 'i_more', label: 'Мне хотелось бы чаще' },
            { value: 'partner_more', label: 'Партнёру хотелось бы чаще' },
            { value: 'varies', label: 'По-разному, зависит от периода' },
            { value: 'not_discuss', label: 'Мы это не обсуждаем' }
        ]
    },
    {
        id: 'intim_affection',
        category: 'intimacy',
        categoryName: 'Близость и интимность',
        categoryIcon: '\uD83D\uDD25',
        text: 'Достаточно ли в ваших отношениях несексуальных прикосновений?',
        hint: 'Объятия, поцелуи, держаться за руки',
        type: 'single',
        options: [
            { value: 'plenty', label: 'Да, мы очень тактильные' },
            { value: 'enough', label: 'В целом да' },
            { value: 'want_more', label: 'Хотелось бы больше' },
            { value: 'little', label: 'Маловато' },
            { value: 'almost_none', label: 'Почти нет, и это проблема' }
        ]
    },
    {
        id: 'intim_hard_to_say',
        category: 'intimacy',
        categoryName: 'Близость и интимность',
        categoryIcon: '\uD83D\uDD25',
        text: 'Есть ли что-то в интимной жизни, что вам сложно обсудить с партнёром?',
        hint: 'AI сохранит это в тайне и деликатно поможет',
        type: 'textarea',
        placeholder: 'Опишите свободно или пропустите этот вопрос...'
    },

    // ─────────────────────────────────────────────────
    // Category 5: Финансы (Finances)
    // ─────────────────────────────────────────────────
    {
        id: 'fin_management',
        category: 'finances',
        categoryName: 'Финансы',
        categoryIcon: '\uD83D\uDCB0',
        text: 'Как вы управляете финансами?',
        hint: '',
        type: 'single',
        options: [
            { value: 'joint', label: 'Общий бюджет' },
            { value: 'split', label: 'Раздельные бюджеты' },
            { value: 'mix', label: 'Частично общий, частично личный' },
            { value: 'one_manages', label: 'Один управляет за обоих' },
            { value: 'no_system', label: 'Нет чёткой системы' }
        ]
    },
    {
        id: 'fin_tension',
        category: 'finances',
        categoryName: 'Финансы',
        categoryIcon: '\uD83D\uDCB0',
        text: 'Бывают ли у вас разногласия из-за денег?',
        hint: '',
        type: 'single',
        options: [
            { value: 'never', label: 'Практически никогда' },
            { value: 'rarely', label: 'Редко' },
            { value: 'sometimes', label: 'Время от времени' },
            { value: 'often', label: 'Довольно часто' },
            { value: 'constant', label: 'Это постоянная тема' }
        ]
    },
    {
        id: 'fin_worry',
        category: 'finances',
        categoryName: 'Финансы',
        categoryIcon: '\uD83D\uDCB0',
        text: 'Что вас больше всего беспокоит в финансовых вопросах с партнёром?',
        hint: 'Выберите все подходящие',
        type: 'multi',
        options: [
            { value: 'spending', label: 'Импульсивные траты партнёра' },
            { value: 'saving', label: 'Недостаточные накопления' },
            { value: 'inequality', label: 'Неравный вклад в общие расходы' },
            { value: 'transparency', label: 'Недостаток открытости в финансах' },
            { value: 'goals', label: 'Разные финансовые цели' },
            { value: 'nothing', label: 'Ничего не беспокоит' }
        ]
    },

    // ─────────────────────────────────────────────────
    // Category 6: Свободное время (Quality Time)
    // ─────────────────────────────────────────────────
    {
        id: 'time_together',
        category: 'quality_time',
        categoryName: 'Время вместе',
        categoryIcon: '\u23F0',
        text: 'Довольны ли вы количеством времени, которое проводите вместе?',
        hint: '',
        type: 'scale',
        scaleLabels: ['Совсем мало', 'Маловато', 'Нормально', 'Достаточно', 'Идеально']
    },
    {
        id: 'time_activities',
        category: 'quality_time',
        categoryName: 'Время вместе',
        categoryIcon: '\u23F0',
        text: 'Чего вам не хватает в совместном времени?',
        hint: 'Выберите все подходящие',
        type: 'multi',
        options: [
            { value: 'dates', label: 'Свиданий и романтических вечеров' },
            { value: 'walks', label: 'Совместных прогулок' },
            { value: 'hobby', label: 'Общего хобби' },
            { value: 'travel', label: 'Путешествий' },
            { value: 'talk', label: 'Просто разговоров' },
            { value: 'fun', label: 'Совместного веселья' },
            { value: 'enough', label: 'Всего хватает' }
        ]
    },
    {
        id: 'time_personal',
        category: 'quality_time',
        categoryName: 'Время вместе',
        categoryIcon: '\u23F0',
        text: 'Достаточно ли у вас личного пространства и времени на себя?',
        hint: '',
        type: 'single',
        options: [
            { value: 'plenty', label: 'Да, более чем достаточно' },
            { value: 'enough', label: 'В целом хватает' },
            { value: 'want_more', label: 'Хотелось бы больше' },
            { value: 'almost_none', label: 'Почти нет' },
            { value: 'partner_issue', label: 'Партнёр обижается, когда я провожу время без него/неё' }
        ]
    },

    // ─────────────────────────────────────────────────
    // Category 7: Семья и окружение (Family & Social)
    // ─────────────────────────────────────────────────
    {
        id: 'fam_inlaws',
        category: 'family',
        categoryName: 'Семья и окружение',
        categoryIcon: '\uD83D\uDC68\u200D\uD83D\uDC69\u200D\uD83D\uDC67',
        text: 'Как складываются ваши отношения с семьёй партнёра?',
        hint: '',
        type: 'scale',
        scaleLabels: ['Плохо', 'Натянуто', 'Нормально', 'Хорошо', 'Отлично']
    },
    {
        id: 'fam_interference',
        category: 'family',
        categoryName: 'Семья и окружение',
        categoryIcon: '\uD83D\uDC68\u200D\uD83D\uDC69\u200D\uD83D\uDC67',
        text: 'Вмешиваются ли родственники в ваши отношения?',
        hint: '',
        type: 'single',
        options: [
            { value: 'never', label: 'Нет, они уважают границы' },
            { value: 'sometimes', label: 'Иногда дают непрошенные советы' },
            { value: 'my_family', label: 'Моя семья бывает навязчива' },
            { value: 'partner_family', label: 'Семья партнёра вмешивается' },
            { value: 'both', label: 'Обе семьи активно участвуют' }
        ]
    },
    {
        id: 'fam_friends',
        category: 'family',
        categoryName: 'Семья и окружение',
        categoryIcon: '\uD83D\uDC68\u200D\uD83D\uDC69\u200D\uD83D\uDC67',
        text: 'Есть ли разногласия из-за друзей?',
        hint: '',
        type: 'single',
        options: [
            { value: 'no', label: 'Нет, мы уважаем друзей друг друга' },
            { value: 'time', label: 'Партнёр проводит слишком много времени с друзьями' },
            { value: 'dislike', label: 'Мне не нравятся некоторые друзья партнёра' },
            { value: 'jealousy', label: 'Есть ревность к друзьям' },
            { value: 'excluded', label: 'Чувствую себя исключённым/-ой из компании партнёра' }
        ]
    },

    // ─────────────────────────────────────────────────
    // Category 8: Ценности и цели (Values & Goals)
    // ─────────────────────────────────────────────────
    {
        id: 'val_future',
        category: 'values',
        categoryName: 'Ценности и цели',
        categoryIcon: '\uD83C\uDF1F',
        text: 'Совпадают ли ваши представления о будущем?',
        hint: 'Дети, карьера, место жизни и т.д.',
        type: 'scale',
        scaleLabels: ['Совсем нет', 'Скорее нет', 'Частично', 'В основном', 'Полностью']
    },
    {
        id: 'val_disagreements',
        category: 'values',
        categoryName: 'Ценности и цели',
        categoryIcon: '\uD83C\uDF1F',
        text: 'В каких жизненных вопросах вы не совпадаете?',
        hint: 'Выберите все подходящие',
        type: 'multi',
        options: [
            { value: 'children', label: 'Вопрос детей' },
            { value: 'career', label: 'Карьерные приоритеты' },
            { value: 'location', label: 'Где жить' },
            { value: 'lifestyle', label: 'Образ жизни' },
            { value: 'religion', label: 'Религия / духовность' },
            { value: 'politics', label: 'Политические взгляды' },
            { value: 'none', label: 'Мы во всём совпадаем' }
        ]
    },

    // ─────────────────────────────────────────────────
    // Category 9: Привычки и триггеры (Habits & Triggers)
    // ─────────────────────────────────────────────────
    {
        id: 'hab_partner_annoys',
        category: 'habits',
        categoryName: 'Привычки и триггеры',
        categoryIcon: '\u26A1',
        text: 'Что в поведении партнёра вызывает у вас наибольшее раздражение?',
        hint: 'Будьте честны — AI поможет решить это деликатно',
        type: 'textarea',
        placeholder: 'Например: забывает мыть за собой посуду, поздно ложится...'
    },
    {
        id: 'hab_my_flaw',
        category: 'habits',
        categoryName: 'Привычки и триггеры',
        categoryIcon: '\u26A1',
        text: 'А что в вашем поведении может раздражать партнёра?',
        hint: 'Самосознание — ключ к гармонии',
        type: 'textarea',
        placeholder: 'Например: много работаю, забываю о планах...'
    },
    {
        id: 'hab_triggers',
        category: 'habits',
        categoryName: 'Привычки и триггеры',
        categoryIcon: '\u26A1',
        text: 'Что гарантированно выводит вас из себя?',
        hint: 'Это поможет AI избегать болезненных тем',
        type: 'multi',
        options: [
            { value: 'ignore', label: 'Когда меня игнорируют' },
            { value: 'late', label: 'Когда партнёр опаздывает' },
            { value: 'lie', label: 'Ложь, даже мелкая' },
            { value: 'tone', label: 'Грубый тон' },
            { value: 'passive', label: 'Пассивная агрессия' },
            { value: 'compare', label: 'Сравнение с другими' },
            { value: 'dismiss', label: 'Обесценивание моих чувств' },
            { value: 'control', label: 'Попытки контролировать' }
        ]
    },

    // ─────────────────────────────────────────────────
    // Category 10: Благодарность и позитив (Appreciation)
    // ─────────────────────────────────────────────────
    {
        id: 'pos_love',
        category: 'appreciation',
        categoryName: 'Что нравится в партнёре',
        categoryIcon: '\uD83D\uDC96',
        text: 'За что вы больше всего цените партнёра?',
        hint: 'Выберите до 3-х ответов',
        type: 'multi',
        maxSelect: 3,
        options: [
            { value: 'humor', label: 'Чувство юмора' },
            { value: 'care', label: 'Заботливость' },
            { value: 'smart', label: 'Ум и интересные разговоры' },
            { value: 'reliable', label: 'Надёжность' },
            { value: 'passion', label: 'Страстность' },
            { value: 'kind', label: 'Доброта' },
            { value: 'ambitious', label: 'Амбициозность' },
            { value: 'calm', label: 'Спокойствие и мудрость' },
            { value: 'creative', label: 'Креативность' }
        ]
    },
    {
        id: 'pos_best_memory',
        category: 'appreciation',
        categoryName: 'Что нравится в партнёре',
        categoryIcon: '\uD83D\uDC96',
        text: 'Опишите ваш лучший совместный момент',
        hint: 'AI будет напоминать о хороших моментах в трудные времена',
        type: 'textarea',
        placeholder: 'Расскажите о моменте, когда вы были особенно счастливы вместе...'
    },

    // ─────────────────────────────────────────────────
    // Category 11: Табу-зоны и границы (Boundaries)
    // ─────────────────────────────────────────────────
    {
        id: 'bound_no_joke',
        category: 'boundaries',
        categoryName: 'Границы и табу',
        categoryIcon: '\uD83D\uDEE1\uFE0F',
        text: 'Есть ли темы, на которые нельзя шутить?',
        hint: 'AI будет уважать ваши границы',
        type: 'multi',
        options: [
            { value: 'appearance', label: 'Внешность и вес' },
            { value: 'family', label: 'Моя семья' },
            { value: 'past', label: 'Моё прошлое' },
            { value: 'career', label: 'Моя работа/карьера' },
            { value: 'insecurities', label: 'Мои комплексы' },
            { value: 'exes', label: 'Бывшие партнёры' },
            { value: 'nothing', label: 'Шучу обо всём' }
        ]
    },
    {
        id: 'bound_dealbreaker',
        category: 'boundaries',
        categoryName: 'Границы и табу',
        categoryIcon: '\uD83D\uDEE1\uFE0F',
        text: 'Что для вас абсолютно неприемлемо в отношениях?',
        hint: '',
        type: 'multi',
        options: [
            { value: 'cheating', label: 'Измена' },
            { value: 'lying', label: 'Системная ложь' },
            { value: 'disrespect', label: 'Неуважение' },
            { value: 'violence', label: 'Любое насилие' },
            { value: 'addiction', label: 'Зависимости' },
            { value: 'neglect', label: 'Полное безразличие' },
            { value: 'control', label: 'Тотальный контроль' }
        ]
    },

    // ─────────────────────────────────────────────────
    // Category 12: Ожидания и желания (Expectations)
    // ─────────────────────────────────────────────────
    {
        id: 'exp_change',
        category: 'expectations',
        categoryName: 'Ожидания',
        categoryIcon: '\u2728',
        text: 'Если бы вы могли изменить одну вещь в партнёре, что бы это было?',
        hint: 'AI аккуратно поработает над этим через челленджи',
        type: 'textarea',
        placeholder: 'Напишите своими словами...'
    },
    {
        id: 'exp_wish',
        category: 'expectations',
        categoryName: 'Ожидания',
        categoryIcon: '\u2728',
        text: 'Чего бы вы хотели больше в ваших отношениях?',
        hint: '',
        type: 'multi',
        options: [
            { value: 'attention', label: 'Больше внимания' },
            { value: 'romance', label: 'Больше романтики' },
            { value: 'independence', label: 'Больше свободы' },
            { value: 'communication', label: 'Более открытого общения' },
            { value: 'intimacy', label: 'Больше близости' },
            { value: 'adventures', label: 'Больше приключений' },
            { value: 'stability', label: 'Больше стабильности' },
            { value: 'humor', label: 'Больше лёгкости и веселья' }
        ]
    },
    {
        id: 'exp_koople_goal',
        category: 'expectations',
        categoryName: 'Ожидания',
        categoryIcon: '\u2728',
        text: 'Что вы хотите получить от Koople?',
        hint: 'Это поможет нам настроить AI-медиатор под ваши цели',
        type: 'multi',
        options: [
            { value: 'communicate', label: 'Научиться лучше общаться' },
            { value: 'conflicts', label: 'Меньше конфликтов' },
            { value: 'understand', label: 'Лучше понимать друг друга' },
            { value: 'closeness', label: 'Стать ближе' },
            { value: 'habits', label: 'Улучшить бытовые привычки' },
            { value: 'spark', label: 'Вернуть искру' },
            { value: 'prevent', label: 'Превентивно решать проблемы' }
        ]
    }
];

// ============================================================
// Demo Data for MVP
// ============================================================

const DEMO_CHALLENGES = [
    {
        id: 'ch1',
        title: 'Неделя комплиментов',
        description: 'Каждый день говорите партнёру один искренний комплимент, который не связан с внешностью',
        category: 'emotional',
        icon: '\uD83D\uDCAC',
        duration: '7 дней',
        difficulty: 'easy',
        progress: 5,
        total: 7,
        status: 'active',
        assignedTo: 'both'
    },
    {
        id: 'ch2',
        title: 'Чистая ванная',
        description: 'Следите за чистотой в ванной после каждого использования: уберите волосы, протрите раковину',
        category: 'household',
        icon: '\u2728',
        duration: '7 дней',
        difficulty: 'easy',
        progress: 3,
        total: 7,
        status: 'active',
        assignedTo: 'partner'
    },
    {
        id: 'ch3',
        title: 'Вечер без телефонов',
        description: 'Проведите вечер вдвоём без телефонов. Поговорите, поиграйте или просто побудьте вместе',
        category: 'quality_time',
        icon: '\uD83D\uDCF5',
        duration: '1 вечер',
        difficulty: 'medium',
        progress: 0,
        total: 1,
        status: 'active',
        assignedTo: 'both'
    },
    {
        id: 'ch4',
        title: 'Благодарность перед сном',
        description: 'Перед сном расскажите партнёру 3 вещи, за которые вы благодарны ему/ей сегодня',
        category: 'appreciation',
        icon: '\uD83C\uDF19',
        duration: '5 дней',
        difficulty: 'easy',
        progress: 5,
        total: 5,
        status: 'completed',
        assignedTo: 'both'
    },
    {
        id: 'ch5',
        title: 'Сюрприз-свидание',
        description: 'Организуйте неожиданное свидание для партнёра. Это может быть что угодно — от ужина дома до прогулки в новом месте',
        category: 'quality_time',
        icon: '\uD83C\uDF39',
        duration: 'Однократно',
        difficulty: 'medium',
        progress: 0,
        total: 1,
        status: 'suggested',
        assignedTo: 'user'
    },
    {
        id: 'ch6',
        title: 'Финансовый вечер',
        description: 'Устройте спокойный разговор о финансах: обсудите траты за месяц, планы и мечты',
        category: 'finances',
        icon: '\uD83D\uDCB0',
        duration: 'Однократно',
        difficulty: 'hard',
        progress: 0,
        total: 1,
        status: 'suggested',
        assignedTo: 'both'
    }
];

const DEMO_ACTIVITY = [];

const DEMO_WISHES = [
    {
        id: 'w1',
        text: 'Хочу, чтобы партнёр чаще спрашивал как прошёл мой день',
        category: 'emotional',
        status: 'active',
        createdAt: '3 дня назад'
    },
    {
        id: 'w2',
        text: 'Хочу больше совместных прогулок по вечерам',
        category: 'quality_time',
        status: 'active',
        createdAt: '1 неделю назад'
    },
    {
        id: 'w3',
        text: 'Хочу чтобы партнёр убирал за собой на кухне сразу после готовки',
        category: 'household',
        status: 'in_progress',
        createdAt: '2 недели назад'
    }
];

const COMPATIBILITY_DATA = [
    { name: 'Эмоциональная связь', value: 88, color: '#4ade80' },
    { name: 'Коммуникация', value: 75, color: '#facc15' },
    { name: 'Быт', value: 62, color: '#fb923c' },
    { name: 'Близость', value: 85, color: '#4ade80' },
    { name: 'Финансы', value: 70, color: '#facc15' },
    { name: 'Время вместе', value: 78, color: '#4ade80' },
    { name: 'Семья', value: 82, color: '#4ade80' },
    { name: 'Ценности', value: 90, color: '#4ade80' },
    { name: 'Привычки', value: 58, color: '#fb923c' }
];

const AI_RECOMMENDATIONS = [
    {
        title: 'Обсудите бытовые обязанности',
        text: 'Ваши ответы показывают разногласия в распределении домашних дел. Попробуйте составить совместный список обязанностей.',
        priority: 'high'
    },
    {
        title: 'Больше качественного времени',
        text: 'Вы оба хотите проводить больше времени вместе. Выделите хотя бы один вечер в неделю только для двоих.',
        priority: 'medium'
    },
    {
        title: 'Практикуйте активное слушание',
        text: 'Когда партнёр рассказывает о своём дне, попробуйте просто слушать без советов и оценок.',
        priority: 'medium'
    }
];

const CHAT_WELCOME_MESSAGES = [
    {
        role: 'assistant',
        text: 'Привет! Я ваш AI-медиатор Koople. \uD83D\uDC4B\n\nЯ здесь, чтобы помочь вам и вашему партнёру лучше понимать друг друга.\n\nВы можете рассказать мне о том, что вас беспокоит, попросить совет или просто поговорить о ваших отношениях. Всё, что вы скажете, останется конфиденциальным.',
        time: 'Сейчас'
    }
];
