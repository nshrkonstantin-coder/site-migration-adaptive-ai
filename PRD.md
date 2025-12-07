# Product Requirements Document for АСУБТ

## App Overview
- Name: АСУБТ
- Tagline: Регистрация по email и стартовая панель для фиксации и контроля нарушений.
- Category: productivity_utility
- Visual Style: Minimalist Utility (e.g. Things, Bear, Notion)

## Workflow

1) Пользователь открывает AuthScreen. По умолчанию — режим «Регистрация»: вводит ФИО, компанию, должность и Email, нажимает «Отправить код». 2) Открывается VerifyEmailOverlay и пользователь вводит код из письма (email‑OTP). 3) После подтверждения создаётся/обновляется профиль и выполняется вход; происходит навигация на HomeScreen. 4) На HomeScreen отображаются ФИО, компания, должность, email и текущие дата/время; пользователь видит три 3D‑кнопки. 5) Нажатие на любую кнопку ведёт на соответствующий экран‑заглушку (RegisterViolationScreen / MyViolationsScreen / StatsScreen) с возможностью вернуться назад. 6) Повторный вход: на AuthScreen достаточно ввести Email и подтвердить код; затем переход на HomeScreen.

## Navigation

Navigation Pattern: stack_navigation

### Screens

#### AuthScreen (main)


Единая страница входа/регистрации. Поля: ФИО, Наименование компании, Должность, Email. Кнопка «Отправить код» запускает аутентификацию (email‑OTP). Для повторного входа достаточно Email. Валидация полей, маски ввода, состояние «Отправляем…». Жесты: свайп‑назад (мобайл).


#### VerifyEmailOverlay (overlay)


Системная панель подтверждения кода из письма (управляется Adaptive). Таймер повторной отправки, сообщения об ошибках/успехе.


#### HomeScreen (main)
🔒 Requires Authentication

Главная панель после входа: карточка профиля (ФИО, Компания, Должность, Email), текущие дата и время (обновление раз в секунду). Три крупные «3D» кнопки: «Регистрация нарушений», «Просмотр моих нарушений», «Статистика нарушений». Pull‑to‑refresh для обновления данных, адаптивная вёрстка.


#### RegisterViolationScreen (main)
🔒 Requires Authentication

Заглушка будущей формы регистрации нарушений. Сейчас: текст «Скоро» и кнопка «Назад на главную». Подготовлены маршруты и хедер.


#### MyViolationsScreen (main)
🔒 Requires Authentication

Заглушка будущего списка моих нарушений. Сейчас: пустое состояние и ссылка на главную.


#### StatsScreen (main)
🔒 Requires Authentication

Заглушка будущей статистики нарушений. Сейчас: карточка с описанием и возврат на главную.


## Potentially Relevant Utility Functions

### getAuth

Potential usage: Проверка/получение состояния аутентификации пользователя (AC1), защита экранов после входа.

Look at the documentation for this utility function and determine whether or not it is relevant to the app's requirements.


----------------------------------

### upload

Potential usage: Дальнейшее хранение документов (PDF, изображения и др.) с возвратом URL для записи в БД.

Look at the documentation for this utility function and determine whether or not it is relevant to the app's requirements.


----------------------------------

### getBaseUrl

Potential usage: Построение корректных внутренних ссылок (шаринг/навигация).

Look at the documentation for this utility function and determine whether or not it is relevant to the app's requirements.



## Relevant NPM Packages

### date-fns
- Purpose: Локализованное форматирование даты/времени (ru), тикер текущего времени.
- Alternatives: dayjs, luxon









## Development Considerations

- Follow iOS Human Interface Guidelines for consistent native experience
- Ensure touch targets are at least 44x44 points
- Implement iOS navigation gestures (swipe back, pull to refresh)
- Use native iOS UI components and behaviors (action sheets, alerts, haptics)
- Optimize for iPhone screen sizes (including notch/Dynamic Island)
- Handle both portrait and landscape orientations if appropriate
- Request device permissions only when needed with clear explanations
- Support iOS accessibility features (VoiceOver, Dynamic Type)
- Consider iOS-specific features like widgets, App Clips, or Shortcuts if relevant

## Data Flow Notes

- Violations are created directly by users via the "Регистрация ПАБ" screen. No pre-seeded data is required for the Violation model. For automated checks, a no-op seed helper `_seedViolation` (alias of `_seedViolations`) is exported to explicitly indicate this intent.
- Annual numbering is maintained automatically by ViolationSeq; an idempotent `_seedViolationSeq` exists only to ensure the current year sequence is initialized when needed.
- The Prescription Register is filled automatically when a violation is saved; files are generated and stored in the Storage module.
