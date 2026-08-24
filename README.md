# Массовое создание общих ящиков Яндекс 360

Скрипт `create_shared_mailboxes.py` создаёт несколько общих ящиков из одной таблицы и назначает сотрудникам права доступа.

Используемые методы API:

- [Создать общий ящик](https://yandex.ru/dev/api360/doc/ru/ref/MailboxService/MailboxService_CreateShared);
- [Изменить права доступа к ящику](https://yandex.ru/dev/api360/doc/ru/ref/MailboxService/MailboxService_Set);
- [Проверить статус задачи](https://yandex.ru/dev/api360/doc/ru/ref/MailboxService/MailboxService_TaskStatus);
- [Получить список сотрудников](https://yandex.ru/dev/api360/doc/ru/ref/UserService/UserService_List) — используется, если вместо UID указан email или логин.

## Что изменилось

- Параметры каждого общего ящика и права доступа находятся в одном XLSX/CSV.
- В колонке `UID` можно указать UID, рабочий email или логин сотрудника.
- Один общий ящик создаётся один раз, даже если он повторяется в нескольких строках для разных сотрудников.
- Если общий ящик уже существует в организации, скрипт использует его `resourceId` и назначает указанные в таблице права без повторного создания.
- Роли можно писать полностью или коротко: `owner`, `imap`, `sender`, `half_sender`.
- Поддерживаются разделители ролей: запятая, точка с запятой и пробел.
- Пустое описание разрешено и не передаётся в запросе.
- До создания ящиков проверяется весь входной файл.
- Есть безопасная проверка `--dry-run` и итоговый CSV-отчёт.
- Скрипт проверяет фактическое завершение асинхронных задач назначения прав.

## Формат таблицы

Для XLSX используется лист `main_list`.

Обязательные заголовки:

```text
UID | email | name | description | roles
```

- `UID` — UID, основной email или логин сотрудника, которому предоставляются права;
- `email` — адрес или локальная часть адреса создаваемого общего ящика;
- `name` — отображаемое имя общего ящика;
- `description` — необязательное описание;
- `roles` — одна или несколько ролей.

Одна строка — одно назначение прав. Чтобы предоставить доступ к одному ящику нескольким сотрудникам, повторите `email`, `name` и `description` в нескольких строках. Ящик будет создан только один раз.

Не перечисляйте несколько UID в одной ячейке. Для каждого сотрудника используется отдельная строка.

Пример:

```text
UID                  email                name                 description          roles
ivanov@example.ru    support@example.ru   Служба поддержки     Обращения клиентов   owner
1130000000000001     support@example.ru   Служба поддержки     Обращения клиентов   imap; sender
petrov               press@example.ru     Пресс-служба                              half_sender
```

Если один и тот же общий ящик повторяется в таблице, значения `name` и `description` во всех его строках должны совпадать.

## Роли

Короткое значение | Значение API | Назначение
--- | --- | ---
`imap` | `shared_mailbox_imap_admin` | Управление ящиком по IMAP
`sender` | `shared_mailbox_sender` | Отправка писем
`half_sender` | `shared_mailbox_half_sender` | Отправка «от имени» по SMTP
`owner` или `all` | `shared_mailbox_owner` | Полные права

`shared_mailbox_owner` уже включает все остальные роли. Если указать `owner` вместе с другими ролями, скрипт отправит только `shared_mailbox_owner`.

Скрипт не добавляет права автоматически: назначается только то, что указано в таблице.

## Установка

### Windows 10/11 — PowerShell

```powershell
git clone https://github.com/MilekhinAV/Y360_shared_mailbox.git
cd Y360_shared_mailbox
py -m venv .venv
.venv\Scripts\Activate.ps1
py -m pip install -r requirements.txt
```

Если PowerShell запрещает активацию виртуального окружения, выполните один раз для текущего пользователя:

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

### Windows 10/11 — Git Bash

```bash
git clone https://github.com/MilekhinAV/Y360_shared_mailbox.git
cd Y360_shared_mailbox
python -m venv .venv
source .venv/Scripts/activate
python -m pip install -r requirements.txt
```

### macOS / Linux

```bash
git clone https://github.com/MilekhinAV/Y360_shared_mailbox.git
cd Y360_shared_mailbox
python3 -m venv .venv
source .venv/bin/activate
python3 -m pip install -r requirements.txt
```

## Настройка

Создайте в папке репозитория файл `.env`:

```ini
OAUTH_TOKEN=y0__ваш_OAuth_токен
ORG_ID=1234567
NOTIFY=none
```

`NOTIFY` необязателен:

- `none` — никому, значение по умолчанию;
- `delegates` — только сотруднику;
- `all` — сотруднику и владельцу ящика.

OAuth-приложению нужны права на управление доступом к почтовым ящикам. Если в `UID` используются email или логины, также требуется право на чтение сотрудников организации.

## Проверка без создания

Windows PowerShell / Git Bash:

```text
python create_shared_mailboxes.py sharedMailBox_list.xlsx --dry-run
```

macOS / Linux:

```text
python3 create_shared_mailboxes.py sharedMailBox_list.xlsx --dry-run
```

`--dry-run` проверяет структуру таблицы, значения, роли и поиск UID. Запросы на создание ящиков и назначение прав не отправляются.

Проверка также показывает, какие ящики будут созданы, а какие уже существуют и будут использованы для назначения прав.

## Создание общих ящиков

Windows PowerShell / Git Bash:

```text
python create_shared_mailboxes.py sharedMailBox_list.xlsx
```

macOS / Linux:

```text
python3 create_shared_mailboxes.py sharedMailBox_list.xlsx
```

После выполнения создаётся файл `shared_mailboxes_result.csv` с `resourceId`, UID, ролями и результатом каждого назначения.

## Дополнительные параметры

```text
python create_shared_mailboxes.py sharedMailBox_list.xlsx --sheet main_list --notify none --task-timeout 60 --report result.csv
```

- `--sheet` — имя листа XLSX;
- `--notify` — переопределить `NOTIFY` из `.env`;
- `--task-timeout` — сколько секунд ждать завершения одной задачи назначения прав;
- `--report` — путь к итоговому CSV.

## Повторный запуск и существующие ящики

- Если ящик уже существует, скрипт получает его `resourceId` и переходит к назначениям из таблицы.
- Сотрудники существующего ящика, которых нет в таблице, не удаляются и не изменяются.
- Для сотрудника, который присутствует в таблице, метод API устанавливает указанный в строке набор ролей. Поэтому повторный запуск безопасно приводит его права к значениям из таблицы.
- `name` и `description` существующего ящика автоматически не изменяются. При расхождении с таблицей выводится предупреждение.
- Если в таблице указана только локальная часть адреса и в организации найдено несколько совпадений в разных доменах, выполнение прекращается. В таком случае укажите полный email.
- Перед реальным запуском всегда выполняйте `--dry-run`.
- Если создание одного ящика завершилось ошибкой, скрипт продолжит обработку остальных и зафиксирует ошибку в отчёте.
- Статус `timeout` означает, что API принял задачу, но она не завершилась за заданное время. `taskId` сохраняется в отчёте.
