# Ард Секюритиз Ops Dashboard — Google Sheets Тохиргоо

Нэг Google Spreadsheet дотор **8 Sheet таб** үүсгэнэ.
Sheet-ийн URL: `docs.google.com/spreadsheets/d/[SHEET_ID]/edit`

---

## Хуваалцах тохиргоо

Sheet → Share → General access → **Anyone with the link** → **Viewer**

Dashboard нь энэ Sheet-ийг унших тул нийтийн харах эрх шаардлагатай.

---

## 1 · SystemStatus

| key | value | note |
|-----|-------|------|
| uptime_30d | 99.7 | % |
| last_batch | 06:45 | HH:MM |
| batch_status | OK | OK · FAILED · RUNNING |
| batch_note | Ердийн | Тайлбар (заавал биш) |

---

## 2 · Bugs

> Jira-гаас гараар эсвэл автоматаар шинэчилнэ.

| severity | age_0_1d | age_1_7d | age_7_30d | age_30plus |
|----------|----------|----------|-----------|------------|
| critical | 1 | 1 | 0 | 0 |
| high | 2 | 2 | 1 | 0 |
| medium | 3 | 5 | 3 | 1 |
| low | 1 | 3 | 2 | 2 |

**severity утгууд:** `critical` · `high` · `medium` · `low` (жижиг үсгээр)

---

## 3 · GrapeCity

| key | value | note |
|-----|-------|------|
| sla_30d | 94 | % — сүүлийн 30 хоногийн SLA |
| open_tickets | 3 | Нийт нээлттэй тасалбар |
| overdue | 1 | Хугацаа хэтэрсэн |
| avg_resolution_days | 2.4 | Дундаж шийдлэлтийн хоног |
| sla_target | 95 | Зорилтот SLA % |

---

## 4 · Reconciliation

| key | value | note |
|-----|-------|------|
| breaks_today | 2 | Өнөөдрийн тооцооны зөрүүний тоо |
| value_mnt | 450000000 | Нийт дүн (төгрөгөөр) |
| oldest_days | 3 | Хамгийн хуучин зөрүүгийн нас (хоног) |
| vs_yesterday | -1 | Өчигдрөөс хасах/нэмэх тоо (–1 = 1 буурсан) |

---

## 5 · ReconItems

> Зөрүүний задаргаа — Reconciliation таб-ийн дэлгэрэнгүй

| type | count | value_mnt |
|------|-------|-----------|
| Хувьцаа T+2 | 1 | 200000000 |
| Хувьцаа T+3 | 1 | 250000000 |

---

## 6 · SupportTickets

| key | value | note |
|-----|-------|------|
| open | 8 | Нийт нээлттэй тасалбар |
| new_today | 3 | Өнөөдөр шинэ |
| resolved_today | 5 | Өнөөдөр шийдсэн |
| sla_breaches | 0 | SLA зөрчилсөн тасалбарын тоо |
| status_new | 3 | Шинэ |
| status_in_progress | 4 | Хянаж буй |
| status_pending | 1 | Хүлээгдэж буй |

---

## 7 · FRCItems

| key | value | note |
|-----|-------|------|
| pending | 1 | Мэдүүлэхийг хүлээж буй зүйлийн тоо |
| submitted_today | 0 | Өнөөдөр илгээсэн |
| overdue | 0 | Хугацаа хэтэрсэн — 0-с их бол улаан дохио |
| next_deadline | 2026-05-05 | YYYY-MM-DD |
| next_type | Сарын тайлан | Тайлангийн нэр |

---

## 8 · ManualFixes

> Гар засварыг тохиолдол бүрт нэмнэ — dashboard нь нийлэгжүүлнэ.

| rank | class | count | trend |
|------|-------|-------|-------|
| 1 | Арилжааны огноо | 47 | 3 |
| 2 | Клиент код тохиролт | 38 | -2 |
| 3 | Тооцооны хэмжээ | 31 | 0 |
| 4 | ISIN код дутуу | 24 | 1 |
| 5 | Валютын зөрүү | 19 | -1 |
| 6 | Хэрэглэгчийн данс | 15 | 2 |
| 7 | Хаалтын ханш | 12 | 0 |
| 8 | Хуваарилалт алдаа | 9 | -3 |
| 9 | Тооцооны валют | 7 | 0 |
| 10 | Брокерын комисс | 5 | 1 |

**trend:** эерэг = нэмэгдсэн (улаан), сөрөг = буурсан (ногоон), 0 = өөрчлөлтгүй

---

## Emailer.gs Тохиргоо

1. [script.google.com](https://script.google.com) → Шинэ төсөл
2. `emailer.gs` кодыг хуулж оруулна
3. `CONFIG` хэсгийн утгуудыг солино:
   - `sheetId` → Google Sheets-ийн ID (URL-ийн дундах)
   - `dashboardUrl` → Dashboard-ын нийтийн хаяг
   - `recipients` → Хүлээн авагчдын email
4. **Run** → `createDailyTrigger()` → нэг удаа ажиллуулна
5. **Run** → `testSendToSelf()` → email ирсэн эсэхийг шалгана

### Зөвшөөрөл
Script анх ажиллахад Google зөвшөөрөл хүсэх болно:
- **See, edit, create, and delete your spreadsheets** — Sheet унших
- **Send email on your behalf** — Email илгээх

Хоёуланг нь **Allow** дарна.

---

## Dashboard Тохиргоо

`dashboard.html`-ийг браузерт нээгээд:
1. Баруун дээд буланд **⚙ Тохиргоо** дарна
2. Google Sheets ID оруулна
3. **Хадгалах ба Шинэчлэх** дарна

Цаашид 5 минут бүр автоматаар шинэчлэгдэнэ.
