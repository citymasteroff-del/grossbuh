/**
 * ГРОССБУХ — синхронизация с Google Таблицей.
 *
 * УСТАНОВКА (один раз):
 * 1. Открой свою Google Таблицу (или создай новую).
 * 2. Меню Расширения → Apps Script.
 * 3. Удали весь код-заглушку и вставь этот файл целиком.
 * 4. Замени строку ниже "МОЙ_СЕКРЕТНЫЙ_ПАРОЛЬ" на свой пароль.
 * 5. Нажми Развернуть → Новое развёртывание → тип "Веб-приложение".
 *    - Выполнять как: Я (твой аккаунт)
 *    - У кого есть доступ: Все (Anyone)
 * 6. Скопируй выданный URL (заканчивается на /exec) — его нужно вставить
 *    в приложении Гроссбух на вкладке «Синхронизация».
 * 7. При каждом изменении этого кода нужно делать новое развёртывание
 *    (Развернуть → Управление развёртываниями → редактировать → новая версия).
 */

const PASSWORD = "МОЙ_СЕКРЕТНЫЙ_ПАРОЛЬ"; // <-- замени на свой пароль перед деплоем
const SHEET_NAME = "Transactions";

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: "Ledger sync endpoint is running" }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);

    if (body.password !== PASSWORD) {
      return jsonOut({ error: "Неверный пароль" });
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (body.action === "push") {
      let sheet = ss.getSheetByName(SHEET_NAME);
      if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
      sheet.clearContents();
      sheet.appendRow(["id", "date", "description", "amount", "type", "category", "source"]);
      const tx = body.transactions || [];
      tx.forEach(t => {
        sheet.appendRow([t.id, t.date, t.description, t.amount, t.type, t.category, t.source || ""]);
      });
      return jsonOut({ ok: true, count: tx.length });
    }

    if (body.action === "pull") {
      const sheet = ss.getSheetByName(SHEET_NAME);
      if (!sheet) return jsonOut({ transactions: [] });
      const data = sheet.getDataRange().getValues();
      const rows = data.slice(1)
        .filter(r => r[0] !== "" && r[0] !== null)
        .map(r => ({
          id: String(r[0]),
          date: formatDateCell(r[1]),
          description: String(r[2]),
          amount: Number(r[3]),
          type: String(r[4]),
          category: String(r[5]),
          source: String(r[6] || "")
        }));
      return jsonOut({ transactions: rows });
    }

    return jsonOut({ error: "Неизвестное действие: " + body.action });
  } catch (err) {
    return jsonOut({ error: String(err) });
  }
}

function formatDateCell(v) {
  if (v instanceof Date) {
    const y = v.getFullYear();
    const m = String(v.getMonth() + 1).padStart(2, "0");
    const d = String(v.getDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  }
  return String(v);
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
