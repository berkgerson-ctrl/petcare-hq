/**
 * 🐾 PetCare HQ — Google Apps Script Backend
 * Tek Google Sheet üzerinde çalışan, çoklu-tablo (sheet) destekli
 * genel amaçlı CRUD API'si.
 *
 * KURULUM:
 * 1) Yeni bir Google Sheet oluştur.
 * 2) Uzantılar > Apps Script menüsünden bu kodu yapıştır.
 * 3) SHEET_ID sabitine ihtiyaç yok; script otomatik olarak bağlı
 *    olduğu Sheet'i kullanır (SpreadsheetApp.getActiveSpreadsheet()).
 * 4) "Dağıt > Yeni Dağıtım > Web Uygulaması" seç.
 *    - Yürüten kişi: Ben (Me)
 *    - Erişimi olanlar: Herkes (Anyone)
 * 5) Verilen Web App URL'sini frontend'deki CONFIG.API_URL alanına yapıştır.
 * 6) İlk çalıştırmada setupSheets() fonksiyonunu manuel çalıştırarak
 *    (Apps Script editöründe fonksiyon seçip ▶️ Çalıştır) tüm sekmeleri
 *    ve başlık satırlarını otomatik oluşturabilirsin.
 * 7) HATIRLATMA E-POSTALARI İÇİN: e-posta adresini artık koda yazmana
 *    gerek yok — uygulama içinde "Pet Bilgilerini Düzenle" (veya Yeni
 *    Pet Ekle) ekranındaki "Hatırlatma E-postası" alanından gir ve
 *    kaydet. Ardından Apps Script'te fonksiyon listesinden
 *    setupDailyReminderTrigger() fonksiyonunu seçip BİR KEZ ▶️ Çalıştır.
 *    İlk çalıştırmada Google senden izin isteyecek (e-posta gönderme
 *    ve tetikleyici oluşturma yetkisi) — onayla. Bundan sonra her gün
 *    saat 09:00 civarında otomatik kontrol edip hatırlatma gönderecek.
 *
 * ÖNEMLİ — ÖZ-ONARIMLI BAŞLIKLAR:
 * Aşağıdaki SCHEMAS listesi sadece "bu tabloda hangi alanlar olmalı"yı
 * tanımlar; GERÇEK sütun sırası her zaman Sheet'in fiziksel başlık
 * satırından okunur. Yeni bir alan eklemek istediğinde SCHEMAS'a
 * dilediğin yere ekleyebilirsin — kod, mevcut sekmelerde eksik olan
 * alanları otomatik olarak sütunların EN SONUNA ekler, var olan
 * sütunların sırasını asla bozmaz. Bu sayede sütun kayması (bir
 * hücredeki verinin yanlış alana yazılması) bir daha yaşanmaz.
 */

// ---- Sheet şemaları (bu tabloda olması gereken alanlar) ----
const SCHEMAS = {
  Pets: ['id', 'name', 'species', 'breed', 'chipNumber', 'birthdate', 'photoUrl', 'vetName', 'vetPhone', 'notes', 'createdAt'],
  Vaccines: ['id', 'petId', 'category', 'name', 'date', 'nextDate', 'completed', 'notes', 'createdAt'],
  Medications: ['id', 'petId', 'name', 'dosage', 'frequency', 'startDate', 'endDate', 'completed', 'completedDates', 'notes', 'createdAt'],
  Boarding: ['id', 'petId', 'hotelName', 'checkIn', 'checkOut', 'cost', 'currency', 'careInstructions', 'notes', 'expenseId', 'createdAt'],
  Expenses: ['id', 'petId', 'category', 'amount', 'currency', 'date', 'insuranceClaim', 'refundAmount', 'notes', 'source', 'linkedHealthType', 'linkedHealthId', 'createdAt'],
  NutritionLog: ['id', 'petId', 'logType', 'title', 'detail', 'estimatedDays', 'amount', 'currency', 'expenseId', 'date', 'createdAt'],
  Settings: ['id', 'reminderEmail', 'createdAt']
};

/** İlk kurulum / onarım: eksik sekmeleri ve eksik başlık sütunlarını oluşturur. */
function setupSheets() {
  Object.keys(SCHEMAS).forEach(function (sheetName) {
    getSheet_(sheetName); // getSheet_ zaten oluşturma + başlık onarımını yapıyor
  });
  // Varsayılan boş "Sheet1" sekmesini temizlemek istersen elle silebilirsin.
}

/**
 * Sheet'i döndürür; yoksa oluşturur. Ayrıca SCHEMAS'ta tanımlı olup
 * fiziksel başlık satırında eksik olan alanları, MEVCUT sütunların
 * sırasını bozmadan sonuna ekler (öz-onarım).
 */
function getSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  ensureHeaders_(sheet, SCHEMAS[name] || []);
  return sheet;
}

/** Sheet'in fiziksel (gerçek) başlık satırını döndürür. */
function getHeaders_(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return [];
  return sheet.getRange(1, 1, 1, lastCol).getValues()[0].filter(function (h) { return h !== ''; });
}

/** Eksik alan başlıklarını sona ekler; var olan sütun sırasını korur. */
function ensureHeaders_(sheet, requiredFields) {
  let headers = getHeaders_(sheet);
  if (headers.length === 0) {
    if (requiredFields.length === 0) return [];
    sheet.getRange(1, 1, 1, requiredFields.length).setValues([requiredFields]);
    sheet.setFrozenRows(1);
    return requiredFields.slice();
  }
  const missing = requiredFields.filter(function (f) { return headers.indexOf(f) === -1; });
  if (missing.length > 0) {
    sheet.getRange(1, headers.length + 1, 1, missing.length).setValues([missing]);
    headers = headers.concat(missing);
  }
  return headers;
}

function sheetToObjects_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  const required = SCHEMAS[sheetName] || [];

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    if (required.length) {
      sheet.getRange(1, 1, 1, required.length).setValues([required]);
      sheet.setFrozenRows(1);
    }
    return [];
  }

  // TEK okuma: hem başlıkları hem veriyi aynı anda al (performans için —
  // önceden başlık kontrolü ve veri okuma ayrı ayrı yapılıyordu).
  const values = sheet.getDataRange().getValues();
  let headers = values.length ? values[0].filter(function (h) { return h !== ''; }) : [];

  // Eksik alan varsa (şema güncellendiyse) sadece bu durumda ek bir yazma yap.
  const missing = required.filter(function (f) { return headers.indexOf(f) === -1; });
  if (missing.length > 0) {
    sheet.getRange(1, headers.length + 1, 1, missing.length).setValues([missing]);
    headers = headers.concat(missing);
  }

  if (values.length < 2) return [];
  const rows = values.slice(1);
  return rows
    .filter(function (row) { return row[0] !== '' && row[0] !== null; })
    .map(function (row) {
      const obj = {};
      headers.forEach(function (h, i) { if (h !== '') obj[h] = row[i]; });
      return obj;
    });
}

function findRowIndexById_(sheet, id) {
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) return i + 1; // 1-based sheet row
  }
  return -1;
}

function jsonResponse_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

/** GET: liste çekme. ?action=list&sheet=Pets&petId=xxx (petId opsiyonel filtre) */
function doGet(e) {
  try {
    const action = e.parameter.action || 'list';
    const sheetName = e.parameter.sheet;

    if (action === 'list') {
      if (!sheetName || !SCHEMAS[sheetName]) {
        return jsonResponse_({ ok: false, error: 'Geçersiz sheet adı: ' + sheetName });
      }
      let data = sheetToObjects_(sheetName);
      if (e.parameter.petId) {
        data = data.filter(function (row) { return String(row.petId) === String(e.parameter.petId); });
      }
      return jsonResponse_({ ok: true, data: data });
    }

    if (action === 'listAll') {
      // Tüm tabloları tek seferde döner (uygulama ilk açılışta hızlı yüklensin diye)
      const result = {};
      Object.keys(SCHEMAS).forEach(function (name) {
        result[name] = sheetToObjects_(name);
      });
      return jsonResponse_({ ok: true, data: result });
    }

    return jsonResponse_({ ok: false, error: 'Bilinmeyen action: ' + action });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err) });
  }
}

/**
 * POST: create / update / delete.
 * Body (JSON, text/plain content-type ile gönderilmeli — CORS preflight'tan kaçınmak için):
 * { action: 'create'|'update'|'delete', sheet: 'Pets', id: '...', data: {...} }
 */
/**
 * ⏰ GÜNLÜK HATIRLATMA E-POSTASI
 * ------------------------------------------------------------------
 * Bu fonksiyonu SAAT 09:00'da otomatik çalışacak şekilde kurmak için,
 * Apps Script editöründe fonksiyon listesinden `setupDailyReminderTrigger`
 * fonksiyonunu SEÇİP BİR KEZ ▶️ Çalıştır'a bas. Bundan sonra her gün
 * saat 09:00 civarında (Apps Script tetikleyicileri tam dakika garanti
 * etmez, birkaç dakika sapma olabilir) sendDailyReminders() otomatik
 * çalışıp aşağıdaki durumlarda e-posta gönderir:
 *   - Yarın (1 gün sonra) tarihi gelen aşı/parazit kayıtları
 *   - Bugün tarihi gelen aşı/parazit kayıtları
 *   - Yarın check-in'i olan otel rezervasyonları
 *   - Bugün check-in'i olan otel rezervasyonları
 *   - Yarın/bugün biten ilaç-tedavi kayıtları
 * ------------------------------------------------------------------
 */
function setupDailyReminderTrigger() {
  // Aynı fonksiyon için önceden kurulmuş tetikleyicileri temizle (tekrar tekrar eklenmesin)
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sendDailyReminders') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('sendDailyReminders')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();
}

function getReminderEmail_() {
  const rows = sheetToObjects_('Settings');
  const row = rows.find(function (r) { return r.reminderEmail; });
  return row ? row.reminderEmail : '';
}

/** Trigger tarafından çağrılır; dönüş değeri önemsenmez. */
function sendDailyReminders() {
  runReminderCheck_();
}

/**
 * Asıl hatırlatma mantığı. Hem günlük trigger hem de uygulamadaki
 * "Hatırlatmaları Şimdi Kontrol Et" butonu bunu çağırır — böylece
 * 09:00'ı beklemeden aynı mantığı test edebilirsin.
 * Dönüş: { sent, count, email, reason }
 */
function runReminderCheck_() {
  const reminderEmail = getReminderEmail_();
  if (!reminderEmail) return { sent: false, count: 0, email: '', reason: 'no-email' };

  const pets = sheetToObjects_('Pets');
  const petNameById = {};
  pets.forEach(function (p) { petNameById[p.id] = p.name; });

  const todayStr = dateOnlyStr_(new Date());
  const tomorrowStr = dateOnlyStr_(addDays_(new Date(), 1));

  // Her hatırlatmayı yapılandırılmış obje olarak topla (şık e-posta şablonu için)
  const reminders = [];

  // 💉 Aşı & Parazit — hedef tarih bugün veya yarınsa
  sheetToObjects_('Vaccines').forEach(function (v) {
    if (!v.nextDate) return;
    const d = dateOnlyStr_(new Date(v.nextDate));
    const urgency = d === todayStr ? 'today' : (d === tomorrowStr ? 'tomorrow' : null);
    if (!urgency) return;
    reminders.push({ icon: '💉', type: (v.category || 'Aşı'), title: v.name, petName: petNameById[v.petId] || 'Bilinmeyen pet', urgency: urgency, dateLabel: 'Tarih: ' + dateOnlyStr_(new Date(v.nextDate)) });
  });

  // 💊 İlaç & Tedavi — DÜZELTİLDİ: artık sadece bitiş tarihine değil,
  // tedavi aktif olduğu HER GÜN "bugün ilaç günü" hatırlatması gönderir.
  // (startDate <= bugün <= endDate ise, endDate boşsa süresiz devam ediyor kabul edilir)
  sheetToObjects_('Medications').forEach(function (m) {
    if (!m.startDate) return; // başlangıç tarihi yoksa hangi gün aktif olduğunu bilemeyiz
    const startStr = dateOnlyStr_(new Date(m.startDate));
    const hasEnd = !!m.endDate;
    const endStr = hasEnd ? dateOnlyStr_(new Date(m.endDate)) : null;
    const petName = petNameById[m.petId] || 'Bilinmeyen pet';

    const isActiveToday = todayStr >= startStr && (!hasEnd || todayStr <= endStr);
    if (isActiveToday) {
      reminders.push({
        icon: '💊',
        type: m.frequency ? ('İlaç zamanı · ' + m.frequency) : 'İlaç zamanı',
        title: m.name,
        petName: petName,
        urgency: 'today',
        dateLabel: m.dosage ? ('Dozaj: ' + m.dosage) : 'Bugün verilmesi gereken ilaç'
      });
    }
    // Tedavi yarın sona eriyorsa ayrıca bir hatırlatma daha ekle
    if (hasEnd && endStr === tomorrowStr) {
      reminders.push({ icon: '💊', type: 'Tedavi yarın sona eriyor', title: m.name, petName: petName, urgency: 'tomorrow', dateLabel: 'Bitiş: ' + endStr });
    }
  });

  // 🏨 Otel check-in — hedef tarih bugün veya yarınsa
  sheetToObjects_('Boarding').forEach(function (b) {
    if (!b.checkIn) return;
    const d = dateOnlyStr_(new Date(b.checkIn));
    const urgency = d === todayStr ? 'today' : (d === tomorrowStr ? 'tomorrow' : null);
    if (!urgency) return;
    reminders.push({ icon: '🏨', type: 'Otel check-in', title: b.hotelName, petName: petNameById[b.petId] || 'Bilinmeyen pet', urgency: urgency, dateLabel: 'Giriş: ' + dateOnlyStr_(new Date(b.checkIn)) });
  });

  // 🥣 Mama stoku — kritik seviyeye düştüğünde (kırmızı bölge başı) ve
  // tükenmek üzereyken (kırmızı bölgenin ortası) hatırlatma gönder.
  // Her pet için sadece EN GÜNCEL stok kaydı dikkate alınır.
  const latestStockByPet = {};
  sheetToObjects_('NutritionLog').forEach(function (n) {
    if (n.logType !== 'stock' || !n.date || !n.estimatedDays) return;
    const existing = latestStockByPet[n.petId];
    if (!existing || new Date(n.date) > new Date(existing.date)) latestStockByPet[n.petId] = n;
  });
  Object.keys(latestStockByPet).forEach(function (petId) {
    const n = latestStockByPet[petId];
    const totalDays = Number(n.estimatedDays) || 0;
    if (totalDays <= 0) return;
    const elapsed = Math.round((new Date(todayStr) - new Date(dateOnlyStr_(new Date(n.date)))) / 86400000);
    const redStartDay = Math.round(totalDays * 0.75);   // kırmızı bölgeye giriş günü
    const redMidDay = Math.round(totalDays * 0.875);    // kırmızı bölgenin ortası
    const petName = petNameById[petId] || 'Bilinmeyen pet';
    const remainingDays = Math.max(totalDays - elapsed, 0);
    if (elapsed === redStartDay) {
      reminders.push({ icon: '🥣', type: 'Mama stoku kritik seviyede', title: n.title, petName: petName, urgency: 'today', dateLabel: 'Tahmini ' + remainingDays + ' gün stok kaldı' });
    } else if (elapsed === redMidDay) {
      reminders.push({ icon: '🥣', type: 'Mama stoku tükenmek üzere', title: n.title, petName: petName, urgency: 'today', dateLabel: 'Tahmini ' + remainingDays + ' gün stok kaldı' });
    }
  });

  if (reminders.length === 0) return { sent: false, count: 0, email: reminderEmail, reason: 'no-due-items' };

  // urgency sırasına göre diz: önce bugün, sonra yarın
  reminders.sort(function (a, b) { return (a.urgency === 'today' ? 0 : 1) - (b.urgency === 'today' ? 0 : 1); });

  const subject = '🐾 PetCare HQ — ' + reminders.length + ' hatırlatman var';
  const plainText = reminders.map(function (r) {
    return (r.urgency === 'today' ? 'BUGÜN' : 'YARIN') + ' — ' + r.petName + ': ' + r.title + ' (' + r.type + ')';
  }).join('\n');
  const htmlBody = buildReminderEmailHtml_(reminders);

  MailApp.sendEmail({ to: reminderEmail, subject: subject, body: plainText, htmlBody: htmlBody });
  return { sent: true, count: reminders.length, email: reminderEmail, reason: 'ok' };
}

/** Tarih/mantık işlemeden, sadece e-posta gönderiminin çalışıp çalışmadığını test eder. */
function sendTestReminderEmail_() {
  const reminderEmail = getReminderEmail_();
  if (!reminderEmail) return { ok: false, error: 'Önce uygulamadan (Pet Bilgilerini Düzenle) bir hatırlatma e-postası kaydet.' };

  const subject = '🐾 PetCare HQ — Test E-postası';
  const plainText = 'Bu bir test e-postasıdır. Bu e-postayı görüyorsan, hatırlatma sistemi ' + reminderEmail + ' adresine başarıyla e-posta gönderebiliyor demektir.';
  const html = ''
    + '<div style="background:#f8fafc;padding:24px 12px;font-family:Arial,sans-serif;">'
    + '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:480px;margin:0 auto;">'
    + '<tr><td><table role="presentation" cellpadding="0" cellspacing="0" style="background:#0f172a;border-radius:16px;width:100%;"><tr><td style="padding:22px 20px;">'
    + '<div style="font-family:Arial,sans-serif;font-weight:800;font-size:18px;color:#ffffff;">🐾 PetCare<span style="color:#93c5fd;">HQ</span></div>'
    + '<div style="font-family:Arial,sans-serif;font-size:13px;color:rgba(255,255,255,.7);margin-top:4px;">Test E-postası</div>'
    + '</td></tr></table></td></tr>'
    + '<tr><td style="padding-top:16px;"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#ffffff;border:1px solid #e6eaf0;border-radius:16px;"><tr><td style="padding:18px;">'
    + '<div style="font-family:Arial,sans-serif;font-size:14px;color:#0f172a;">✅ Tebrikler! Hatırlatma sistemi bu adrese e-posta gönderebiliyor: <b>' + escapeHtml_(reminderEmail) + '</b></div>'
    + '</td></tr></table></td></tr>'
    + '</table></div>';

  MailApp.sendEmail({ to: reminderEmail, subject: subject, body: plainText, htmlBody: html });
  return { ok: true, data: { email: reminderEmail } };
}

/** Uygulamanın tasarım diliyle uyumlu (lacivert/mavi, yuvarlak köşeli kart) HTML e-posta şablonu. */
function buildReminderEmailHtml_(reminders) {
  const NAVY = '#0f172a', BLUE = '#2563eb', BLUE_TINT = '#eff6ff', BG = '#f8fafc', BORDER = '#e6eaf0';
  const DANGER = '#dc2626', DANGER_TINT = '#fef2f2', AMBER = '#d97706', AMBER_TINT = '#fffbeb';

  function cardHtml(r) {
    const isToday = r.urgency === 'today';
    const badgeColor = isToday ? DANGER : AMBER;
    const badgeTint = isToday ? DANGER_TINT : AMBER_TINT;
    const badgeLabel = isToday ? 'BUGÜN' : 'YARIN';
    return ''
      + '<tr><td style="padding:0 0 12px 0;">'
      + '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#ffffff;border:1px solid ' + BORDER + ';border-radius:16px;">'
      + '<tr><td style="padding:16px 18px;">'
      + '<table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr>'
      + '<td style="font-size:22px;width:36px;vertical-align:top;">' + r.icon + '</td>'
      + '<td style="vertical-align:top;">'
      + '<span style="display:inline-block;background:' + badgeTint + ';color:' + badgeColor + ';font-size:11px;font-weight:700;padding:3px 10px;border-radius:999px;font-family:Arial,sans-serif;letter-spacing:.3px;">' + badgeLabel + '</span>'
      + '<div style="font-family:Arial,sans-serif;font-weight:800;font-size:15px;color:' + NAVY + ';margin-top:8px;">' + escapeHtml_(r.petName) + '</div>'
      + '<div style="font-family:Arial,sans-serif;font-size:13px;color:' + NAVY + ';margin-top:2px;">' + escapeHtml_(r.title) + ' <span style="color:#94a3b8;">· ' + escapeHtml_(r.type) + '</span></div>'
      + '<div style="font-family:Arial,sans-serif;font-size:12px;color:#94a3b8;margin-top:4px;">' + escapeHtml_(r.dateLabel) + '</div>'
      + '</td></tr></table>'
      + '</td></tr></table>'
      + '</td></tr>';
  }

  return ''
    + '<div style="background:' + BG + ';padding:24px 12px;font-family:Arial,sans-serif;">'
    + '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:480px;margin:0 auto;">'
    + '<tr><td style="padding-bottom:16px;">'
    + '<table role="presentation" cellpadding="0" cellspacing="0" style="background:' + NAVY + ';border-radius:16px;width:100%;">'
    + '<tr><td style="padding:22px 20px;">'
    + '<div style="font-family:Arial,sans-serif;font-weight:800;font-size:18px;color:#ffffff;">🐾 PetCare<span style="color:#93c5fd;">HQ</span></div>'
    + '<div style="font-family:Arial,sans-serif;font-size:13px;color:rgba(255,255,255,.7);margin-top:4px;">Bugünün hatırlatmaları — ' + reminders.length + ' kayıt</div>'
    + '</td></tr></table>'
    + '</td></tr>'
    + reminders.map(cardHtml).join('')
    + '<tr><td style="padding-top:6px;text-align:center;">'
    + '<div style="font-family:Arial,sans-serif;font-size:11px;color:#94a3b8;">Bu e-posta PetCare HQ uygulamandaki Ayarlar\'da girdiğin adrese otomatik gönderildi.</div>'
    + '</td></tr>'
    + '</table></div>';
}

function escapeHtml_(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function dateOnlyStr_(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}
function addDays_(d, n) {
  const copy = new Date(d);
  copy.setDate(copy.getDate() + n);
  return copy;
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;

    // Sheet gerektirmeyen özel action'lar (hatırlatma test/kontrol)
    if (action === 'runReminderCheckNow') {
      return jsonResponse_({ ok: true, data: runReminderCheck_() });
    }
    if (action === 'sendTestEmail') {
      return jsonResponse_(sendTestReminderEmail_());
    }

    const sheetName = body.sheet;

    if (!sheetName || !SCHEMAS[sheetName]) {
      return jsonResponse_({ ok: false, error: 'Geçersiz sheet adı: ' + sheetName });
    }

    const sheet = getSheet_(sheetName);
    const headers = getHeaders_(sheet); // her zaman GERÇEK fiziksel sütun sırası

    if (action === 'create') {
      const id = Utilities.getUuid();
      const now = new Date().toISOString();
      const record = Object.assign({}, body.data, { id: id, createdAt: now });
      const row = headers.map(function (h) { return record[h] !== undefined ? record[h] : ''; });
      sheet.appendRow(row);
      return jsonResponse_({ ok: true, data: record });
    }

    if (action === 'update') {
      const rowIndex = findRowIndexById_(sheet, body.id);
      if (rowIndex === -1) return jsonResponse_({ ok: false, error: 'Kayıt bulunamadı: ' + body.id });
      const existingValues = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
      const existing = {};
      headers.forEach(function (h, i) { existing[h] = existingValues[i]; });
      const updated = Object.assign({}, existing, body.data, { id: body.id });
      const row = headers.map(function (h) { return updated[h] !== undefined ? updated[h] : ''; });
      sheet.getRange(rowIndex, 1, 1, headers.length).setValues([row]);
      return jsonResponse_({ ok: true, data: updated });
    }

    if (action === 'delete') {
      const rowIndex = findRowIndexById_(sheet, body.id);
      if (rowIndex === -1) return jsonResponse_({ ok: false, error: 'Kayıt bulunamadı: ' + body.id });
      sheet.deleteRow(rowIndex);
      return jsonResponse_({ ok: true });
    }

    return jsonResponse_({ ok: false, error: 'Bilinmeyen action: ' + action });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err) });
  }
}
