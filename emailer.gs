/**
 * Ард Секюритиз ҮЦК · Өдөр Тутмын Ops Email
 * ─────────────────────────────────────────────────────────────────
 * ТОХИРГОО (3 алхам):
 *  1. CONFIG доорх утгуудыг өөрийн мэдээллээр солино
 *  2. createDailyTrigger() → Run → 09:00 автомат trigger үүснэ
 *  3. testSendToSelf()    → Run → тест email ирсэн эсэхийг шалгана
 * ─────────────────────────────────────────────────────────────────
 */

// ══════════════════════════════════════════════════════════════════
//  ⚙  ТОХИРГОО — Энд өөрчилнө үү
// ══════════════════════════════════════════════════════════════════
const CONFIG = {
  // Google Sheets ID (URL-ийн дундах урт тэмдэгтийн мөр)
  sheetId: 'YOUR_GOOGLE_SHEETS_ID_HERE',

  // Дашбоардын нийтийн URL (GitHub Pages эсвэл бусад)
  dashboardUrl: 'https://YOUR_SITE/dashboard.html',

  // Хүлээн авагчид
  recipients: {
    ceo:     'bilguun@ardsecurities.mn',   // Х.Билгүүн — Гүйцэтгэх Захирал
    cbdo:    'anujin@ardsecurities.mn',    // Б.Анужин — ББХЗ
    manager: 'manager@ardsecurities.mn',   // Шууд удирдлага
  },

  // GrapeCity SLA зорилт (хувиар)
  gcSlaTarget: 95,
};

// ══════════════════════════════════════════════════════════════════
//  ▶  ҮНДСЭН ФУНКЦ — Trigger дуудна
// ══════════════════════════════════════════════════════════════════
function sendMorningSummary() {
  try {
    const ss   = SpreadsheetApp.openById(CONFIG.sheetId);
    const data = readAllData(ss);
    MailApp.sendEmail({
      to:       Object.values(CONFIG.recipients).join(','),
      subject:  buildSubject(data),
      htmlBody: buildEmail(data),
      name:     'Ард Секюритиз · Ops',
    });
    Logger.log('✅ Email sent at ' + new Date().toISOString());
  } catch (e) {
    Logger.log('❌ sendMorningSummary error: ' + e);
    throw e;
  }
}

// Зөвхөн өөртөө тест илгээнэ
function testSendToSelf() {
  const ss   = SpreadsheetApp.openById(CONFIG.sheetId);
  const data = readAllData(ss);
  MailApp.sendEmail({
    to:       Session.getActiveUser().getEmail(),
    subject:  '[ТЕСТ] ' + buildSubject(data),
    htmlBody: buildEmail(data),
    name:     'Ард Секюритиз · Ops',
  });
  Logger.log('Test email → ' + Session.getActiveUser().getEmail());
}

// ══════════════════════════════════════════════════════════════════
//  ⏰  TRIGGER — НЭГ УДАА ажиллуулна
// ══════════════════════════════════════════════════════════════════
function createDailyTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'sendMorningSummary')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('sendMorningSummary')
    .timeBased()
    .atHour(9)
    .everyDays(1)
    .inTimezone('Asia/Ulaanbaatar')
    .create();

  Logger.log('✅ Daily trigger created: 09:00 Asia/Ulaanbaatar');
}

// ══════════════════════════════════════════════════════════════════
//  📊  ӨГӨГДӨЛ УНШИХ
// ══════════════════════════════════════════════════════════════════
function readTab(ss, name) {
  try {
    const sh = ss.getSheetByName(name);
    if (!sh) { Logger.log('⚠️ Tab not found: ' + name); return []; }
    const vals = sh.getDataRange().getValues();
    if (vals.length < 2) return [];
    const hdrs = vals[0].map(h => String(h).trim().toLowerCase().replace(/\s+/g, '_'));
    return vals.slice(1)
      .filter(r => r.some(c => c !== '' && c != null))
      .map(row => {
        const obj = {};
        hdrs.forEach((h, i) => { obj[h] = row[i]; });
        return obj;
      });
  } catch (e) { Logger.log('readTab [' + name + ']: ' + e); return []; }
}

function kv(rows) {
  const obj = {};
  rows.forEach(r => { if (r.key != null && r.key !== '') obj[String(r.key).trim()] = r.value; });
  return obj;
}
const n  = v => parseFloat(v) || 0;
const ni = v => parseInt(v)   || 0;

function readAllData(ss) {
  const sys = kv(readTab(ss, 'SystemStatus'));
  const gc  = kv(readTab(ss, 'GrapeCity'));
  const rec = kv(readTab(ss, 'Reconciliation'));
  const sup = kv(readTab(ss, 'SupportTickets'));
  const frc = kv(readTab(ss, 'FRCItems'));
  const bugR = readTab(ss, 'Bugs');
  const riR  = readTab(ss, 'ReconItems');
  const mfR  = readTab(ss, 'ManualFixes');

  return {
    system: {
      uptime30d:   n(sys.uptime_30d),
      lastBatch:   sys.last_batch   || '—',
      batchStatus: sys.batch_status || '—',
      batchNote:   sys.batch_note   || '',
    },
    bugs: {
      rows: bugR.map(r => ({
        severity: String(r.severity || '').toLowerCase(),
        a: ni(r.age_0_1d), b: ni(r.age_1_7d), c: ni(r.age_7_30d), d: ni(r.age_30plus),
      })),
    },
    grapeCity: {
      sla30d:    n(gc.sla_30d),
      open:      ni(gc.open_tickets),
      overdue:   ni(gc.overdue),
      avgDays:   n(gc.avg_resolution_days),
      slaTarget: n(gc.sla_target) || CONFIG.gcSlaTarget,
    },
    recon: {
      breaks:     ni(rec.breaks_today),
      valueMNT:   ni(rec.value_mnt),
      oldestDays: ni(rec.oldest_days),
      vsYday:     ni(rec.vs_yesterday),
      items: riR.map(r => ({ type: r.type || '', count: ni(r.count), value: ni(r.value_mnt) })),
    },
    support: {
      open:          ni(sup.open),
      newToday:      ni(sup.new_today),
      resolvedToday: ni(sup.resolved_today),
      slaBreaches:   ni(sup.sla_breaches),
      statusNew:     ni(sup.status_new),
      statusIP:      ni(sup.status_in_progress),
      statusPend:    ni(sup.status_pending),
    },
    frc: {
      pending:        ni(frc.pending),
      submittedToday: ni(frc.submitted_today),
      overdue:        ni(frc.overdue),
      nextDeadline:   frc.next_deadline || '—',
      nextType:       frc.next_type     || '—',
    },
    manualFixes: mfR.slice(0, 10).map(r => ({
      cls:   r['class'] || r.class_name || '',
      count: ni(r.count),
      trend: ni(r.trend),
    })),
  };
}

// ══════════════════════════════════════════════════════════════════
//  📧  EMAIL БҮТЭЭГЧ
// ══════════════════════════════════════════════════════════════════
function calcHealth(d) {
  const crit = d.bugs.rows.reduce((s, r) => s + (r.severity === 'critical' ? r.a+r.b+r.c+r.d : 0), 0);
  if (crit > 0 || d.frc.overdue > 0 || d.recon.oldestDays > 3)                                  return 'red';
  if (d.grapeCity.overdue > 0 || d.recon.breaks > 2 ||
      d.grapeCity.sla30d < d.grapeCity.slaTarget || d.support.slaBreaches > 0)                  return 'amber';
  return 'green';
}

function fmtMNT(v) {
  if (v >= 1e9) return '₮' + (v / 1e9).toFixed(1) + ' тэрбум';
  if (v >= 1e6) return '₮' + (v / 1e6).toFixed(0) + ' сая';
  return '₮' + v.toLocaleString();
}

function buildSubject(d) {
  const emoji = { red: '🔴', amber: '🟡', green: '🟢' }[calcHealth(d)];
  const date  = Utilities.formatDate(new Date(), 'Asia/Ulaanbaatar', 'yyyy-MM-dd');
  return emoji + ' [Ард Секюритиз] Өглөөний Тойм — ' + date;
}

// ── HTML template helpers ──────────────────────────────────────
function card(title, icon, body) {
  return `
  <tr><td style="padding-top:14px">
    <table width="100%" cellpadding="0" cellspacing="0"
           style="background:#0D1829;border:1px solid rgba(255,255,255,.07);border-radius:10px;overflow:hidden;border-collapse:separate">
      <tr><td style="background:rgba(201,168,76,.07);border-bottom:1px solid rgba(255,255,255,.06);
                     padding:11px 18px;font-size:12.5px;font-weight:700;color:#EEF2F8">
        ${icon}&nbsp; ${title}
      </td></tr>
      <tr><td style="padding:14px 18px">${body}</td></tr>
    </table>
  </td></tr>`;
}

function kpiRow(items) {
  const w = Math.floor(100 / items.length) + '%';
  const cells = items.map((it, i) =>
    `<td width="${w}" align="center"
         style="padding:10px 8px;${i < items.length - 1 ? 'border-right:1px solid rgba(255,255,255,.05)' : ''}">
       <div style="font-size:26px;font-weight:800;color:${it.color || '#EEF2F8'};font-family:Arial,sans-serif;line-height:1">${it.val}</div>
       <div style="font-size:9px;color:#405566;font-weight:700;text-transform:uppercase;letter-spacing:.8px;margin-top:4px">${it.lbl}</div>
     </td>`
  ).join('');
  return `<table width="100%" cellpadding="0" cellspacing="0"><tr>${cells}</tr></table>`;
}

function drow(lbl, val, col) {
  return `<tr>
    <td style="padding:6px 0;border-bottom:1px solid rgba(255,255,255,.04);font-size:12px;color:#8FA3BE">${lbl}</td>
    <td align="right" style="padding:6px 0;border-bottom:1px solid rgba(255,255,255,.04);font-size:12.5px;font-weight:700;color:${col || '#EEF2F8'}">${val}</td>
  </tr>`;
}

function statusOf(ok, warn, bad, val) {
  if (val <= ok)   return '#4ADE80';
  if (val <= warn) return '#FCD34D';
  return '#FF5A52';
}

// ── Main buildEmail ────────────────────────────────────────────
function buildEmail(d) {
  const h     = calcHealth(d);
  const date  = Utilities.formatDate(new Date(), 'Asia/Ulaanbaatar', 'yyyy-MM-dd');
  const time  = Utilities.formatDate(new Date(), 'Asia/Ulaanbaatar', 'HH:mm');

  const hColor = { red: '#FF5A52', amber: '#FCD34D', green: '#4ADE80' }[h];
  const hBg    = { red: 'rgba(230,51,41,.1)', amber: 'rgba(245,158,11,.1)', green: 'rgba(34,197,94,.1)' }[h];
  const hBdr   = { red: 'rgba(230,51,41,.4)', amber: 'rgba(245,158,11,.4)', green: 'rgba(34,197,94,.4)' }[h];
  const hLabel = { red: '🔴 АНХААРАХ ШААРДЛАГАТАЙ', amber: '🟡 СОНОР БАЙЛГАНА УУ', green: '🟢 ХЭВИЙН' }[h];

  // Bug totals
  const crit  = d.bugs.rows.reduce((s,r)=>s+(r.severity==='critical'?r.a+r.b+r.c+r.d:0),0);
  const high  = d.bugs.rows.reduce((s,r)=>s+(r.severity==='high'    ?r.a+r.b+r.c+r.d:0),0);
  const med   = d.bugs.rows.reduce((s,r)=>s+(r.severity==='medium'  ?r.a+r.b+r.c+r.d:0),0);
  const low   = d.bugs.rows.reduce((s,r)=>s+(r.severity==='low'     ?r.a+r.b+r.c+r.d:0),0);

  // Bug matrix rows HTML
  const bugMatrixRows = [
    ['critical','Critical','#FF5A52'],['high','High','#FCD34D'],
    ['medium','Medium','#60A5FA'],['low','Low','#8FA3BE'],
  ].map(([sv,lbl,col]) => {
    const row = d.bugs.rows.find(r => r.severity === sv) || {a:0,b:0,c:0,d:0};
    const t   = row.a+row.b+row.c+row.d;
    const c   = (v) => `<td align="center" style="padding:5px 8px;border-bottom:1px solid rgba(255,255,255,.04);font-size:12px;font-weight:700;color:${v>0?col:'#405566'}">${v||'—'}</td>`;
    return `<tr>
      <td style="padding:5px 8px;border-bottom:1px solid rgba(255,255,255,.04)">
        <span style="font-size:11px;font-weight:700;color:${col};background:${col}22;padding:2px 8px;border-radius:4px">${lbl}</span>
      </td>
      ${c(row.a)}${c(row.b)}${c(row.c)}${c(row.d)}
      <td align="center" style="padding:5px 8px;border-bottom:1px solid rgba(255,255,255,.04);font-size:12px;font-weight:800;color:#E5C46A">${t}</td>
    </tr>`;
  }).join('');

  // Recon items HTML
  const reconRows = d.recon.items.length
    ? d.recon.items.map(it =>
        `<tr>
          <td style="padding:5px 0;border-bottom:1px solid rgba(255,255,255,.04);font-size:12px;color:#8FA3BE">${it.type}</td>
          <td align="right" style="padding:5px 0;border-bottom:1px solid rgba(255,255,255,.04);font-size:12px;font-weight:700;color:#FCD34D">${it.count} зөрүү</td>
          <td align="right" style="padding:5px 0;border-bottom:1px solid rgba(255,255,255,.04);font-size:11px;color:#405566">${fmtMNT(it.value)}</td>
        </tr>`).join('')
    : `<tr><td colspan="3" style="padding:8px 0;font-size:12px;color:#4ADE80">✓ Зөрүүгүй</td></tr>`;

  // Top 5 manual fixes
  const top5 = d.manualFixes.slice(0, 5).map((f, i) => {
    const td = f.trend > 0 ? `<span style="color:#FF5A52">+${f.trend}</span>`
             : f.trend < 0 ? `<span style="color:#4ADE80">${f.trend}</span>`
             : `<span style="color:#405566">—</span>`;
    return `<tr>
      <td style="padding:5px 0;border-bottom:1px solid rgba(255,255,255,.04);font-size:12px;color:#405566;width:20px">${i+1}</td>
      <td style="padding:5px 0;border-bottom:1px solid rgba(255,255,255,.04);font-size:12px;color:#8FA3BE">${f.cls}</td>
      <td align="right" style="padding:5px 0;border-bottom:1px solid rgba(255,255,255,.04);font-size:12px;font-weight:700;color:#E5C46A">${f.count}</td>
      <td align="right" style="padding:5px 0;border-bottom:1px solid rgba(255,255,255,.04);font-size:11px;width:28px">${td}</td>
    </tr>`;
  }).join('');

  const bsOk = d.system.batchStatus === 'OK';
  const gcPct = d.grapeCity.sla30d;

  // Summary snapshot table
  const snapRows = [
    ['CASPO / Polaris', bsOk ? '✅ OK' : '❌ ' + d.system.batchStatus, `Uptime ${d.system.uptime30d}% · Batch ${d.system.lastBatch}`, bsOk ? '#4ADE80' : '#FF5A52'],
    ['Нээлттэй Bug',  crit > 0 ? '🔴 ' + crit + ' Critical' : high > 0 ? '🟡 ' + high + ' High' : '🟢 OK',
     `C:${crit} H:${high} M:${med} L:${low} — Нийт ${crit+high+med+low}`,
     crit > 0 ? '#FF5A52' : high > 0 ? '#FCD34D' : '#4ADE80'],
    ['GrapeCity SLA',  gcPct >= d.grapeCity.slaTarget ? '✅ ' + gcPct + '%' : '🟡 ' + gcPct + '%',
     `Зорилт ${d.grapeCity.slaTarget}% · Нээлттэй ${d.grapeCity.open} · Хэтэрсэн ${d.grapeCity.overdue}`,
     gcPct >= d.grapeCity.slaTarget ? '#4ADE80' : '#FCD34D'],
    ['MSE/MCSD Recon',  d.recon.breaks === 0 ? '✅ OK' : '🟡 ' + d.recon.breaks + ' зөрүү',
     `${fmtMNT(d.recon.valueMNT)} · Хамгийн хуучин: ${d.recon.oldestDays} өдөр`,
     d.recon.breaks === 0 ? '#4ADE80' : d.recon.breaks <= 2 ? '#FCD34D' : '#FF5A52'],
    ['ҮЦК Дэмжлэг',  d.support.slaBreaches > 0 ? '🔴 SLA зөрчил' : '✅ ' + d.support.open + ' нээлттэй',
     `Шинэ ${d.support.newToday} · Шийдсэн ${d.support.resolvedToday} · SLA зөрчил ${d.support.slaBreaches}`,
     d.support.slaBreaches > 0 ? '#FF5A52' : '#4ADE80'],
    ['FRC Мэдүүлэг',  d.frc.overdue > 0 ? '🔴 ' + d.frc.overdue + ' хэтэрсэн' : d.frc.pending > 0 ? '🟡 ' + d.frc.pending + ' хүлээгдэж буй' : '✅ OK',
     `Дараагийн хугацаа ${d.frc.nextDeadline} · ${d.frc.nextType}`,
     d.frc.overdue > 0 ? '#FF5A52' : d.frc.pending > 0 ? '#FCD34D' : '#4ADE80'],
    ['Гар засвар (Top)', '📊 ' + (d.manualFixes[0] ? d.manualFixes[0].cls : '—'),
     d.manualFixes[0] ? `${d.manualFixes[0].count} удаа (энэ сар)` : '—', '#E5C46A'],
  ].map(([domain, status, detail, col]) => `
    <tr>
      <td style="padding:8px 12px;border-bottom:1px solid rgba(255,255,255,.05);font-size:12px;font-weight:700;color:#EEF2F8;width:160px">${domain}</td>
      <td style="padding:8px 12px;border-bottom:1px solid rgba(255,255,255,.05);font-size:12px;font-weight:700;color:${col}">${status}</td>
      <td style="padding:8px 12px;border-bottom:1px solid rgba(255,255,255,.05);font-size:11px;color:#8FA3BE">${detail}</td>
    </tr>`).join('');

  // ── Full HTML ──────────────────────────────────────────────
  return `<!DOCTYPE html>
<html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0"></head>
<body style="margin:0;padding:0;background:#080F1C;font-family:Arial,sans-serif;color:#EEF2F8">
<table border="0" cellpadding="0" cellspacing="0" width="100%" style="background:#080F1C">
<tr><td align="center" style="padding:20px 12px">
<table border="0" cellpadding="0" cellspacing="0" width="620" style="max-width:620px;width:100%">

  <!-- HEADER -->
  <tr><td style="background:#0D1829;border:1px solid rgba(201,168,76,.3);border-radius:12px 12px 0 0;padding:22px 24px">
    <table width="100%" cellpadding="0" cellspacing="0"><tr>
      <td>
        <div style="font-size:18px;font-weight:700;color:#EEF2F8;margin-bottom:3px">Ард Секюритиз ҮЦК</div>
        <div style="font-size:9.5px;font-weight:700;color:#C9A84C;text-transform:uppercase;letter-spacing:2px">ӨГЛӨӨНИЙ ҮЙЛАЖИЛЛАГААНЫ ТОЙМ</div>
      </td>
      <td align="right">
        <div style="font-size:13px;font-weight:700;color:#EEF2F8">${date}</div>
        <div style="font-size:10px;color:#405566">Илгээсэн: ${time}</div>
      </td>
    </tr></table>
  </td></tr>

  <!-- STATUS STRIP -->
  <tr><td style="background:${hBg};border-left:1px solid ${hBdr};border-right:1px solid ${hBdr};padding:12px 24px">
    <table width="100%" cellpadding="0" cellspacing="0"><tr>
      <td><span style="font-size:13px;font-weight:800;color:${hColor}">${hLabel}</span></td>
      <td align="right" style="font-size:11px;color:#8FA3BE">
        CASPO: <strong style="color:${bsOk ? '#4ADE80' : '#FF5A52'}">${d.system.batchStatus}</strong>
        &nbsp;·&nbsp; Batch: <strong style="color:#EEF2F8">${d.system.lastBatch}</strong>
      </td>
    </tr></table>
  </td></tr>

  <!-- SNAPSHOT TABLE -->
  ${card('Өнөөдрийн Хураангуй', '📋',
    `<table width="100%" cellpadding="0" cellspacing="0">${snapRows}</table>`
  )}

  <!-- BUGS MATRIX -->
  ${card('Bug Дутагдалуудын Матриц — Jira', '🐛',
    `<table width="100%" cellpadding="0" cellspacing="0">
      <tr>
        <th style="padding:5px 8px;font-size:9px;font-weight:700;text-transform:uppercase;color:#405566;letter-spacing:1px;border-bottom:1px solid rgba(255,255,255,.08);text-align:left">Ноцтой байдал</th>
        <th style="padding:5px 8px;font-size:9px;font-weight:700;text-transform:uppercase;color:#405566;letter-spacing:1px;border-bottom:1px solid rgba(255,255,255,.08);text-align:center">&lt;1 өдөр</th>
        <th style="padding:5px 8px;font-size:9px;font-weight:700;text-transform:uppercase;color:#405566;letter-spacing:1px;border-bottom:1px solid rgba(255,255,255,.08);text-align:center">1–7 өдөр</th>
        <th style="padding:5px 8px;font-size:9px;font-weight:700;text-transform:uppercase;color:#405566;letter-spacing:1px;border-bottom:1px solid rgba(255,255,255,.08);text-align:center">7–30 өдөр</th>
        <th style="padding:5px 8px;font-size:9px;font-weight:700;text-transform:uppercase;color:#405566;letter-spacing:1px;border-bottom:1px solid rgba(255,255,255,.08);text-align:center">&gt;30 өдөр</th>
        <th style="padding:5px 8px;font-size:9px;font-weight:700;text-transform:uppercase;color:#405566;letter-spacing:1px;border-bottom:1px solid rgba(255,255,255,.08);text-align:center">Нийт</th>
      </tr>
      ${bugMatrixRows}
    </table>`
  )}

  <!-- RECON -->
  ${card('MSE/MCSD Тооцооны Зөрүү', '⚖️',
    kpiRow([
      {val: d.recon.breaks,   lbl: 'Зөрүү өнөөдөр',   color: statusOf(0, 2, 99, d.recon.breaks)},
      {val: fmtMNT(d.recon.valueMNT), lbl: 'Нийт дүн', color: '#FCD34D'},
      {val: d.recon.oldestDays + ' өдөр', lbl: 'Хамгийн хуучин', color: statusOf(0, 2, 99, d.recon.oldestDays)},
    ]) +
    `<div style="margin-top:12px"><table width="100%" cellpadding="0" cellspacing="0">${reconRows}</table></div>`
  )}

  <!-- GC + SUPPORT side by side -->
  <tr><td style="padding-top:14px">
    <table width="100%" cellpadding="0" cellspacing="0"><tr>

      <td width="48%" valign="top" style="padding-right:8px">
        <table width="100%" cellpadding="0" cellspacing="0"
               style="background:#0D1829;border:1px solid rgba(255,255,255,.07);border-radius:10px;overflow:hidden;border-collapse:separate">
          <tr><td style="background:rgba(201,168,76,.07);border-bottom:1px solid rgba(255,255,255,.06);padding:11px 16px;font-size:12.5px;font-weight:700;color:#EEF2F8">📊 GrapeCity SLA</td></tr>
          <tr><td style="padding:14px 16px">
            ${kpiRow([
              {val: gcPct + '%',         lbl: 'SLA (30 хоног)', color: gcPct >= d.grapeCity.slaTarget ? '#4ADE80' : '#FCD34D'},
              {val: d.grapeCity.open,    lbl: 'Нээлттэй',       color: '#EEF2F8'},
              {val: d.grapeCity.overdue, lbl: 'Хэтэрсэн',       color: d.grapeCity.overdue > 0 ? '#FF5A52' : '#4ADE80'},
            ])}
            <div style="margin-top:10px">
            <table width="100%" cellpadding="0" cellspacing="0">
              ${drow('Зорилт', d.grapeCity.slaTarget + '%', '#8FA3BE')}
              ${drow('Дундаж шийдлэлт', d.grapeCity.avgDays + ' өдөр', '#8FA3BE')}
            </table></div>
          </td></tr>
        </table>
      </td>

      <td width="4%"></td>

      <td width="48%" valign="top" style="padding-left:8px">
        <table width="100%" cellpadding="0" cellspacing="0"
               style="background:#0D1829;border:1px solid rgba(255,255,255,.07);border-radius:10px;overflow:hidden;border-collapse:separate">
          <tr><td style="background:rgba(201,168,76,.07);border-bottom:1px solid rgba(255,255,255,.06);padding:11px 16px;font-size:12.5px;font-weight:700;color:#EEF2F8">🎫 ҮЦК Дэмжлэг</td></tr>
          <tr><td style="padding:14px 16px">
            ${kpiRow([
              {val: d.support.open,          lbl: 'Нээлттэй',    color: '#E5C46A'},
              {val: d.support.newToday,      lbl: 'Шинэ',        color: '#60A5FA'},
              {val: d.support.resolvedToday, lbl: 'Шийдсэн',     color: '#4ADE80'},
            ])}
            <div style="margin-top:10px">
            <table width="100%" cellpadding="0" cellspacing="0">
              ${drow('Шинэ', d.support.statusNew, '#8FA3BE')}
              ${drow('Хянаж буй', d.support.statusIP, '#8FA3BE')}
              ${drow('SLA зөрчил', d.support.slaBreaches, d.support.slaBreaches > 0 ? '#FF5A52' : '#4ADE80')}
            </table></div>
          </td></tr>
        </table>
      </td>

    </tr></table>
  </td></tr>

  <!-- FRC -->
  ${card('FRC Мэдүүлэх Зүйлс — СЗХ', '🏛️',
    kpiRow([
      {val: d.frc.pending,        lbl: 'Хүлээгдэж буй',   color: d.frc.pending > 0 ? '#FCD34D' : '#4ADE80'},
      {val: d.frc.submittedToday, lbl: 'Өнөөдөр илгээсэн',color: '#4ADE80'},
      {val: d.frc.overdue,        lbl: 'Хугацаа хэтэрсэн', color: d.frc.overdue > 0 ? '#FF5A52' : '#4ADE80'},
    ]) +
    `<div style="margin-top:10px"><table width="100%" cellpadding="0" cellspacing="0">
      ${drow('Дараагийн хугацаа', d.frc.nextDeadline, '#E5C46A')}
      ${drow('Тайлангийн төрөл',  d.frc.nextType,      '#8FA3BE')}
    </table></div>`
  )}

  <!-- MANUAL FIXES TOP 5 -->
  ${card('Гар Засварын Шилдэг 5 Ангилал', '🔧',
    `<table width="100%" cellpadding="0" cellspacing="0">
      <tr>
        <th style="padding:4px 8px;font-size:9px;color:#405566;text-align:left;border-bottom:1px solid rgba(255,255,255,.08)">#</th>
        <th style="padding:4px 8px;font-size:9px;color:#405566;text-align:left;border-bottom:1px solid rgba(255,255,255,.08)">Ангилал</th>
        <th style="padding:4px 8px;font-size:9px;color:#405566;text-align:right;border-bottom:1px solid rgba(255,255,255,.08)">Тоо</th>
        <th style="padding:4px 8px;font-size:9px;color:#405566;text-align:right;border-bottom:1px solid rgba(255,255,255,.08)">Өөрчлөлт</th>
      </tr>
      ${top5}
    </table>`
  )}

  <!-- CTA BUTTON -->
  <tr><td align="center" style="padding:24px 0 16px">
    <a href="${CONFIG.dashboardUrl}"
       style="display:inline-block;padding:13px 32px;background:linear-gradient(135deg,#C9A84C,#8A6A1A);color:#040B15;font-weight:700;font-size:14px;text-decoration:none;border-radius:10px;letter-spacing:.3px">
      Дашбоард Нэвтрэх →
    </a>
    <div style="margin-top:8px;font-size:10.5px;color:#405566">${CONFIG.dashboardUrl}</div>
  </td></tr>

  <!-- FOOTER -->
  <tr><td style="background:#0D1829;border:1px solid rgba(201,168,76,.15);border-radius:0 0 12px 12px;padding:16px 24px">
    <table width="100%" cellpadding="0" cellspacing="0"><tr>
      <td style="font-size:11px;color:#405566">
        <strong style="color:#8FA3BE">Ард Секюритиз ҮЦК</strong> · Ops Dashboard<br>
        Автомат мэйл — Хариулах шаардлагагүй
      </td>
      <td align="right" style="font-size:10px;color:#405566">
        <a href="${CONFIG.dashboardUrl}" style="color:#C9A84C;text-decoration:none">Дашбоард</a>
        &nbsp;·&nbsp;
        <a href="mailto:${CONFIG.recipients.manager}" style="color:#C9A84C;text-decoration:none">Холбоо барих</a>
      </td>
    </tr></table>
  </td></tr>

</table>
</td></tr></table>
</body></html>`;
}
