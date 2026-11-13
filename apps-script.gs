/**
 * ダーツの旅 スポット帳 — スプレッドシート用スクリプト
 * スプレッドシートの「拡張機能 → Apps Script」に、これを丸ごと貼り付けて使います。
 * ページ（index.html）からの読み込み・書き込みを受け付けます。
 */
const SHEETS = ['スポット', 'コメント', '設定'];
const KANTO = [34.8, 138.3, 37.3, 141.0]; // 位置検索を関東周辺に寄せる範囲（南,西,北,東）

function doGet() {
  const ss = SpreadsheetApp.getActive();
  const data = {};
  SHEETS.forEach(function (name) {
    const sh = ss.getSheetByName(name);
    if (!sh || sh.getLastRow() < 1) { data[name] = []; return; }
    const v = sh.getDataRange().getDisplayValues();
    const head = v.shift();
    data[name] = v
      .filter(function (r) { return r.some(function (c) { return c !== ''; }); })
      .map(function (r) {
        const o = {};
        head.forEach(function (h, i) { if (h) o[h] = r[i]; });
        return o;
      });
  });
  return out({ ok: true, data: data });
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return out({ ok: false, error: 'bad request' }); }
  if (SHEETS.indexOf(req.sheet) < 0) return out({ ok: false, error: 'unknown sheet' });

  const row = clean(req.row || {});
  if (req.sheet === 'スポット' && !(row['緯度'] && row['経度']) &&
      (req.action === 'append' || row['住所'] || row['GoogleマップURL'])) {
    const found = {};
    const ll = locate(row, found);
    if (ll) { row['緯度'] = ll[0]; row['経度'] = ll[1]; }
    if (!row['名前'] && found.name) row['名前'] = found.name;   // リンクだけのときは店名も入れる
  }
  if (req.sheet === 'スポット' && req.action === 'append' && !row['名前']) row['名前'] = 'Googleマップのスポット';

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheetOf(req.sheet);
    if (req.action === 'append') {
      row['ID'] = Utilities.getUuid().slice(0, 8);
      row['日時'] = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
      const head = headers(sh, Object.keys(row));
      sh.appendRow(head.map(function (h) { return row[h] != null ? row[h] : ''; }));
      return out({ ok: true, id: row['ID'], row: row });
    }
    if (req.action === 'update') {
      const id = String(req.id || '');
      if (!id) return out({ ok: false, error: 'no id' });
      row['ID'] = id;
      const head = headers(sh, Object.keys(row));
      const col = head.indexOf('ID') + 1;
      const ids = sh.getLastRow() > 1 ? sh.getRange(2, col, sh.getLastRow() - 1, 1).getDisplayValues().map(function (r) { return r[0]; }) : [];
      const r = ids.indexOf(id);
      if (r < 0) {
        sh.appendRow(head.map(function (h) { return row[h] != null ? row[h] : ''; }));
      } else {
        Object.keys(row).forEach(function (k) { sh.getRange(r + 2, head.indexOf(k) + 1).setValue(row[k]); });
      }
      return out({ ok: true, id: id, row: row });
    }
    return out({ ok: false, error: 'unknown action' });
  } finally {
    lock.releaseLock();
  }
}

/* ---- 以下は内部で使う関数 ---- */

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function sheetOf(name) {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

// 1行目の見出しに足りない項目を右端に足し、見出しの並びを返す
function headers(sh, keys) {
  let head = sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0] : [];
  if (head.indexOf('ID') < 0 && head.every(function (h) { return h === ''; })) head = ['ID', '日時'];
  const add = keys.filter(function (k) { return head.indexOf(k) < 0; });
  if (add.length || sh.getLastRow() === 0) {
    head = head.concat(add);
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return head;
}

// 文字列にそろえ、長さを制限し、数式として解釈されないようにする
function clean(row) {
  const o = {};
  Object.keys(row).forEach(function (k) {
    if (!k || k === 'ID' || k === '日時') return;
    let v = row[k] == null ? '' : String(row[k]).trim().slice(0, 1000);
    if (/^[=+\-@]/.test(v)) v = "'" + v;
    o[String(k).slice(0, 40)] = v;
  });
  return o;
}

// GoogleマップのURL、住所、名前の順に位置（緯度・経度）を探す
function locate(row, found) {
  found = found || {};
  const url = row['GoogleマップURL'] || '';
  if (/^https?:\/\//.test(url)) {
    try {
      let u = url;
      for (let i = 0; i < 5; i++) {
        const pn = placeName(u);
        if (pn) found.name = pn;
        const ll = parseLL(u);
        if (ll) return ll;
        const res = UrlFetchApp.fetch(u, { followRedirects: false, muteHttpExceptions: true });
        const h = res.getHeaders();
        const next = h['Location'] || h['location'];
        if (!next) {
          const body = res.getContentText().slice(0, 300000);
          if (!found.name) {
            const t = body.match(/<meta[^>]+property="og:title"[^>]+content="([^"]+)"/) || body.match(/<title>([^<]+)<\/title>/);
            if (t) found.name = t[1].replace(/\s*[-·]\s*Google\s*(マップ|Maps).*$/, '').trim();
          }
          const b = parseLL(body); if (b) return b; break;
        }
        u = next;
      }
    } catch (err) {}
  }
  const nameQ = row['名前'] || found.name;
  const queries = [row['住所'], nameQ && row['住所'] ? nameQ + ' ' + row['住所'] : '', nameQ].filter(Boolean);
  for (let i = 0; i < queries.length; i++) {
    try {
      const g = Maps.newGeocoder().setLanguage('ja').setRegion('jp')
        .setBounds(KANTO[0], KANTO[1], KANTO[2], KANTO[3]).geocode(queries[i]);
      if (g.status === 'OK' && g.results.length) {
        const l = g.results[0].geometry.location;
        return [l.lat, l.lng];
      }
    } catch (err) {}
  }
  return null;
}

// URLの /maps/place/店名/ の部分から店名を取り出す
function placeName(s) {
  const m = String(s).match(/\/maps\/place\/([^/@?]+)/);
  if (!m) return '';
  try { return decodeURIComponent(m[1].replace(/\+/g, ' ')).trim(); } catch (err) { return ''; }
}

function parseLL(s) {
  try { s = decodeURIComponent(s); } catch (err) {}
  const m = s.match(/!3d(-?\d+\.\d+)!4d(-?\d+\.\d+)/) ||
            s.match(/@(-?\d+\.\d+),(-?\d+\.\d+)/) ||
            s.match(/[?&](?:q|query|ll|center|destination)=(-?\d+\.\d+),\s*(-?\d+\.\d+)/);
  return m ? [parseFloat(m[1]), parseFloat(m[2])] : null;
}

// 貼り付け後に一度だけ実行すると、権限の確認画面が出ます（デプロイ前の確認用）
function setup() {
  SHEETS.forEach(sheetOf);
  Logger.log(Maps.newGeocoder().setLanguage('ja').geocode('西千葉駅').status);
}
