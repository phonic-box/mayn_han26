/**
 * May'n ひらがな はんぐじゃむ 披露楽曲管理 — サーバー側（Google Apps Script）
 *
 * 永続データは Songs / Shows の 2 シートのみ。
 * 「全公演」「未披露」は画面側で算出する仮想選択肢で、シートには保存しない。
 */

var APP_TITLE = "May'n ひらがな はんぐじゃむ｜楽曲管理";
var SONGS_SHEET = 'Songs';
var SHOWS_SHEET = 'Shows';
var SONG_HEADERS = ['song_id', 'title', 'reading', 'performed_show_ids'];
var SHOW_HEADERS = ['show_id', 'show_name', 'show_date', 'venue', 'sort_order'];
var LOCK_TIMEOUT_MS = 10000;

// 説明・動作確認用のダミーデータ。実データではない。
var SAMPLE_SHOWS = [
  ['P001', '東京公演（サンプル）', '', '', 1],
  ['P002', '愛知公演（サンプル）', '', '', 2],
  ['P003', '京都公演（サンプル）', '', '', 3],
  ['P004', '東京公演 2日目（サンプル）', '', '', 4]
];
var SAMPLE_SONGS = [
  ['S001', '楽曲A', 'がっきょくえー', '[]'],
  ['S002', '楽曲B', 'がっきょくびー', '[]'],
  ['S003', '楽曲C', 'がっきょくしー', '[]'],
  ['S004', 'あいうえおの歌（サンプル）', 'あいうえおのうた', '[]'],
  ['S005', 'Sample <Song> & "記号"', 'さんぷるそんぐ', '[]'],
  ['S006', 'よみ未登録のサンプル曲', '', '[]']
];

// ---------------------------------------------------------------------------
// Web アプリ
// ---------------------------------------------------------------------------

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle(APP_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

/** Index.html から Styles / Logic / Scripts を取り込む。 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/** スプレッドシートを開いたときに管理メニューを追加する（コンテナバインド時）。 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('楽曲管理')
    .addItem('シートを初期化（サンプルデータ投入）', 'setupSheets')
    .addToUi();
}

// ---------------------------------------------------------------------------
// 画面から呼ぶ関数（google.script.run）
// ---------------------------------------------------------------------------

/** @return {{shows: Object[], songs: Object[], fetchedAt: string}} */
function getInitialData() {
  var ss = getSpreadsheet_();
  var shows = readShows_(ss);
  var songs = readSongs_(ss).songs.map(toClientSong_);
  return { shows: shows, songs: songs, fetchedAt: new Date().toISOString() };
}

/**
 * 指定公演での披露状態を ON/OFF する。他公演の記録は保持する。
 * @return {Object} 更新後の楽曲
 */
function setSongPerformance(songId, showId, performed) {
  songId = requireId_(songId, 'songId');
  showId = requireId_(showId, 'showId');
  if (typeof performed !== 'boolean') {
    throw new Error('performed は true / false で指定してください。');
  }

  return withLock_(function () {
    var ss = getSpreadsheet_();
    var shows = readShows_(ss);
    if (!shows.some(function (s) { return s.show_id === showId; })) {
      throw new Error('存在しない公演 ID です: ' + showId);
    }

    // ロック取得後に最新の値を読み直し、対象公演だけを差分更新する。
    var table = readSongs_(ss);
    var song = findById_(table.songs, songId);
    if (!song) {
      throw new Error('存在しない楽曲 ID です: ' + songId);
    }
    if (song.error) {
      throw new Error('「' + song.title + '」の performed_show_ids が壊れているため更新できません。' +
        'スプレッドシート ' + SONGS_SHEET + ' の ' + song.rowNumber + ' 行目を確認してください。');
    }

    var ids = unique_(song.ids.filter(function (id) { return id !== showId; }));
    if (performed) ids.push(showId);
    ids = sortShowIds_(ids, shows);

    table.sheet
      .getRange(song.rowNumber, table.columns.performed_show_ids)
      .setValue(JSON.stringify(ids));
    SpreadsheetApp.flush();

    song.ids = ids;
    return toClientSong_(song);
  });
}

// ---------------------------------------------------------------------------
// 初期設定（エディタまたはメニューから手動実行）
// ---------------------------------------------------------------------------

/** Songs / Shows シートが無ければ作成し、ヘッダーとダミーデータを入れる。既存データは上書きしない。 */
function setupSheets() {
  var ss = getSpreadsheet_();
  var created = [];
  if (ensureSheet_(ss, SHOWS_SHEET, SHOW_HEADERS, SAMPLE_SHOWS, 4)) created.push(SHOWS_SHEET);
  if (ensureSheet_(ss, SONGS_SHEET, SONG_HEADERS, SAMPLE_SONGS, 4)) created.push(SONGS_SHEET);
  return created.length
    ? created.join(', ') + ' シートを作成しました。'
    : '両シートとも既にデータがあるため、変更していません。';
}

function ensureSheet_(ss, name, headers, sampleRows, textColumnCount) {
  var sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sheet.getLastRow() > 0) return false;

  sheet.getRange(1, 1, sheet.getMaxRows(), textColumnCount).setNumberFormat('@');
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, sampleRows.length, headers.length).setValues(sampleRows);
  sheet.setFrozenRows(1);
  return true;
}

// ---------------------------------------------------------------------------
// シート読み込み
// ---------------------------------------------------------------------------

function getSpreadsheet_() {
  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  var ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('スプレッドシートが見つかりません。スクリプトプロパティ SPREADSHEET_ID を設定してください。');
  }
  return ss;
}

/** ヘッダー行で列を特定し、表示値（文字列）で全行を読む。 */
function readTable_(ss, name, requiredHeaders) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    throw new Error(name + ' シートがありません。setupSheets を実行してください。');
  }
  var values = sheet.getDataRange().getDisplayValues();
  if (!values.length || !values[0].join('')) {
    throw new Error(name + ' シートにヘッダー行がありません。setupSheets を実行してください。');
  }

  var columns = {};
  values[0].forEach(function (h, i) {
    var key = String(h).trim();
    if (key && !columns[key]) columns[key] = i + 1;
  });
  var missing = requiredHeaders.filter(function (h) { return !columns[h]; });
  if (missing.length) {
    throw new Error(name + ' シートに列 ' + missing.join(', ') + ' がありません。');
  }

  var rows = [];
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    if (!row.join('').trim()) continue;
    rows.push({
      rowNumber: r + 1,
      get: (function (cells) {
        return function (key) { return String(cells[columns[key] - 1]).trim(); };
      })(row)
    });
  }
  return { sheet: sheet, columns: columns, rows: rows };
}

function readShows_(ss) {
  var table = readTable_(ss, SHOWS_SHEET, SHOW_HEADERS);
  var seen = {};
  var shows = [];
  table.rows.forEach(function (row) {
    var id = row.get('show_id');
    if (!id) return;
    if (seen[id]) throw new Error('Shows シートで show_id "' + id + '" が重複しています。');
    seen[id] = true;
    var order = Number(row.get('sort_order'));
    shows.push({
      show_id: id,
      show_name: row.get('show_name') || id,
      show_date: normalizeDate_(row.get('show_date')),
      venue: row.get('venue'),
      sort_order: isFinite(order) && row.get('sort_order') !== '' ? order : Number.MAX_SAFE_INTEGER,
      _row: row.rowNumber
    });
  });
  shows.sort(function (a, b) { return a.sort_order - b.sort_order || a._row - b._row; });
  return shows.map(function (s) {
    return { show_id: s.show_id, show_name: s.show_name, show_date: s.show_date, venue: s.venue, sort_order: s.sort_order };
  });
}

function readSongs_(ss) {
  var table = readTable_(ss, SONGS_SHEET, SONG_HEADERS);
  var seen = {};
  var songs = [];
  table.rows.forEach(function (row) {
    var id = row.get('song_id');
    if (!id) return;
    if (seen[id]) throw new Error('Songs シートで song_id "' + id + '" が重複しています。');
    seen[id] = true;
    var parsed = parseShowIds_(row.get('performed_show_ids'));
    songs.push({
      song_id: id,
      title: row.get('title') || id,
      reading: row.get('reading'),
      ids: parsed.ids,
      error: parsed.error,
      rowNumber: row.rowNumber
    });
  });
  return { sheet: table.sheet, columns: table.columns, songs: songs };
}

/** 破損データは空配列に読み替えず、error を付けて返す（更新を拒否するため）。 */
function parseShowIds_(raw) {
  if (raw === '') return { ids: [], error: '' };
  var value;
  try {
    value = JSON.parse(raw);
  } catch (e) {
    return { ids: [], error: 'JSON として読めません: ' + raw };
  }
  if (!Array.isArray(value) || value.some(function (v) { return typeof v !== 'string'; })) {
    return { ids: [], error: '文字列の配列ではありません: ' + raw };
  }
  return { ids: unique_(value), error: '' };
}

function toClientSong_(song) {
  return {
    song_id: song.song_id,
    title: song.title,
    reading: song.reading,
    performed_show_ids: song.ids.slice(),
    data_error: song.error || ''
  };
}

// ---------------------------------------------------------------------------
// 補助
// ---------------------------------------------------------------------------

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(LOCK_TIMEOUT_MS)) {
    throw new Error('他の人が更新中です。少し待ってからもう一度お試しください。');
  }
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function requireId_(value, label) {
  if (typeof value !== 'string' || !value.trim()) {
    throw new Error(label + ' が指定されていません。');
  }
  return value.trim();
}

function findById_(songs, songId) {
  for (var i = 0; i < songs.length; i++) {
    if (songs[i].song_id === songId) return songs[i];
  }
  return null;
}

function unique_(list) {
  var seen = {};
  return list.filter(function (v) {
    if (seen[v]) return false;
    seen[v] = true;
    return true;
  });
}

/** 公演 ID を Shows の並び順に揃える（マスタに無い ID は末尾に残す）。 */
function sortShowIds_(ids, shows) {
  var rank = {};
  shows.forEach(function (s, i) { rank[s.show_id] = i; });
  return ids.slice().sort(function (a, b) {
    var ra = a in rank ? rank[a] : shows.length;
    var rb = b in rank ? rank[b] : shows.length;
    return ra - rb;
  });
}

/** 2026/1/5 のような表示値を 2026-01-05 に揃える。解釈できなければそのまま返す。 */
function normalizeDate_(text) {
  var m = /^(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})$/.exec(text);
  if (!m) return text;
  return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
}
