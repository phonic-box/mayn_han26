/**
 * May'n ひらがな はんぐじゃむ 披露楽曲管理 — サーバー側（Google Apps Script）
 *
 * データは Songs / Shows / Setlists の 3 シート。管理者がスプレッドシートで登録し、
 * Web アプリは閲覧専用（データを書き換える関数は公開しない）。
 * 曲ごとの「披露済み／未披露」は Setlists から都度算出し、シートには保存しない。
 */

var APP_TITLE = "May'n ひらがな はんぐじゃむ｜楽曲管理";
var SONGS_SHEET = 'Songs';
var SHOWS_SHEET = 'Shows';
var SETLISTS_SHEET = 'Setlists';
var SONG_HEADERS = ['song_id', 'title', 'reading'];
var SHOW_HEADERS = ['show_id', 'show_name', 'show_date', 'venue', 'sort_order'];
var SETLIST_HEADERS = ['show', 'no', 'title'];
var MIGRATED_PROPERTY = 'CHECKS_MIGRATED';

// 説明・動作確認用のダミーデータ。実データではない。
var SAMPLE_SHOWS = [
  ['P001', '東京公演（サンプル）', '', '', 1],
  ['P002', '愛知公演（サンプル）', '', '', 2],
  ['P003', '京都公演（サンプル）', '', '', 3],
  ['P004', '東京公演 2日目（サンプル）', '', '', 4]
];
var SAMPLE_SONGS = [
  ['S001', '楽曲A', 'がっきょくえー'],
  ['S002', '楽曲B', 'がっきょくびー'],
  ['S003', '楽曲C', 'がっきょくしー'],
  ['S004', 'あいうえおの歌（サンプル）', 'あいうえおのうた'],
  ['S005', 'Sample <Song> & "記号"', 'さんぷるそんぐ'],
  ['S006', 'よみ未登録のサンプル曲', '']
];
var SAMPLE_SETLISTS = [
  ['東京公演（サンプル）', 1, '楽曲C'],
  ['東京公演（サンプル）', 2, '楽曲A'],
  ['愛知公演（サンプル）', 1, '楽曲A']
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
    .addItem('シートを初期化・プルダウンを更新', 'setupSheets')
    .addItem('旧チェックを Setlists に移す（1 回だけ）', 'migrateChecksToSetlists')
    .addToUi();
}

// ---------------------------------------------------------------------------
// 画面から呼ぶ関数（google.script.run）— 読み取りのみ
// ---------------------------------------------------------------------------

/**
 * @return {{
 *   shows: Object[],
 *   songs: Object[],              // performed_show_ids は Setlists から算出
 *   setlists: Object<string, Object[]>,  // show_id → 曲順に並べた [{no, title, reading, song_id}]
 *   warnings: string[],
 *   fetchedAt: string
 * }}
 */
function getInitialData() {
  var ss = getSpreadsheet_();
  var shows = readShows_(ss);
  var songs = readSongs_(ss).songs;
  var result = readSetlists_(ss, shows, songs);

  var performed = {};
  shows.forEach(function (show) {
    result.setlists[show.show_id].forEach(function (entry) {
      if (!entry.song_id) return;
      var ids = performed[entry.song_id] || (performed[entry.song_id] = []);
      if (ids.indexOf(show.show_id) === -1) ids.push(show.show_id);
    });
  });

  return {
    shows: shows,
    songs: songs.map(function (song) {
      return {
        song_id: song.song_id,
        title: song.title,
        reading: song.reading,
        performed_show_ids: performed[song.song_id] || []
      };
    }),
    setlists: result.setlists,
    warnings: result.warnings,
    fetchedAt: new Date().toISOString()
  };
}

// ---------------------------------------------------------------------------
// 管理者用（エディタまたはスプレッドシートのメニューから手動実行）
// ---------------------------------------------------------------------------

/**
 * シートが無ければヘッダー付きで作成し、Setlists のプルダウン（入力規則）を設定する。
 * 既存データは上書きしない。ダミーデータは Songs / Shows を新規作成したときだけ入れる。
 */
function setupSheets() {
  var ss = getSpreadsheet_();
  var created = [];
  var fresh = false;
  if (ensureSheet_(ss, SHOWS_SHEET, SHOW_HEADERS, SAMPLE_SHOWS, 4)) created.push(SHOWS_SHEET);
  if (ensureSheet_(ss, SONGS_SHEET, SONG_HEADERS, SAMPLE_SONGS, 3)) {
    created.push(SONGS_SHEET);
    fresh = true;
  }
  if (ensureSheet_(ss, SETLISTS_SHEET, SETLIST_HEADERS, fresh ? SAMPLE_SETLISTS : [], 1)) {
    created.push(SETLISTS_SHEET);
  }
  applySetlistValidation_(ss);
  return (created.length ? created.join(', ') + ' シートを作成しました。' : 'シートは作成済みです。') +
    'Setlists のプルダウンを更新しました。';
}

/**
 * 画面でチェックしていた頃の Songs.performed_show_ids を Setlists に移す（曲順 no は空欄）。
 * 2 回目以降は何もしない。移した後は Songs の performed_show_ids 列を削除してよい。
 */
function migrateChecksToSetlists() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty(MIGRATED_PROPERTY)) {
    return '移行は実行済みのため、何もしていません。';
  }

  var ss = getSpreadsheet_();
  setupSheets();
  var shows = readShows_(ss);
  var songsTable = readTable_(ss, SONGS_SHEET, SONG_HEADERS);
  if (!songsTable.columns.performed_show_ids) {
    props.setProperty(MIGRATED_PROPERTY, new Date().toISOString());
    return 'Songs に performed_show_ids 列が無いため、移すものはありません。';
  }

  var existing = {};
  readTable_(ss, SETLISTS_SHEET, SETLIST_HEADERS).rows.forEach(function (row) {
    existing[row.get('show') + '\n' + row.get('title')] = true;
  });

  var rows = [];
  shows.forEach(function (show) {
    songsTable.rows.forEach(function (row) {
      var ids;
      try {
        ids = JSON.parse(row.get('performed_show_ids') || '[]');
      } catch (e) {
        throw new Error('Songs ' + row.rowNumber + ' 行目の performed_show_ids が読めません。直してから再実行してください。');
      }
      var title = row.get('title');
      if (!Array.isArray(ids) || ids.indexOf(show.show_id) === -1 || !title) return;
      if (existing[show.show_name + '\n' + title] || existing[show.show_id + '\n' + title]) return;
      rows.push([show.show_name, '', title]);
    });
  });

  if (rows.length) {
    var sheet = ss.getSheetByName(SETLISTS_SHEET);
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, SETLIST_HEADERS.length).setValues(rows);
  }
  props.setProperty(MIGRATED_PROPERTY, new Date().toISOString());
  return rows.length + ' 件を Setlists に移しました。曲順（no）を入力してください。';
}

function ensureSheet_(ss, name, headers, sampleRows, textColumnCount) {
  var sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sheet.getLastRow() > 0) return false;

  sheet.getRange(1, 1, sheet.getMaxRows(), textColumnCount).setNumberFormat('@');
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (sampleRows.length) {
    sheet.getRange(2, 1, sampleRows.length, headers.length).setValues(sampleRows);
  }
  sheet.setFrozenRows(1);
  return true;
}

/** Setlists の show 列は Shows.show_name、title 列は Songs.title から選べるようにする。 */
function applySetlistValidation_(ss) {
  var setlists = readTable_(ss, SETLISTS_SHEET, SETLIST_HEADERS);
  var shows = readTable_(ss, SHOWS_SHEET, SHOW_HEADERS);
  var songs = readTable_(ss, SONGS_SHEET, SONG_HEADERS);
  var rowCount = setlists.sheet.getMaxRows() - 1;

  var pairs = [
    [setlists.columns.show, shows.sheet, shows.columns.show_name],
    [setlists.columns.title, songs.sheet, songs.columns.title]
  ];
  pairs.forEach(function (p) {
    var source = p[1].getRange(2, p[2], p[1].getMaxRows() - 1, 1);
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(source, true)
      .setAllowInvalid(false)
      .build();
    setlists.sheet.getRange(2, p[0], rowCount, 1).setDataValidation(rule);
  });
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
        return function (key) { return columns[key] ? String(cells[columns[key] - 1]).trim() : ''; };
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
    songs.push({ song_id: id, title: row.get('title') || id, reading: row.get('reading') });
  });
  return { sheet: table.sheet, songs: songs };
}

/**
 * Setlists を公演ごとに曲順で並べる。show は show_name か show_id、title は曲名か song_id で照合する。
 * 照合できない行は warnings に入れる（公演不明の行は捨て、曲不明の行は曲名だけで表示する）。
 */
function readSetlists_(ss, shows, songs) {
  var table = readTable_(ss, SETLISTS_SHEET, SETLIST_HEADERS);
  var showByKey = {};
  shows.forEach(function (s) {
    showByKey[s.show_name] = s;
    showByKey[s.show_id] = s;
  });
  var songByKey = {};
  songs.forEach(function (s) {
    songByKey[s.title] = s;
    songByKey[s.song_id] = s;
  });

  var setlists = {};
  shows.forEach(function (s) { setlists[s.show_id] = []; });
  var warnings = [];

  table.rows.forEach(function (row) {
    var where = SETLISTS_SHEET + ' ' + row.rowNumber + ' 行目';
    var showKey = row.get('show');
    var title = row.get('title');
    if (!showKey && !title) return;
    var show = showByKey[showKey];
    if (!show) {
      warnings.push(where + ': 公演「' + showKey + '」が Shows にありません。');
      return;
    }
    if (!title) {
      warnings.push(where + ': 曲名が空です。');
      return;
    }
    var song = songByKey[title];
    if (!song) warnings.push(where + ': 曲名「' + title + '」が Songs にありません。');
    var no = row.get('no');
    setlists[show.show_id].push({
      no: no,
      title: song ? song.title : title,
      reading: song ? song.reading : '',
      song_id: song ? song.song_id : '',
      _order: no !== '' && isFinite(Number(no)) ? Number(no) : Number.MAX_SAFE_INTEGER,
      _row: row.rowNumber
    });
  });

  Object.keys(setlists).forEach(function (id) {
    setlists[id] = setlists[id]
      .sort(function (a, b) { return a._order - b._order || a._row - b._row; })
      .map(function (e) { return { no: e.no, title: e.title, reading: e.reading, song_id: e.song_id }; });
  });
  return { setlists: setlists, warnings: warnings };
}

/** 2026/1/5 のような表示値を 2026-01-05 に揃える。解釈できなければそのまま返す。 */
function normalizeDate_(text) {
  var m = /^(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})$/.exec(text);
  if (!m) return text;
  return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
}
