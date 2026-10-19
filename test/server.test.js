const test = require('node:test');
const assert = require('node:assert/strict');
const { bootServer } = require('./helpers');

function setup() {
  const server = bootServer();
  server.call('setupSheets');
  return server;
}

function songCell(ss, songId) {
  const sheet = ss.getSheetByName('Songs');
  const row = sheet.cells.findIndex((r) => r[0] === songId);
  return sheet.cells[row][3];
}

test('setupSheets はダミーデータを作り、2 回目は上書きしない', () => {
  const { call, ss } = bootServer();
  assert.match(call('setupSheets'), /作成しました/);
  ss.getSheetByName('Songs').write(2, 2, '編集済み');
  assert.match(call('setupSheets'), /変更していません/);
  assert.equal(ss.getSheetByName('Songs').cell(2, 2), '編集済み');
  assert.equal(ss.getSheetByName('Songs').formats['2,2'], '@');
});

test('getInitialData は公演を sort_order 順、楽曲の performed_show_ids を配列で返す', () => {
  const { call, ss } = setup();
  ss.getSheetByName('Shows').load([
    ['show_id', 'show_name', 'show_date', 'venue', 'sort_order'],
    ['P002', '愛知', '2026/1/5', '', '2'],
    ['P001', '東京', '', '', '1']
  ]);
  const data = call('getInitialData');
  assert.deepEqual(data.shows.map((s) => s.show_id), ['P001', 'P002']);
  assert.equal(data.shows[1].show_date, '2026-01-05');
  assert.equal(data.songs.length, 6);
  assert.ok(data.songs.every((s) => Array.isArray(s.performed_show_ids) && s.performed_show_ids.length === 0));
});

test('T02/T03: 東京→愛知の順にチェックすると両方保持される', () => {
  const { call, ss } = setup();
  let song = call('setSongPerformance', 'S001', 'P001', true);
  assert.deepEqual(song.performed_show_ids, ['P001']);
  song = call('setSongPerformance', 'S001', 'P002', true);
  assert.deepEqual(song.performed_show_ids, ['P001', 'P002']);
  assert.equal(songCell(ss, 'S001'), '["P001","P002"]');
});

test('T04/T05: チェックを外すとその公演だけ消え、最後の 1 件で未披露に戻る', () => {
  const { call, ss } = setup();
  call('setSongPerformance', 'S001', 'P001', true);
  call('setSongPerformance', 'S001', 'P002', true);
  let song = call('setSongPerformance', 'S001', 'P001', false);
  assert.deepEqual(song.performed_show_ids, ['P002']);
  song = call('setSongPerformance', 'S001', 'P002', false);
  assert.deepEqual(song.performed_show_ids, []);
  assert.equal(songCell(ss, 'S001'), '[]');
});

test('T09: 同じ公演を連続で ON にしても ID は 1 件だけ', () => {
  const { call } = setup();
  call('setSongPerformance', 'S001', 'P001', true);
  const song = call('setSongPerformance', 'S001', 'P001', true);
  assert.deepEqual(song.performed_show_ids, ['P001']);
});

test('T10: 更新はロック下で最新値を読み直すので、別の人の別公演の更新が消えない', () => {
  const { call, ss, lock } = setup();
  // A さんの画面は古いデータのまま。その間に B さんが愛知を追加（シート上で直接起きたのと同じ）
  ss.getSheetByName('Songs').write(2, 4, '["P002"]');
  const song = call('setSongPerformance', 'S001', 'P001', true);
  assert.deepEqual(song.performed_show_ids, ['P001', 'P002']);
  assert.equal(lock.acquired, 1);
  assert.equal(lock.held, false, '処理後にロックが解放されている');
});

test('T10: ロックが取れない場合は更新せずエラー', () => {
  const { call, ss, lock } = setup();
  lock.fail = true;
  assert.throws(() => call('setSongPerformance', 'S001', 'P001', true), /他の人が更新中/);
  assert.equal(songCell(ss, 'S001'), '[]');
});

test('T12: 同じ都市の別日公演は別 ID で区別して保持される', () => {
  const { call } = setup();
  call('setSongPerformance', 'S002', 'P001', true);
  const song = call('setSongPerformance', 'S002', 'P004', true);
  assert.deepEqual(song.performed_show_ids, ['P001', 'P004']);
  const data = call('getInitialData');
  assert.notEqual(data.shows[0].show_id, data.shows[3].show_id);
});

test('T14: 存在しない公演・楽曲 ID、真偽値以外は更新せずエラー', () => {
  const { call, ss } = setup();
  assert.throws(() => call('setSongPerformance', 'S001', 'P999', true), /存在しない公演 ID/);
  assert.throws(() => call('setSongPerformance', 'S999', 'P001', true), /存在しない楽曲 ID/);
  assert.throws(() => call('setSongPerformance', 'S001', 'P001', 'true'), /true \/ false/);
  assert.throws(() => call('setSongPerformance', '', 'P001', true), /songId/);
  assert.equal(songCell(ss, 'S001'), '[]');
});

test('performed_show_ids が壊れている曲は空配列で上書きせずエラー', () => {
  const { call, ss } = setup();
  ss.getSheetByName('Songs').write(2, 4, '["P001",');
  const data = call('getInitialData');
  assert.match(data.songs.find((s) => s.song_id === 'S001').data_error, /JSON/);
  assert.throws(() => call('setSongPerformance', 'S001', 'P002', true), /壊れている/);
  assert.equal(songCell(ss, 'S001'), '["P001",', '破損データはそのまま残る');
});

test('他公演の ID を壊さず、マスタから消えた公演 ID も保持する', () => {
  const { call, ss } = setup();
  ss.getSheetByName('Songs').write(2, 4, '["P900","P003"]');
  const song = call('setSongPerformance', 'S001', 'P001', true);
  assert.deepEqual(song.performed_show_ids, ['P001', 'P003', 'P900']);
});

test('T13: 記号を含む曲名・数式のような曲名も文字列のまま返す', () => {
  const { call, ss } = setup();
  const title = '=HYPERLINK("x") <b>&amp;</b>';
  ss.getSheetByName('Songs').write(8, 1, 'S007');
  ss.getSheetByName('Songs').write(8, 2, title);
  ss.getSheetByName('Songs').write(8, 4, '[]');
  const data = call('getInitialData');
  assert.ok(data.songs.some((s) => s.title === title));
});

test('シートに手で追加した行は、performed_show_ids が空欄でも未披露として扱う', () => {
  const { call, ss } = setup();
  ss.getSheetByName('Songs').write(8, 1, 'S007');
  ss.getSheetByName('Songs').write(8, 2, '手で追加した曲');
  const song = call('getInitialData').songs.find((s) => s.song_id === 'S007');
  assert.deepEqual(song.performed_show_ids, []);
  assert.deepEqual(call('setSongPerformance', 'S007', 'P001', true).performed_show_ids, ['P001']);
});

test('画面から楽曲を追加するサーバー関数は公開していない', () => {
  const { context } = setup();
  assert.equal(typeof context.addSong, 'undefined');
});

test('シートの列順が入れ替わっていてもヘッダー名で読み書きできる', () => {
  const { call, ss } = setup();
  ss.getSheetByName('Songs').load([
    ['performed_show_ids', 'reading', 'title', 'song_id', 'memo'],
    ['[]', 'よみ', '曲', 'S010', 'メモ']
  ]);
  call('setSongPerformance', 'S010', 'P001', true);
  assert.equal(ss.getSheetByName('Songs').cell(2, 1), '["P001"]');
  assert.equal(ss.getSheetByName('Songs').cell(2, 5), 'メモ', '関係ない列は触らない');
});

test('song_id / show_id の重複はエラーとして検出する', () => {
  const { call, ss } = setup();
  ss.getSheetByName('Songs').write(3, 1, 'S001');
  assert.throws(() => call('getInitialData'), /song_id "S001" が重複/);
});
