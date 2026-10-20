const test = require('node:test');
const assert = require('node:assert/strict');
const { bootServer } = require('./helpers');

const SONGS = [
  ['song_id', 'title', 'reading'],
  ['S001', '楽曲A', 'がっきょくえー'],
  ['S002', '楽曲B', ''],
  ['S003', '楽曲C', '']
];
const SHOWS = [
  ['show_id', 'show_name', 'show_date', 'venue', 'sort_order'],
  ['P002', '愛知', '2026/1/5', '', '2'],
  ['P001', '東京 1回目', '', '', '1'],
  ['P003', '東京 2回目', '', '', '3']
];

/** Songs / Shows を指定内容にし、Setlists を rows で埋める */
function boot(setlistRows) {
  const server = bootServer();
  server.call('setupSheets');
  server.ss.getSheetByName('Songs').load(SONGS);
  server.ss.getSheetByName('Shows').load(SHOWS);
  server.ss.getSheetByName('Setlists').load([['show', 'no', 'title'], ...(setlistRows || [])]);
  return server;
}

test('setupSheets: 新規なら 3 シートをダミーデータ付きで作り、既存データは上書きしない', () => {
  const { call, ss } = bootServer();
  assert.match(call('setupSheets'), /Shows, Songs, Setlists シートを作成しました/);
  const data = call('getInitialData');
  assert.equal(data.songs.length, 6);
  assert.deepEqual(data.setlists.P001.map((e) => e.title), ['楽曲C', '楽曲A']);

  ss.getSheetByName('Songs').write(2, 2, '編集済み');
  assert.match(call('setupSheets'), /シートは作成済み/);
  assert.equal(ss.getSheetByName('Songs').cell(2, 2), '編集済み');
});

test('setupSheets: 既存の Songs / Shows に Setlists を足すときはダミーを入れない', () => {
  const { call, ss } = bootServer();
  ss.insertSheet('Songs').load(SONGS);
  ss.insertSheet('Shows').load(SHOWS);
  call('setupSheets');
  assert.deepEqual(ss.getSheetByName('Setlists').cells, [['show', 'no', 'title']]);
});

test('setupSheets: Setlists の show / title 列に Shows.show_name / Songs.title のプルダウンを付ける', () => {
  const { ss } = boot();
  const v = ss.getSheetByName('Setlists').validations.slice(-2);
  assert.equal(v[0].col, 1);
  assert.equal(v[0].rule.source.sheet.getName(), 'Shows');
  assert.equal(v[0].rule.source.col, 2);
  assert.equal(v[1].col, 3);
  assert.equal(v[1].rule.source.sheet.getName(), 'Songs');
  assert.equal(v[1].rule.source.col, 2);
  assert.equal(v[1].rule.allowInvalid, false);
});

test('セットリストは no 順、no が空の行は後ろに入力順で並ぶ', () => {
  const { call } = boot([
    ['東京 1回目', '2', '楽曲A'],
    ['東京 1回目', '', '楽曲C'],
    ['東京 1回目', '1', '楽曲B']
  ]);
  const list = call('getInitialData').setlists.P001;
  assert.deepEqual(list.map((e) => [e.no, e.title]), [['1', '楽曲B'], ['2', '楽曲A'], ['', '楽曲C']]);
  assert.equal(list[1].reading, 'がっきょくえー');
});

test('披露済みはセットリストから算出し、公演は Shows の並び順', () => {
  const { call } = boot([
    ['愛知', '1', '楽曲A'],
    ['東京 1回目', '1', '楽曲A'],
    ['東京 1回目', '2', '楽曲A']
  ]);
  const data = call('getInitialData');
  assert.deepEqual(data.shows.map((s) => s.show_id), ['P001', 'P002', 'P003']);
  assert.equal(data.shows[1].show_date, '2026-01-05');
  const byId = Object.fromEntries(data.songs.map((s) => [s.song_id, s.performed_show_ids]));
  assert.deepEqual(byId.S001, ['P001', 'P002'], '同じ公演で 2 回歌っても公演 ID は 1 件');
  assert.deepEqual(byId.S002, []);
  assert.deepEqual(data.setlists.P003, []);
});

test('同じ都市の別公演は区別される', () => {
  const { call } = boot([['東京 2回目', '1', '楽曲B']]);
  const data = call('getInitialData');
  assert.deepEqual(data.setlists.P001, []);
  assert.deepEqual(data.setlists.P003.map((e) => e.title), ['楽曲B']);
});

test('show は show_id、title は song_id でも照合できる', () => {
  const { call } = boot([['P002', '1', 'S003']]);
  const data = call('getInitialData');
  assert.deepEqual(data.setlists.P002.map((e) => e.title), ['楽曲C']);
  assert.deepEqual(data.songs.find((s) => s.song_id === 'S003').performed_show_ids, ['P002']);
});

test('照合できない行は警告し、公演不明は捨て、曲不明は曲名だけ表示する', () => {
  const { call } = boot([
    ['大阪', '1', '楽曲A'],
    ['東京 1回目', '1', '未登録の曲'],
    ['東京 1回目', '2', ''],
    ['', '', '']
  ]);
  const data = call('getInitialData');
  assert.deepEqual(data.warnings, [
    'Setlists 2 行目: 公演「大阪」が Shows にありません。',
    'Setlists 3 行目: 曲名「未登録の曲」が Songs にありません。',
    'Setlists 4 行目: 曲名が空です。'
  ]);
  assert.deepEqual(data.setlists.P001, [{ no: '1', title: '未登録の曲', reading: '', song_id: '' }]);
  assert.ok(data.songs.every((s) => s.performed_show_ids.length === 0));
});

test('記号を含む曲名・数式のような曲名も文字列のまま返す', () => {
  const title = '=HYPERLINK("x") <b>&amp;</b>';
  const { call, ss } = boot([['東京 1回目', '1', title]]);
  ss.getSheetByName('Songs').write(5, 1, 'S004');
  ss.getSheetByName('Songs').write(5, 2, title);
  const data = call('getInitialData');
  assert.equal(data.setlists.P001[0].title, title);
  assert.equal(data.setlists.P001[0].song_id, 'S004');
});

test('Songs に旧 performed_show_ids 列やメモ列が残っていても読める', () => {
  const { call, ss } = boot([['東京 1回目', '1', '楽曲B']]);
  ss.getSheetByName('Songs').load([
    ['memo', 'performed_show_ids', 'reading', 'title', 'song_id'],
    ['x', '["P002"]', '', '楽曲B', 'S002']
  ]);
  const data = call('getInitialData');
  assert.deepEqual(data.songs, [{ song_id: 'S002', title: '楽曲B', reading: '', performed_show_ids: ['P001'] }]);
});

test('migrateChecksToSetlists: 旧チェックを公演順・曲順で移し、既存行は重複させず、2 回目は何もしない', () => {
  const { call, ss } = boot([['東京 1回目', '1', '楽曲C']]);
  ss.getSheetByName('Songs').load([
    ['song_id', 'title', 'reading', 'performed_show_ids'],
    ['S001', '楽曲A', '', '["P002","P001"]'],
    ['S002', '楽曲B', '', '[]'],
    ['S003', '楽曲C', '', '["P001"]']
  ]);
  assert.match(call('migrateChecksToSetlists'), /2 件を Setlists に移しました/);
  assert.deepEqual(ss.getSheetByName('Setlists').cells.slice(1), [
    ['東京 1回目', '1', '楽曲C'],
    ['東京 1回目', '', '楽曲A'],
    ['愛知', '', '楽曲A']
  ]);
  assert.match(call('migrateChecksToSetlists'), /実行済み/);
  assert.equal(ss.getSheetByName('Setlists').cells.length, 4);
});

test('migrateChecksToSetlists: performed_show_ids が壊れていたら何も書かずに止まる', () => {
  const { call, ss, props } = boot();
  ss.getSheetByName('Songs').load([
    ['song_id', 'title', 'reading', 'performed_show_ids'],
    ['S001', '楽曲A', '', '["P001",']
  ]);
  assert.throws(() => call('migrateChecksToSetlists'), /2 行目の performed_show_ids が読めません/);
  assert.equal(ss.getSheetByName('Setlists').cells.length, 1);
  assert.equal(props.CHECKS_MIGRATED, undefined, '再実行できる');
});

test('画面からデータを書き換える関数は公開していない', () => {
  const { context } = boot();
  assert.equal(typeof context.setSongPerformance, 'undefined');
  assert.equal(typeof context.addSong, 'undefined');
});

test('song_id / show_id の重複はエラーとして検出する', () => {
  const { call, ss } = boot();
  ss.getSheetByName('Songs').write(3, 1, 'S001');
  assert.throws(() => call('getInitialData'), /song_id "S001" が重複/);
});
