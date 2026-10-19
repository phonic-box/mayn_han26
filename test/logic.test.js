const test = require('node:test');
const assert = require('node:assert/strict');
const { loadLogic } = require('./helpers');

const L = loadLogic();

const songs = [
  { song_id: 'S001', title: '楽曲C', reading: 'がっきょくしー', performed_show_ids: ['P001', 'P002'] },
  { song_id: 'S002', title: '楽曲A', reading: 'がっきょくえー', performed_show_ids: ['P002'] },
  { song_id: 'S003', title: 'Zebra', reading: '', performed_show_ids: [] },
  { song_id: 'S004', title: 'あいうえお', reading: 'アイウエオ', performed_show_ids: [] }
];
const ids = (view) => view.items.map((s) => s.song_id);
const build = (state) => L.buildView(songs, { selection: L.ALL, query: '', filter: 'all', ...state });

test('reading の 50 音順（カタカナもひらがな扱い）、無ければ title で並ぶ', () => {
  assert.deepEqual(ids(build({})), ['S003', 'S004', 'S002', 'S001']);
});

test('T01: 全曲未披露なら、全公演の未披露に全曲が出る', () => {
  const fresh = songs.map((s) => ({ ...s, performed_show_ids: [] }));
  const view = L.buildView(fresh, { selection: L.ALL, query: '', filter: 'unperformed' });
  assert.equal(view.items.length, 4);
});

test('T06: 全公演ではどれかの公演で歌われた曲が披露済み', () => {
  const view = build({ filter: 'performed' });
  assert.deepEqual(ids(view), ['S002', 'S001']);
});

test('T07: 全公演の未披露はどの公演でも歌われていない曲だけ', () => {
  assert.deepEqual(ids(build({ filter: 'unperformed' })), ['S003', 'S004']);
});

test('個別公演では、その公演に対して披露済み／未披露を判定する', () => {
  assert.deepEqual(ids(build({ selection: 'P001', filter: 'performed' })), ['S001']);
  assert.deepEqual(ids(build({ selection: 'P001', filter: 'unperformed' })), ['S003', 'S004', 'S002']);
});

test('keepIds の曲は絞り込みに合わなくても残る（保存中の曲）', () => {
  assert.deepEqual(ids(build({ selection: 'P001', filter: 'unperformed', keepIds: ['S001'] })), ['S003', 'S004', 'S002', 'S001']);
  assert.deepEqual(ids(build({ selection: 'P001', filter: 'unperformed', keepIds: ['S001'], query: 'Zeb' })), ['S003']);
});

test('T08: 曲名・読みの部分一致検索（全角半角・カタカナを吸収）', () => {
  assert.deepEqual(ids(build({ query: '楽曲' })), ['S002', 'S001']);
  assert.deepEqual(ids(build({ query: 'ガッキョクエ' })), ['S002']);
  assert.deepEqual(ids(build({ query: 'ｚｅｂ' })), ['S003']);
  assert.deepEqual(ids(build({ query: 'いうえ' })), ['S004']);
  assert.deepEqual(ids(build({ query: '存在しない' })), []);
});

test('検索・絞り込みで元データは変わらない', () => {
  const before = JSON.stringify(songs);
  build({ selection: 'P001', query: '楽', filter: 'unperformed' });
  assert.equal(JSON.stringify(songs), before);
});
