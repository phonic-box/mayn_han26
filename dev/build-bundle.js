/**
 * Apps Script エディタに手で貼るための 2 ファイル版を dist/gas/ に生成する。
 *   dist/gas/Code.gs     … src/Code.gs と同じ
 *   dist/gas/Index.html  … Styles / Logic / Scripts を埋め込み済み
 */
const fs = require('node:fs');
const path = require('node:path');

const root = path.join(__dirname, '..');
const src = (name) => fs.readFileSync(path.join(root, 'src', name), 'utf8');
const out = path.join(root, 'dist', 'gas');

const index = src('Index.html')
  .replace(/<\?!=\s*include\('(\w+)'\);\s*\?>/g, (_, name) => src(name + '.html').trim());
if (/<\?/.test(index.replace(/<\?xml/g, ''))) {
  throw new Error('Index.html にテンプレート記法が残っています');
}

fs.mkdirSync(out, { recursive: true });
fs.writeFileSync(path.join(out, 'Code.gs'), src('Code.gs'));
fs.writeFileSync(path.join(out, 'Index.html'), index);
console.log('dist/gas/Code.gs と dist/gas/Index.html を生成しました');
