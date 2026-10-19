const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { createGasGlobals } = require('./gas-mock');

const SRC = path.join(__dirname, '..', 'src');

/** Code.gs をモック付きの VM で読み込む */
function bootServer() {
  const gas = createGasGlobals();
  const context = vm.createContext({ ...gas.globals });
  vm.runInContext(fs.readFileSync(path.join(SRC, 'Code.gs'), 'utf8'), context, { filename: 'Code.gs' });
  // VM 内のオブジェクトは別 realm なので、JSON 経由で素のオブジェクトにして比較しやすくする
  const call = (name, ...args) => JSON.parse(JSON.stringify(context[name](...args)));
  return { context, call, ss: gas.spreadsheet, lock: gas.lock };
}

/** Logic.html の <script> 部分を読み込む */
function loadLogic() {
  const html = fs.readFileSync(path.join(SRC, 'Logic.html'), 'utf8');
  const code = html.replace(/^\s*<script>/, '').replace(/<\/script>\s*$/, '');
  const context = vm.createContext({});
  vm.runInContext(code, context, { filename: 'Logic.html' });
  return context.SongLogic;
}

module.exports = { bootServer, loadLogic };
