/**
 * GAS にデプロイせずに画面を確認するためのプレビューを dist/preview.html に生成する。
 * Code.gs をブラウザ内のモック（test/gas-mock.js）上で動かし、google.script.run を模倣する。
 *
 *   ?fail=1   すべての保存を失敗させる（T11 の確認用）
 *   ?delay=ms サーバー応答の遅延（既定 600ms）
 */
const fs = require('node:fs');
const path = require('node:path');

const root = path.join(__dirname, '..');
const src = (name) => fs.readFileSync(path.join(root, 'src', name), 'utf8');

const shim = `
<script>${fs.readFileSync(path.join(root, 'test', 'gas-mock.js'), 'utf8')}</script>
<script>
  (function () {
    var gas = GasMock.createGasGlobals();
    Object.assign(window, gas.globals);
  })();
</script>
<script>${src('Code.gs')}</script>
<script>
  (function () {
    var params = new URLSearchParams(location.search);
    var delay = Number(params.get('delay') || 600);
    var fail = params.get('fail') === '1';
    setupSheets();
    function makeRunner(success, failure) {
      return new Proxy({}, {
        get: function (_, name) {
          if (name === 'withSuccessHandler') return function (fn) { return makeRunner(fn, failure); };
          if (name === 'withFailureHandler') return function (fn) { return makeRunner(success, fn); };
          return function () {
            var args = JSON.parse(JSON.stringify(Array.prototype.slice.call(arguments)));
            setTimeout(function () {
              try {
                if (fail && name !== 'getInitialData') throw new Error('（プレビュー）通信に失敗しました');
                var result = JSON.parse(JSON.stringify(window[name].apply(null, args)));
                success && success(result);
              } catch (e) {
                failure && failure(e);
              }
            }, delay);
          };
        }
      });
    }
    window.google = { script: { run: makeRunner(null, null) } };
  })();
</script>`;

const html = src('Index.html')
  .replace('<head>', '<head>\n' + shim)
  .replace(/<\?!=\s*include\('(\w+)'\);\s*\?>/g, (_, name) => src(name + '.html'));

fs.mkdirSync(path.join(root, 'dist'), { recursive: true });
fs.writeFileSync(path.join(root, 'dist', 'preview.html'), html);
console.log('dist/preview.html を生成しました → http://localhost:8765/preview.html');
