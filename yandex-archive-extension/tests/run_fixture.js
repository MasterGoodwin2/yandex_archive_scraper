// Самопроверка: node tests/run_fixture.js
// Прогоняет страницу 165 (ЦГА Москвы, ф. 51, оп. 8, д. 588) через revision-rows.js и layout.js.
const fs = require('fs'), path = require('path');
const root = path.join(__dirname, '..');
const Logic = require(path.join(root, 'logic.js'));
const Layout = require(path.join(root, 'layout.js'));
const fx = JSON.parse(fs.readFileSync(path.join(__dirname, 'page165_fixture.json'), 'utf8'));
const st = Layout.newState(), fams = [];
const how = Layout.processNode(fx.props.pageProps.currentNode, { url: 'test', page: 165 }, st, 'revision', fams);
Layout.finish(st, fams);
const ok = [
  ['метка режима', /__NEXT_DATA__/.test(how)],
  ['2 двора (23 и 24)', fams.map((f) => f.yard).join(',') === '23,24'],
  ['населённый пункт', fams.every((f) => f.settlement === 'деревня Черепкова')],
  ['28 строк таблицы', st.rows.length === 28],
  ['блоков в дворах: 3 и 4', fams.map((f) => f.persons.length).join(',') === '3,4'],
  ['итоги страницы сошлись', /Муж.*сошлось.*Жен.*сошлось/.test(st.checks[165])],
];
ok.forEach(([n, v]) => console.log((v ? 'OK   ' : 'FAIL ') + n));
console.log('\n' + Layout.familyCell(fams[0]).split('\n').slice(0, 5).join('\n') + '\n...');
process.exit(ok.every(([, v]) => v) ? 0 : 1);
