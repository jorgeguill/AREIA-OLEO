// Valida a sintaxe do JS inline de páginas HTML do Painel do Forno.
// Uso: node scripts/check.js painel-forno-COMPLETO-SYNC/*.html
const fs = require('fs');
const vm = require('vm');
let anyBad = false;
for (const f of process.argv.slice(2)) {
  const html = fs.readFileSync(f, 'utf8');
  const re = /<script>([\s\S]*?)<\/script>/g;
  let m, n = 0, bad = 0;
  while ((m = re.exec(html)) !== null) {
    n++;
    try { new vm.Script(m[1]); }
    catch (e) { bad++; anyBad = true; console.log(`FAIL ${f} script#${n}: ${e.message}`); }
  }
  console.log(`${f}: ${n} inline script(s), ${bad} com erro de sintaxe`);
}
process.exit(anyBad ? 1 : 0);
