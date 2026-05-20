const fs = require('fs');
const path = require('path');
const p = path.join(__dirname, '..', 'index.html');
let s = fs.readFileSync(p, 'utf8');
const marker = '__HTML_DUP_RM__';
const i = s.indexOf(marker);
if (i >= 0) {
  s = s.slice(0, i).trimEnd() + '\n';
}
s = s.replace(/\n>>>>>>>[^\n]+\n?/g, '\n');
fs.writeFileSync(p, s);
console.log('trimmed index.html');
