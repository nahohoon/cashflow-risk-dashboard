import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const p = path.join(__dirname, 'index.html');
let s = fs.readFileSync(p, 'utf8');
const tag = '</html>';
const i = s.indexOf(tag);
if (i < 0) throw new Error('no ' + tag);
const out = s.slice(0, i + tag.length).replace(/\s+$/, '') + '\n';
fs.writeFileSync(p, out, 'utf8');
console.log('wrote', out.length, 'bytes');
