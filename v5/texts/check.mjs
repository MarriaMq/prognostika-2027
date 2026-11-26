// Проверка текстов: node v5/texts/check.mjs
import { readFileSync, existsSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
const root = dirname(fileURLToPath(import.meta.url));
const SECT = { ru: ['Общий фон', 'Карьера и бизнес', 'Финансы', 'Риски и сложности'],
               uk: ['Головні тенденції року', 'Кар’єра та бізнес', 'Фінанси', 'Ризики та труднощі'] };
let bad = 0, ok = 0;
for (const lang of ['ru', 'uk']) for (const g of ['female', 'male']) for (let a = 1; a <= 22; a++) {
  const f = join(root, 'year/general', lang, g, String(a).padStart(2, '0') + '.md');
  if (!existsSync(f)) { console.log('НЕТ ФАЙЛА', f); bad++; continue; }
  const t = readFileSync(f, 'utf8');
  const heads = [...t.matchAll(/^## (.+)$/gm)].map(m => m[1].trim());
  const errs = [];
  if (!new RegExp(`^# ${a} аркан`).test(t)) errs.push('заголовок не совпадает с номером файла');
  if (heads.join('|') !== SECT[lang].join('|')) errs.push('рубрики: ' + heads.join(', '));
  if ((t.match(/\*\*/g) || []).length % 2) errs.push('незакрытые ** (жирный)');
  t.split(/^## .+$/m).slice(1).forEach((s, i) => { if (!s.trim()) errs.push('пустая рубрика ' + heads[i]); });
  if (errs.length) { console.log(f, '\n  ' + errs.join('\n  ')); bad++; } else ok++;
}
console.log(`Готово: ${ok} в порядке, ${bad} с ошибками`);
process.exit(bad ? 1 : 0);
