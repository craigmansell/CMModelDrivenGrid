import {mkdirSync, writeFileSync} from 'node:fs';
import {resolve} from 'node:path';
import {initialR200sSetup} from './blindtype-setup-model';
const output = resolve(__dirname, '../out/blindtype-setup');
mkdirSync(output, {recursive: true});
const seed = initialR200sSetup();
for (const name of ['catalogue', 'document'] as const) writeFileSync(resolve(output, `${name}.json`), JSON.stringify(seed[name], null, 2));
console.log('Exported validated R200S catalogue and setup seed.');
