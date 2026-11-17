const fs = require('node:fs');
const path = require('node:path');
const esbuild = require(require.resolve('esbuild', {paths: [path.dirname(require.resolve('tsx/package.json'))]}));
const root = path.resolve(__dirname, '..');
const out = path.join(root, 'out', 'blindtype-setup');
async function build() {
    fs.mkdirSync(out, {recursive: true});
    await esbuild.build({absWorkingDir: root, entryPoints: ['./scripts/blindtype-setup-entry.tsx'], outfile: path.join(out, 'cm_blindtypesetup.js'),
        bundle: true, minify: true, platform: 'browser', target: 'es2020', define: {'process.env.NODE_ENV': '"production"'}, logLevel: 'info'});
    fs.copyFileSync(path.join(root, 'assets', 'blindtype-setup.html'), path.join(out, 'cm_blindtypesetup.html'));
    console.log(`Blind type setup resources: ${out}`);
}
build().catch(error => {console.error(error); process.exitCode = 1;});
