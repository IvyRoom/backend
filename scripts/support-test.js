'use strict';

const fs = require('node:fs');
const path = require('node:path');
const { spawn } = require('node:child_process');

const root = path.resolve(__dirname, '..');
const tests = fs.readdirSync(path.join(root, 'test'))
    .filter((name) => /^support-.*\.test\.js$/u.test(name) || name === 'pagarme.test.js')
    .sort()
    .map((name) => path.join(root, 'test', name));
const frontendTests = path.resolve(root, '../sistemas/.agents/tests');
tests.push(...fs.readdirSync(frontendTests)
    .filter((name) => /^support-.*\.test\.js$/u.test(name))
    .sort()
    .map((name) => path.join(frontendTests, name)));

const child = spawn(process.execPath, [
    '--require', path.join(root, 'test/deny-production-network.js'), '--test', ...tests,
], { cwd: root, stdio: 'inherit', windowsHide: true });
child.once('error', () => {
    console.error('Support checks could not start. Check Node.js and the sibling sistemas checkout.');
    process.exitCode = 1;
});
child.once('exit', (code) => { process.exitCode = code ?? 1; });
