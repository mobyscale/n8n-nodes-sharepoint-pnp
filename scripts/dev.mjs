// `n8n-node dev` starts n8n with npx from this folder, so npx reads `.npmrc` (ignore-scripts=true).
// n8n then gets sqlite3 without its native binary and cannot open its database. Turn install scripts
// back on for that process only.
import { spawn } from 'node:child_process';

const args = process.argv.slice(2).join(' ');
const child = spawn(`n8n-node dev ${args}`.trim(), {
	stdio: 'inherit',
	shell: true,
	env: { ...process.env, npm_config_ignore_scripts: 'false' },
});

child.on('exit', (code) => process.exit(code ?? 0));
