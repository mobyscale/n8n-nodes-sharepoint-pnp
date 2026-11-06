// Bundles each node entry point with its PnPjs imports into a single CommonJS file.
// PnPjs is ESM-only and n8n Cloud does not allow runtime dependencies, so it must be inlined.
import { build } from 'esbuild';
import { readFileSync } from 'node:fs';

const pkg = JSON.parse(readFileSync(new URL('../package.json', import.meta.url), 'utf8'));

const entryPoints = pkg.n8n.nodes.map((output) => ({
	in: output.replace(/^dist\//, '').replace(/\.js$/, '.ts'),
	out: output.replace(/^dist\//, '').replace(/\.js$/, ''),
}));

await build({
	entryPoints,
	outdir: 'dist',
	bundle: true,
	platform: 'node',
	format: 'cjs',
	target: 'node20',
	external: ['n8n-workflow'],
	sourcemap: true,
	legalComments: 'linked',
	logLevel: 'info',
});
