import { existsSync, readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { loadEnvFile } from 'node:process';
import { defineConfig } from 'vitest/config';

// Integration settings come from an uncommitted .env file (see .env.example).
// They are read here because the n8n lint rules do not allow process or fs in test files.
const envPath = resolve(import.meta.dirname, '.env');
if (existsSync(envPath)) loadEnvFile(envPath);

const read = (path?: string) =>
	path ? readFileSync(resolve(import.meta.dirname, path), 'utf8') : '';

const integration = process.env.SP_SITE_URL
	? {
			tenantId: process.env.SP_TENANT_ID ?? '',
			clientId: process.env.SP_CLIENT_ID ?? '',
			subdomain: process.env.SP_SUBDOMAIN ?? '',
			privateKey: read(process.env.SP_PRIVATE_KEY_PATH),
			passphrase: process.env.SP_PRIVATE_KEY_PASSPHRASE ?? '',
			certificate: read(process.env.SP_CERTIFICATE_PATH),
			thumbprint: process.env.SP_CERTIFICATE_THUMBPRINT ?? '',
			siteUrl: process.env.SP_SITE_URL,
			libraryPath: process.env.SP_TEST_LIBRARY_PATH ?? 'Shared Documents',
			webhookNotificationUrl: process.env.SP_WEBHOOK_NOTIFICATION_URL ?? '',
			runWriteTests: process.env.SP_RUN_WRITE_TESTS === 'true',
		}
	: null;

export default defineConfig({
	test: {
		include: ['test/integration/**/*.test.ts'],
		testTimeout: 120_000,
		hookTimeout: 300_000,
		fileParallelism: false,
		provide: { integration },
	},
});
