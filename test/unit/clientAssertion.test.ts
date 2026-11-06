import { createVerify, generateKeyPairSync } from 'node:crypto';
import { describe, expect, it } from 'vitest';

import {
	createClientAssertion,
	getCertificateX5t,
	getTokenEndpoint,
} from '../../credentials/helpers/clientAssertion';

const { privateKey, publicKey } = generateKeyPairSync('rsa', {
	modulusLength: 2048,
	privateKeyEncoding: { type: 'pkcs8', format: 'pem' },
	publicKeyEncoding: { type: 'spki', format: 'pem' },
});

const thumbprint = 'A1B2C3D4E5F60718293A4B5C6D7E8F9012345678';
const decode = (part: string) => JSON.parse(Buffer.from(part, 'base64url').toString('utf8'));

describe('client assertion', () => {
	it('converts a hex thumbprint to x5t', () => {
		expect(
			getCertificateX5t({ privateKey, thumbprint: 'a1:b2 c3d4e5f60718293a4b5c6d7e8f9012345678' }),
		).toBe(Buffer.from(thumbprint, 'hex').toString('base64url'));
	});

	it('requires a certificate or a valid thumbprint', () => {
		expect(() => getCertificateX5t({ privateKey, thumbprint: 'abc' })).toThrow(/thumbprint/);
	});

	it('creates a signed JWT for the Entra token endpoint', () => {
		const jwt = createClientAssertion({
			tenantId: 'tenant-id',
			clientId: 'client-id',
			privateKey,
			thumbprint,
			nowSeconds: 1_700_000_000,
		});
		const [header, payload, signature] = jwt.split('.');

		expect(decode(header)).toEqual({
			alg: 'RS256',
			typ: 'JWT',
			x5t: Buffer.from(thumbprint, 'hex').toString('base64url'),
		});
		expect(decode(payload)).toMatchObject({
			aud: getTokenEndpoint('tenant-id'),
			iss: 'client-id',
			sub: 'client-id',
			nbf: 1_700_000_000 - 60,
			exp: 1_700_000_000 + 600,
		});

		const valid = createVerify('RSA-SHA256')
			.update(`${header}.${payload}`)
			.verify(publicKey, Buffer.from(signature, 'base64url'));
		expect(valid).toBe(true);
	});
});
