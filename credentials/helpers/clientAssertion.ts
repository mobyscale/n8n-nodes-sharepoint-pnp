import { createHash, createPrivateKey, createSign, randomUUID, X509Certificate } from 'node:crypto';

export interface ClientAssertionOptions {
	tenantId: string;
	clientId: string;
	privateKey: string;
	passphrase?: string;
	certificate?: string;
	thumbprint?: string;
	nowSeconds?: number;
	lifetimeSeconds?: number;
}

const CERTIFICATE_PATTERN = /-----BEGIN CERTIFICATE-----[\s\S]+?-----END CERTIFICATE-----/;

const base64Url = (value: Buffer | string): string => Buffer.from(value).toString('base64url');

export function getTokenEndpoint(tenantId: string): string {
	return `https://login.microsoftonline.com/${encodeURIComponent(tenantId.trim())}/oauth2/v2.0/token`;
}

/**
 * Entra ID identifies the signing certificate by the base64url SHA-1 hash of its DER bytes (x5t).
 * The certificate PEM is preferred because a typed thumbprint is easy to get wrong.
 */
export function getCertificateX5t(
	options: Pick<ClientAssertionOptions, 'certificate' | 'thumbprint' | 'privateKey'>,
): string {
	const certificatePem =
		options.certificate?.match(CERTIFICATE_PATTERN)?.[0] ??
		options.privateKey.match(CERTIFICATE_PATTERN)?.[0];

	if (certificatePem) {
		const der = new X509Certificate(certificatePem).raw;
		return base64Url(createHash('sha1').update(der).digest());
	}

	const hex = (options.thumbprint ?? '').replace(/[^0-9a-fA-F]/g, '');
	if (hex.length !== 40) {
		throw new Error('Provide the certificate PEM or a 40-character SHA-1 certificate thumbprint');
	}
	return base64Url(Buffer.from(hex, 'hex'));
}

export function createClientAssertion(options: ClientAssertionOptions): string {
	const now = options.nowSeconds ?? Math.floor(Date.now() / 1000);
	const lifetime = options.lifetimeSeconds ?? 600;

	const header = { alg: 'RS256', typ: 'JWT', x5t: getCertificateX5t(options) };
	const payload = {
		aud: getTokenEndpoint(options.tenantId),
		iss: options.clientId.trim(),
		sub: options.clientId.trim(),
		jti: randomUUID(),
		nbf: now - 60,
		iat: now,
		exp: now + lifetime,
	};

	const signingInput = `${base64Url(JSON.stringify(header))}.${base64Url(JSON.stringify(payload))}`;
	const key = createPrivateKey({
		key: options.privateKey,
		format: 'pem',
		passphrase: options.passphrase || undefined,
	});
	const signature = createSign('RSA-SHA256').update(signingInput).sign(key);

	return `${signingInput}.${base64Url(signature)}`;
}
