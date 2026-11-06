import type {
	IAuthenticateGeneric,
	ICredentialDataDecryptedObject,
	ICredentialTestRequest,
	ICredentialType,
	IHttpRequestHelper,
	Icon,
	INodeProperties,
} from 'n8n-workflow';

import { createClientAssertion, getTokenEndpoint } from './helpers/clientAssertion';

export class SharePointPnpCertificateApi implements ICredentialType {
	name = 'sharePointPnpCertificateApi';

	displayName = 'SharePoint PnP Certificate API';

	icon: Icon = { light: 'file:icons/sharepoint.svg', dark: 'file:icons/sharepoint.svg' };

	documentationUrl =
		'https://github.com/octanesoftware/n8n-nodes-sharepoint-pnp#certificate-app-only';

	properties: INodeProperties[] = [
		{
			displayName: 'Directory (Tenant) ID',
			name: 'tenantId',
			type: 'string',
			default: '',
			required: true,
		},
		{
			displayName: 'Application (Client) ID',
			name: 'clientId',
			type: 'string',
			default: '',
			required: true,
		},
		{
			displayName: 'Subdomain',
			name: 'subdomain',
			type: 'string',
			default: '',
			required: true,
			placeholder: 'contoso',
			hint: 'The part before ".sharepoint.com". For "https://contoso.sharepoint.com" enter "contoso".',
		},
		{
			displayName: 'Private Key',
			name: 'privateKey',
			type: 'string',
			typeOptions: { password: true, rows: 6 },
			default: '',
			required: true,
			placeholder: '-----BEGIN PRIVATE KEY-----',
			description:
				'The PEM private key of the certificate registered on the Entra app. You can paste the certificate PEM after the key, then the thumbprint is not necessary.',
		},
		{
			displayName: 'Private Key Passphrase',
			name: 'passphrase',
			type: 'string',
			typeOptions: { password: true },
			default: '',
			description: 'Only necessary if the private key is encrypted',
		},
		{
			displayName: 'Certificate',
			name: 'certificate',
			type: 'string',
			typeOptions: { password: true, rows: 4 },
			default: '',
			placeholder: '-----BEGIN CERTIFICATE-----',
			description:
				'The public certificate PEM. Use this field or Certificate Thumbprint, unless the certificate is in the Private Key field.',
		},
		{
			displayName: 'Certificate Thumbprint',
			name: 'thumbprint',
			type: 'string',
			default: '',
			description: 'The SHA-1 thumbprint shown on the Entra app "Certificates & secrets" page',
		},
		{
			displayName: 'Access Token',
			name: 'accessToken',
			type: 'hidden',
			typeOptions: { expirable: true, password: true },
			default: '',
		},
	];

	async preAuthentication(this: IHttpRequestHelper, credentials: ICredentialDataDecryptedObject) {
		const assertion = createClientAssertion({
			tenantId: credentials.tenantId as string,
			clientId: credentials.clientId as string,
			privateKey: credentials.privateKey as string,
			passphrase: credentials.passphrase as string | undefined,
			certificate: credentials.certificate as string | undefined,
			thumbprint: credentials.thumbprint as string | undefined,
		});

		const form = new URLSearchParams({
			grant_type: 'client_credentials',
			client_id: (credentials.clientId as string).trim(),
			client_assertion_type: 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer',
			client_assertion: assertion,
			scope: `https://${(credentials.subdomain as string).trim()}.sharepoint.com/.default`,
		});

		const response = (await this.helpers.httpRequest({
			method: 'POST',
			url: getTokenEndpoint(credentials.tenantId as string),
			headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
			body: form.toString(),
		})) as { access_token: string };

		return { accessToken: response.access_token };
	}

	authenticate: IAuthenticateGeneric = {
		type: 'generic',
		properties: {
			headers: {
				Authorization: '=Bearer {{$credentials.accessToken}}',
			},
		},
	};

	test: ICredentialTestRequest = {
		request: {
			baseURL: '=https://{{$credentials.subdomain}}.sharepoint.com',
			url: '/_api/web?$select=Title',
			headers: { Accept: 'application/json;odata=nometadata' },
		},
	};
}
