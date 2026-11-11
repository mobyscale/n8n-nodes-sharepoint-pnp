import type { Icon, ICredentialTestRequest, ICredentialType, INodeProperties } from 'n8n-workflow';

export class SharePointPnpOAuth2Api implements ICredentialType {
	name = 'sharePointPnpOAuth2Api';

	extends = ['microsoftOAuth2Api'];

	displayName = 'SharePoint PnP OAuth2 API';

	icon: Icon = { light: 'file:icons/sharepoint-pnp.svg', dark: 'file:icons/sharepoint-pnp.svg' };

	documentationUrl =
		'https://github.com/mobyscale/n8n-nodes-sharepoint-pnp/blob/main/docs/credentials.md#oauth2-delegated';

	properties: INodeProperties[] = [
		{
			displayName: 'Tenant ID',
			name: 'tenantId',
			type: 'string',
			default: 'common',
			placeholder: '00000000-0000-0000-0000-000000000000',
			hint: 'The directory (tenant) ID of the app registration. Keep "common" only for a multi-tenant app.',
		},
		// The base credential uses the /common endpoint, which single-tenant apps cannot use.
		{
			displayName: 'Authorization URL',
			name: 'authUrl',
			type: 'hidden',
			default:
				'=https://login.microsoftonline.com/{{$self["tenantId"] || "common"}}/oauth2/v2.0/authorize',
		},
		{
			displayName: 'Access Token URL',
			name: 'accessTokenUrl',
			type: 'hidden',
			default:
				'=https://login.microsoftonline.com/{{$self["tenantId"] || "common"}}/oauth2/v2.0/token',
		},
		{
			displayName: 'Scope',
			name: 'scope',
			type: 'hidden',
			default: '=openid offline_access https://{{$self.subdomain}}.sharepoint.com/.default',
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
	];

	test: ICredentialTestRequest = {
		request: {
			baseURL: '=https://{{$credentials.subdomain}}.sharepoint.com',
			url: '/_api/web?$select=Title',
			headers: { Accept: 'application/json;odata=nometadata' },
		},
	};
}
