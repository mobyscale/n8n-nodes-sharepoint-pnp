import type { Icon, ICredentialTestRequest, ICredentialType, INodeProperties } from 'n8n-workflow';

export class SharePointPnpOAuth2Api implements ICredentialType {
	name = 'sharePointPnpOAuth2Api';

	extends = ['microsoftOAuth2Api'];

	displayName = 'SharePoint PnP OAuth2 API';

	icon: Icon = { light: 'file:icons/sharepoint.svg', dark: 'file:icons/sharepoint.svg' };

	documentationUrl = 'https://github.com/octanesoftware/n8n-nodes-sharepoint-pnp#oauth2-delegated';

	properties: INodeProperties[] = [
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
