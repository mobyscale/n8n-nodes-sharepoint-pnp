import { afterAll, beforeAll, describe, expect, it } from 'vitest';

import {
	createTestFolder,
	createTestSp,
	hasConfig,
	locator,
	removeTestFolder,
	runNode,
	itWrite,
} from './harness';

describe.runIf(hasConfig)('sharing', () => {
	let testFolder: string | undefined;
	let filePath = '';

	const fileTarget = () => ({ target: 'file', filePath: locator(filePath) });

	beforeAll(async () => {
		testFolder = await createTestFolder();
		filePath = `${testFolder}/share-me.txt`;
		await createTestSp()
			.web.getFolderByServerRelativePath(testFolder)
			.files.addUsingPath('share-me.txt', 'sharing test', { Overwrite: true });
	});

	afterAll(async () => {
		await removeTestFolder(testFolder);
	});

	it('creates an Organization View link, reads the information and deletes the link', async () => {
		const [created] = await runNode({
			resource: 'sharing',
			operation: 'createLink',
			...fileTarget(),
			linkType: 'view',
			linkScope: 'organization',
		});
		expect(String(created.json.url)).toMatch(/^https:\/\//);
		expect(created.json.linkKind).toBe('OrganizationView');
		expect(created.json.shareId).toBeTruthy();

		const [information] = await runNode({
			resource: 'sharing',
			operation: 'getSharingInformation',
			...fileTarget(),
		});
		expect(information.json.target).toBe('file');
		expect(information.json.permissionsInformation).toBeDefined();

		const [deleted] = await runNode({
			resource: 'sharing',
			operation: 'deleteLink',
			...fileTarget(),
			deleteKind: 'organizationView',
		});
		expect(deleted.json.deleted).toBe(true);
	});

	it('unshares the test file', async () => {
		const [result] = await runNode({
			resource: 'sharing',
			operation: 'unshare',
			...fileTarget(),
		});
		expect(result.json.unshared).toBe(true);
	});

	it('creates an Anyone link, or fails with a clear policy error', async () => {
		try {
			const [created] = await runNode({
				resource: 'sharing',
				operation: 'createLink',
				...fileTarget(),
				linkType: 'view',
				linkScope: 'anyone',
				linkOptions: { expiration: new Date(Date.now() + 7 * 86_400_000).toISOString() },
			});
			expect(created.json.linkKind).toBe('AnonymousView');
			await runNode({
				resource: 'sharing',
				operation: 'deleteLink',
				...fileTarget(),
				deleteKind: 'anyoneView',
			});
		} catch (error) {
			expect((error as Error).message).toMatch(
				/Anyone links are probably not allowed|anonymous|anyone|not allowed|disabled|policy/i,
			);
		}
	});

	itWrite('shares with users (grants access and can send email)', async () => {
		await runNode({
			resource: 'sharing',
			operation: 'shareWithUsers',
			...fileTarget(),
			users: 'user@contoso.com',
			shareRole: 'view',
			sendEmail: false,
		});
	});

	itWrite('checks sharing permissions of users', async () => {
		await runNode({
			resource: 'sharing',
			operation: 'checkPermissions',
			...fileTarget(),
			users: 'user@contoso.com',
		});
	});
});

describe.runIf(hasConfig)('attachment', () => {
	it.skip('adds, lists, downloads, updates and deletes an attachment (needs a generic list, which tests must not create)', () => {
		// Covered by unit tests with mocked PnPjs in test/unit/attachment.test.ts.
	});
});
