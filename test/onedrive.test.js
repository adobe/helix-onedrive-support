/*
 * Copyright 2019 Adobe. All rights reserved.
 * This file is licensed to you under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License. You may obtain a copy
 * of the License at http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software distributed under
 * the License is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR REPRESENTATIONS
 * OF ANY KIND, either express or implied. See the License for the specific language
 * governing permissions and limitations under the License.
 */

/* eslint-env mocha */
import assert from 'assert';
import { UnsecuredJWT } from 'jose';
import { OneDrive } from '../src/OneDrive.js';
import { OneDriveAuth } from '../src/OneDriveAuth.js';
import { OneDriveMock as MockDrive } from '../src/OneDriveMock.js';
import { StatusCodeError } from '../src/index.js';
import { Nock } from './utils.js';

/**
 * @param {OneDriveAuthOptions} opts
 * @returns {OneDriveAuth}
 */
const DEFAULT_AUTH = (opts = {}) => new OneDriveAuth({
  clientId: 'foo',
  localAuthCache: true,
  noTenantCache: true,
  ...opts,
}).setAccessToken(new UnsecuredJWT({
  email: 'bob',
  tid: 'test-tenantid',
}).encode());

describe('OneDrive Tests', () => {
  let nock;
  beforeEach(() => {
    nock = new Nock();
    delete process.env.HELIX_ONEDRIVE_NO_SHARE_LINK_CACHE;
  });

  afterEach(() => {
    nock.done();
  });

  it('throws when required parameters are not specified.', async () => {
    assert.throws(() => new OneDrive({}));
  });

  it('can be constructed.', async () => {
    const drive = new OneDrive({
      auth: DEFAULT_AUTH(),
    });
    assert.ok(drive);
  });

  it('can be disposed.', async () => {
    const drive = new OneDrive({
      auth: DEFAULT_AUTH(),
    });
    await assert.doesNotReject(async () => drive.dispose());
  });

  it('can encode a share link', () => {
    assert.deepStrictEqual(
      OneDrive.encodeSharingUrl('https://onedrive.com/a/b/c/d'),
      'u!aHR0cHM6Ly9vbmVkcml2ZS5jb20vYS9iL2MvZA=',
    );
  });

  it('can encode a share link as url', () => {
    assert.deepStrictEqual(
      OneDrive.encodeSharingUrl(new URL('https://onedrive.com/a/b/c/d')),
      'u!aHR0cHM6Ly9vbmVkcml2ZS5jb20vYS9iL2MvZA=',
    );
  });

  it('can convert a drive item to a uri', () => {
    assert.deepStrictEqual(OneDrive.driveItemToURL({
      id: 'item-id',
      parentReference: {
        driveId: 'drive-id',
      },
    }), new URL('onedrive:/drives/drive-id/items/item-id'));
  });

  it('can convert an url to a drive item', () => {
    assert.deepStrictEqual({
      id: 'item-id',
      parentReference: {
        driveId: 'drive-id',
      },
    }, OneDrive.driveItemFromURL(new URL('onedrive:/drives/drive-id/items/item-id')));
  });

  it('can convert an url string to a drive item', () => {
    assert.deepStrictEqual({
      id: 'item-id',
      parentReference: {
        driveId: 'drive-id',
      },
    }, OneDrive.driveItemFromURL('onedrive:/drives/drive-id/items/item-id'));
  });

  it('returns null for non onedrive urls', () => {
    assert.strictEqual(OneDrive.driveItemFromURL('https://www.example.com/drives/drive-id/items/item-id'), null);
  });

  it('throws an error for onedrive with wrong format (missing drives)', () => {
    assert.throws(
      () => OneDrive.driveItemFromURL('onedrive:/drive-id/items/item-id'),
      new Error('URI not supported (missing \'drives\' segment): onedrive:/drive-id/items/item-id'),
    );
  });

  it('throws an error for onedrive with wrong format (missing items)', () => {
    assert.throws(
      () => OneDrive.driveItemFromURL('onedrive:/drives/drive-id/item-id'),
      new Error('URI not supported (missing \'items\' segment): onedrive:/drives/drive-id/item-id'),
    );
  });

  it('fuzzyGetDriveItem throws an error if relPath does not start with /', async () => {
    const drive = new MockDrive();
    const folderItem = OneDrive.driveItemFromURL('onedrive:/drives/123/items/456');

    await assert.rejects(
      drive.fuzzyGetDriveItem(folderItem, 'foo'),
      new Error('relPath must be empty or start with /'),
    );
  });

  it('fuzzyGetDriveItem returns folder item when relpath is missing', async () => {
    const folderItem = OneDrive.driveItemFromURL('onedrive:/drives/123/items/456');

    const data = {
      folder: { childCount: 1 },
      name: 'test',
    };
    const drive = new MockDrive()
      .registerDriveItem('123', '456', data);
    const res = await drive.fuzzyGetDriveItem(folderItem);

    assert.deepStrictEqual(res, [data]);
  });

  it('fuzzyGetDriveItem returns exact item', async () => {
    const folderItem = OneDrive.driveItemFromURL('onedrive:/drives/123/items/456');
    const drive = new MockDrive()
      .registerDriveItem('123', '456:/document.docx', {
        file: { mimeType: 'dummy' },
        name: 'document.docx',
      });
    const res = await drive.fuzzyGetDriveItem(folderItem, '/document.docx');

    assert.deepStrictEqual(res, [{
      extension: 'docx',
      file: {
        mimeType: 'dummy',
      },
      name: 'document.docx',
    }]);
  });

  it('fuzzyGetDriveItem throw error on direct item', async () => {
    const folderItem = OneDrive.driveItemFromURL('onedrive:/drives/123/items/456');
    const drive = new MockDrive()
      .registerDriveItem('123', '456:/document.docx', new StatusCodeError('rate limit', 429));
    await assert.rejects(drive.fuzzyGetDriveItem(folderItem, '/document.docx'), new Error('rate limit'));
  });

  it('fuzzyGetDriveItem returns item for deep path', async () => {
    const folderItem = OneDrive.driveItemFromURL('onedrive:/drives/123/items/456');
    const data = {
      value: [{
        file: { mimeType: 'dummy' },
        name: 'document.docx',
      }],
    };

    const drive = new MockDrive()
      .registerDriveItemChildren('123', '456:/publish/en:', data); // this is a bit a hack to trick OneDriveMock
    const res = await drive.fuzzyGetDriveItem(folderItem, '/publish/en/document.docx');

    assert.deepStrictEqual(res, [{
      extension: 'docx',
      file: {
        mimeType: 'dummy',
      },
      fuzzyDistance: 0,
      name: 'document.docx',
    }]);
  });

  it('fuzzyGetDriveItem returns empty array for non existing item', async () => {
    const folderItem = OneDrive.driveItemFromURL('onedrive:/drives/123/items/456');
    const data = {
      value: [],
    };

    const drive = new MockDrive()
      .registerDriveItemChildren('123', '456', data);
    const res = await drive.fuzzyGetDriveItem(folderItem, '/document.docx');

    assert.deepStrictEqual(res, []);
  });

  it('fuzzyGetDriveItem returns matching items', async () => {
    const folderItem = OneDrive.driveItemFromURL('onedrive:/drives/123/items/456');
    const data = {
      value: [{
        file: { mimeType: 'dummy' },
        name: 'My 1. Document.docx',
      }, {
        file: { mimeType: 'dummy' },
        name: 'my-1-document.docx',
      }, {
        file: { mimeType: 'dummy' },
        name: 'my-1-document".docx',
      }, {
        file: { mimeType: 'dummy' },
        name: 'My-1-Document.docx',
      }, {
        file: { mimeType: 'dummy' },
        name: 'My-1-Document.md',
      }, {
        folder: { childCount: 1 },
        name: 'folder',
      }],
    };

    const drive = new MockDrive()
      .registerDriveItemChildren('123', '456', data);
    const res = await drive.fuzzyGetDriveItem(folderItem, '/my-1-document.docx');

    assert.deepStrictEqual(res, [
      {
        extension: 'docx',
        file: { mimeType: 'dummy' },
        fuzzyDistance: 0,
        name: 'my-1-document.docx',
      },
      {
        extension: 'docx',
        file: { mimeType: 'dummy' },
        fuzzyDistance: 1,
        name: 'my-1-document".docx',
      },
      {
        extension: 'docx',
        file: { mimeType: 'dummy' },
        fuzzyDistance: 2,
        name: 'My-1-Document.docx',
      },
      {
        extension: 'docx',
        file: { mimeType: 'dummy' },
        fuzzyDistance: 5,
        name: 'My 1. Document.docx',
      },
    ]);
  });

  it('fuzzyGetDriveItem returns matching items w/o extension', async () => {
    const folderItem = OneDrive.driveItemFromURL('onedrive:/drives/123/items/456');
    const data = {
      value: [{
        file: { mimeType: 'dummy' },
        name: 'My 1. Document.docx',
      }, {
        file: { mimeType: 'dummy' },
        name: 'my-1-document.docx',
      }, {
        file: { mimeType: 'dummy' },
        name: 'my-1-document.md',
      }, {
        file: { mimeType: 'dummy' },
        name: 'My-1-Document.docx',
      }],
    };

    const drive = new MockDrive()
      .registerDriveItemChildren('123', '456', data);
    const res = await drive.fuzzyGetDriveItem(folderItem, '/my-1-document');

    assert.deepStrictEqual(res, [
      {
        extension: 'docx',
        file: { mimeType: 'dummy' },
        fuzzyDistance: 0,
        name: 'my-1-document.docx',
      },
      {
        extension: 'md',
        file: { mimeType: 'dummy' },
        fuzzyDistance: 0,
        name: 'my-1-document.md',
      },
      {
        extension: 'docx',
        file: { mimeType: 'dummy' },
        fuzzyDistance: 2,
        name: 'My-1-Document.docx',
      },
      {
        extension: 'docx',
        file: { mimeType: 'dummy' },
        fuzzyDistance: 5,
        name: 'My 1. Document.docx',
      },
    ]);
  });

  it('fuzzyGetDriveItem returns matching items (ignore extension)', async () => {
    const folderItem = OneDrive.driveItemFromURL('onedrive:/drives/123/items/456');
    const data = {
      value: [{
        file: { mimeType: 'dummy' },
        name: 'My 1. Document.docx',
      }, {
        file: { mimeType: 'dummy' },
        name: 'my-1-document.docx',
      }, {
        file: { mimeType: 'dummy' },
        name: 'my-1-document.md',
      }, {
        file: { mimeType: 'dummy' },
        name: 'My-1-Document.docx',
      }],
    };

    const drive = new MockDrive()
      .registerDriveItemChildren('123', '456', data);
    const res = await drive.fuzzyGetDriveItem(folderItem, '/my-1-document.docx', true);

    assert.deepStrictEqual(res, [
      {
        extension: 'docx',
        file: { mimeType: 'dummy' },
        fuzzyDistance: 0,
        name: 'my-1-document.docx',
      },
      {
        extension: 'md',
        file: { mimeType: 'dummy' },
        fuzzyDistance: 0,
        name: 'my-1-document.md',
      },
      {
        extension: 'docx',
        file: { mimeType: 'dummy' },
        fuzzyDistance: 2,
        name: 'My-1-Document.docx',
      },
      {
        extension: 'docx',
        file: { mimeType: 'dummy' },
        fuzzyDistance: 5,
        name: 'My 1. Document.docx',
      },
    ]);
  });

  it('fuzzyGetDriveItem iterates over pages', async () => {
    const folderItem = OneDrive.driveItemFromURL('onedrive:/drives/123/items/456');
    const data = {
      value: [],
    };
    for (let i = 0; i < 5000; i += 1) {
      data.value.push({
        file: { mimeType: 'dummy' },
        name: `dummy-document-${i}.docx`,
      });
    }
    data.value.push({
      file: { mimeType: 'dummy' },
      name: 'My 1. Document.docx',
    });

    const drive = new MockDrive()
      .registerDriveItemChildren('123', '456', data);
    const res = await drive.fuzzyGetDriveItem(folderItem, '/my-1-document');

    assert.deepStrictEqual(res, [{
      extension: 'docx',
      file: { mimeType: 'dummy' },
      fuzzyDistance: 5,
      name: 'My 1. Document.docx',
    }]);
  });

  it('can get the user profile: me', async () => {
    const expected = {
      '@odata.context': 'https://graph.microsoft.com/v1.0/$metadata#users/$entity',
      businessPhones: [],
      displayName: 'Project Helix Integration',
      givenName: 'Project',
      id: 'c96b3b1f-5489-4639-8100-d67739af7d3e',
      mail: 'helix@adobe.com',
      surname: 'Helix Integration',
      userPrincipalName: 'helix@adobe.com',
    };

    nock('https://graph.microsoft.com/v1.0')
      .get('/me')
      .reply(200, expected);

    const od = new OneDrive({
      auth: DEFAULT_AUTH(),
    });
    const me = await od.me();
    assert.deepStrictEqual(me, expected);
  });

  it('can get the root item', async () => {
    const expected = {
      '@odata.context': 'https://graph.microsoft.com/v1.0/$metadata#users(\'48d31887-5fad-4d73-a9f5-3c356e68a038\')/drive/root/$entity',
      createdDateTime: '2017-07-27T02:41:36Z',
      id: '01BYE5RZ56Y2GOVW7725BZO354PWSELRRZ',
      lastModifiedDateTime: '2022-11-10T06:33:51Z',
      name: 'root',
      webUrl: 'https://m365x214355-my.sharepoint.com/personal/meganb_m365x214355_onmicrosoft_com/Documents',
      size: 106329756,
      parentReference: {
        driveType: 'business',
        driveId: 'b!-RIj2DuyvEyV1T4NlOaMHk8XkS_I8MdFlUCq1BlcjgmhRfAj3-Z8RY2VpuvV_tpd',
      },
      fileSystemInfo: {
        createdDateTime: '2017-07-27T02:41:36Z',
        lastModifiedDateTime: '2022-11-10T06:33:51Z',
      },
      folder: {
        childCount: 38,
      },
      root: {},
    };

    nock('https://graph.microsoft.com/v1.0')
      .get(`/drives/${expected.id}/root`)
      .reply(200, expected);

    const od = new OneDrive({
      auth: DEFAULT_AUTH(),
    });
    const root = await od.getDriveRootItem(expected.id);
    assert.deepStrictEqual(root, expected);
  });

  it('uses global share link cache by default', async () => {
    nock('https://graph.microsoft.com/v1.0')
      .get('/shares/u!aHR0cHM6Ly9vbmVkcml2ZS5jb20vYS9iL2MvZDA/driveItem')
      .reply(200, {
        id: 'some-id',
      });

    const od = new OneDrive({
      auth: DEFAULT_AUTH(),
    });
    const item1 = await od.getDriveItemFromShareLink('https://onedrive.com/a/b/c/d0');
    assert.deepStrictEqual(item1, {
      id: 'some-id',
    });
    const item2 = await od.getDriveItemFromShareLink('https://onedrive.com/a/b/c/d0');
    assert.strictEqual(item1, item2);

    const item3 = await od.getDriveItemFromShareLink('onedrive:/drives/drive-id/items/item-id');
    assert.deepStrictEqual(item3, {
      id: 'item-id',
      parentReference: {
        driveId: 'drive-id',
      },
    });
  });

  it('share link cache can be disabled via option', async () => {
    nock('https://graph.microsoft.com/v1.0')
      .get('/shares/u!aHR0cHM6Ly9vbmVkcml2ZS5jb20vYS9iL2MvZDAx/driveItem')
      .twice()
      .reply(200, {
        id: 'some-id',
      });

    const od = new OneDrive({
      auth: DEFAULT_AUTH(),
      noShareLinkCache: true,
    });
    const item1 = await od.getDriveItemFromShareLink('https://onedrive.com/a/b/c/d01');
    assert.deepStrictEqual(item1, {
      id: 'some-id',
    });
    const item2 = await od.getDriveItemFromShareLink('https://onedrive.com/a/b/c/d01');
    assert.deepStrictEqual(item1, item2);
    assert.notStrictEqual(item1, item2);
  });

  it('share link cache can be disabled via env', async () => {
    process.env.HELIX_ONEDRIVE_NO_SHARE_LINK_CACHE = true;
    nock('https://graph.microsoft.com/v1.0')
      .get('/shares/u!aHR0cHM6Ly9vbmVkcml2ZS5jb20vYS9iL2MvZDE/driveItem')
      .twice()
      .reply(200, {
        id: 'some-id',
      });

    const od = new OneDrive({
      auth: DEFAULT_AUTH(),
    });
    const item1 = await od.getDriveItemFromShareLink('https://onedrive.com/a/b/c/d1');
    assert.deepStrictEqual(item1, {
      id: 'some-id',
    });
    const item2 = await od.getDriveItemFromShareLink('https://onedrive.com/a/b/c/d1');
    assert.deepStrictEqual(item1, item2);
    assert.notStrictEqual(item1, item2);
  });

  it('share link cache can be supplied via opts', async () => {
    nock('https://graph.microsoft.com/v1.0')
      .get('/shares/u!aHR0cHM6Ly9vbmVkcml2ZS5jb20vYS9iL2MvZDI/driveItem')
      .reply(200, {
        id: 'some-id',
      });

    const shareLinkCache = new Map();

    const od = new OneDrive({
      auth: DEFAULT_AUTH(),
      shareLinkCache,
    });
    const item1 = await od.getDriveItemFromShareLink('https://onedrive.com/a/b/c/d2');
    assert.deepStrictEqual(item1, {
      id: 'some-id',
    });
    const item2 = await od.getDriveItemFromShareLink('https://onedrive.com/a/b/c/d2');
    assert.strictEqual(item1, item2);

    assert.deepStrictEqual(Object.fromEntries(shareLinkCache.entries()), {
      'https://onedrive.com/a/b/c/d2': {
        id: 'some-id',
      },
    });
  });

  it('returns a 404 when fetching a share link is forbidden', async () => {
    nock('https://graph.microsoft.com/v1.0')
      .get('/shares/u!aHR0cHM6Ly9vbmVkcml2ZS5jb20vYS9iL2MvZDI/driveItem')
      .reply(403, 'not allowed');

    const od = new OneDrive({
      auth: DEFAULT_AUTH(),
    });
    const error = new StatusCodeError('not allowed', 404);
    await assert.rejects(
      async () => od.getDriveItemFromShareLink('https://onedrive.com/a/b/c/d2'),
      error,
    );
  });

  it('returns original error when fetching a share link returns neither 401 nor 403', async () => {
    nock('https://graph.microsoft.com/v1.0')
      .get('/shares/u!aHR0cHM6Ly9vbmVkcml2ZS5jb20vYS9iL2MvZDI/driveItem')
      .reply(500, 'kaputt');

    const od = new OneDrive({
      auth: DEFAULT_AUTH(),
    });
    const error = new StatusCodeError('kaputt', 500);
    await assert.rejects(
      async () => od.getDriveItemFromShareLink('https://onedrive.com/a/b/c/d2'),
      error,
    );
  });

  it('propagates errors', async () => {
    nock('https://graph.microsoft.com/v1.0')
      .get('/me')
      .reply(400, 'wrong input');

    const od = new OneDrive({
      auth: DEFAULT_AUTH(),
    });
    await assert.rejects(od.me(), new StatusCodeError('wrong input', 400));
  }).timeout(5000);

  it('propagates 404s as warnings', async () => {
    nock('https://graph.microsoft.com/v1.0')
      .get('/me')
      .reply(404, 'not found');

    const od = new OneDrive({
      auth: DEFAULT_AUTH(),
    });
    await assert.rejects(od.me(), new StatusCodeError('not found', 404));
  }).timeout(5000);

  it('propagates other failures in fetch', async () => {
    nock('https://graph.microsoft.com/v1.0')
      .get('/me')
      .replyWithError(new Error('kaputt'));

    const od = new OneDrive({
      auth: DEFAULT_AUTH(),
    });
    await assert.rejects(od.me(), new Error('kaputt', 500));
  }).timeout(5000);

  it('uploadDriveItem fails with bad conflict behaviour', async () => {
    const drive = new OneDrive({
      auth: DEFAULT_AUTH(),
    });
    await assert.rejects(
      async () => drive.uploadDriveItem(Buffer.alloc(0), 'item', '', 'guess'),
      /Error: Bad confict behaviour/,
    );
  });

  it('uploadDriveItem uploads file', async () => {
    const driveItem = {
      parentReference: {
        driveId: '1',
      },
      id: '2',
    };
    nock('https://graph.microsoft.com/v1.0')
      .put(`/drives/${driveItem.parentReference.driveId}/items/${driveItem.id}/content`)
      .query(true)
      .reply(200, {
        id: driveItem.id,
      });
    const drive = new OneDrive({
      auth: DEFAULT_AUTH(),
    });
    const result = await drive.uploadDriveItem(Buffer.alloc(2), driveItem);
    assert(result);
  });

  it('uploadDriveItem uploads file with relative path', async () => {
    const relPath = '/3';
    const driveItem = {
      parentReference: {
        driveId: '1',
      },
      id: '2',
    };
    nock('https://graph.microsoft.com/v1.0')
      .put(`/drives/${driveItem.parentReference.driveId}/items/${driveItem.id}:${relPath}:/content`)
      .query(true)
      .reply(200, {
        id: driveItem.id,
      });
    const drive = new OneDrive({
      auth: DEFAULT_AUTH(),
    });
    const result = await drive.uploadDriveItem(Buffer.alloc(2), driveItem, relPath);
    assert(result);
  });

  it('downloadDriveItem returns buffer', async () => {
    const driveItem = {
      parentReference: {
        driveId: '1',
      },
      id: '2',
    };
    nock('https://graph.microsoft.com/v1.0')
      .get(`/drives/${driveItem.parentReference.driveId}/items/${driveItem.id}/content`)
      .reply(200, 'This is the contents of that file');
    const od = new OneDrive({
      auth: DEFAULT_AUTH(),
    });
    const result = await od.downloadDriveItem(driveItem);
    assert(Buffer.isBuffer(result));
  });

  it('getDriveItem returns buffer when asked to download', async () => {
    const relPath = '/3';
    const folderItem = {
      parentReference: {
        driveId: '1',
      },
      id: '2',
    };
    nock('https://graph.microsoft.com/v1.0')
      .get(`/drives/${folderItem.parentReference.driveId}/items/${folderItem.id}:${relPath}:/content`)
      .reply(200, 'This is the contents of that file');
    const od = new OneDrive({
      auth: DEFAULT_AUTH(),
    });
    const result = await od.getDriveItem(folderItem, relPath, true);
    assert(Buffer.isBuffer(result));
  });

  it('can getWorkbook', async () => {
    const fileItem = {
      parentReference: {
        driveId: '1',
      },
      id: '2',
    };
    const od = new OneDrive({
      auth: DEFAULT_AUTH(),
    });
    const workbook = od.getWorkbook(fileItem);
    assert.strictEqual(workbook.uri, `/drives/${fileItem.parentReference.driveId}/items/${fileItem.id}/workbook`);
  });

  it('can list subscriptions', async () => {
    const subscriptions = {
      '@odata.context': 'https://graph.microsoft.com/v1.0/$metadata#subscriptions',
      value: [
        {
          id: '0fc0d6db-0073-42e5-a186-853da75fb308',
          resource: 'Users',
          applicationId: '24d3b144-21ae-4080-943f-7067b395b913',
          changeType: 'updated,deleted',
          clientState: null,
          notificationUrl: 'https://webhookappexample.azurewebsites.net/api/notifications',
          lifecycleNotificationUrl: 'https://webhook.azurewebsites.net/api/send/lifecycleNotifications',
          expirationDateTime: '2018-03-12T05:00:00Z',
          creatorId: '8ee44408-0679-472c-bc2a-692812af3437',
          latestSupportedTlsVersion: 'v1_2',
          encryptionCertificate: '',
          encryptionCertificateId: '',
          includeResourceData: false,
          notificationContentType: 'application/json',
        },
      ],
    };
    nock('https://graph.microsoft.com/v1.0')
      .get('/subscriptions')
      .reply(200, subscriptions);
    const od = new OneDrive({
      auth: DEFAULT_AUTH(),
    });
    const result = await od.listSubscriptions();
    assert(result, subscriptions);
  });

  it('can create subscriptions', async () => {
    const opts = {
      resource: '/me',
      notificationUrl: 'https://www.hlx.live/',
      clientState: 'confirmed',
      changeType: 'updated',
    };
    nock('https://graph.microsoft.com/v1.0')
      .post('/subscriptions')
      .reply((_, body) => {
        delete body.expirationDateTime;
        assert.deepStrictEqual(body, opts);
        return [200, {
          id: '7f105c7d-2dc5-4530-97cd-4e7ae6534c07',
        }];
      });
    const od = new OneDrive({
      auth: DEFAULT_AUTH(),
    });
    const result = await od.createSubscription({
      ...opts,
      expiresIn: 60000,
    });
    assert(result);
  });

  it('can refresh subscriptions', async () => {
    const id = '7f105c7d-2dc5-4530-97cd-4e7ae6534c07';
    const expiresIn = 60000;
    const now = Date.now();

    nock('https://graph.microsoft.com/v1.0')
      .patch(`/subscriptions/${id}`)
      .reply((_, body) => {
        const expires = Date.parse(body.expirationDateTime);
        assert(expires >= now + expiresIn);
        return [200, { id }];
      });
    const od = new OneDrive({
      auth: DEFAULT_AUTH(),
    });
    const result = await od.refreshSubscription(id, expiresIn);
    assert(result);
  });

  it('can delete subscriptions', async () => {
    const id = '7f105c7d-2dc5-4530-97cd-4e7ae6534c07';

    nock('https://graph.microsoft.com/v1.0')
      .delete(`/subscriptions/${id}`)
      .reply(200);
    const od = new OneDrive({
      auth: DEFAULT_AUTH(),
    });
    const result = await od.deleteSubscription(id);
    assert(result);
  });

  it('can fetch changes', async () => {
    const changes = {
      value: [{
        id: '123010204abac',
        name: 'file.txt',
        file: {},
      }],
      '@odata.nextLink': 'https://graph.microsoft.com/v1.0/me/drive/delta?token=1230919asd190410jlka',
    };
    nock('https://graph.microsoft.com/v1.0')
      .get('/me/drive/root/delta')
      .reply(200, changes)
      .get('/me/drive/root/delta?token=1230919asd190410jlka')
      .reply(200, {
        value: [],
        '@odata.deltaLink': 'https://graph.microsoft.com/v1.0/me/drive/delta?token=1230919asd190410jlkb',
      });
    const od = new OneDrive({
      auth: DEFAULT_AUTH(),
    });
    const result = await od.fetchChanges('/me/drive/root');
    assert.deepStrictEqual(result.token, '1230919asd190410jlkb');
  });

  it('can fetch changes with a token', async () => {
    const token = '1230919asd190410jlka';
    nock('https://graph.microsoft.com/v1.0')
      .get(`/me/drive/root/delta?token=${token}`)
      .reply(200, {
        value: [],
        '@odata.deltaLink': 'https://graph.microsoft.com/v1.0/me/drive/delta?token=1230919asd190410jlkb',
      });
    const od = new OneDrive({
      auth: DEFAULT_AUTH(),
    });
    const result = await od.fetchChanges('/me/drive/root', token);
    assert.deepStrictEqual(result.token, '1230919asd190410jlkb');
  });

  it('throws when neither next nor delta link is received', async () => {
    const token = '1230919asd190410jlka';
    nock('https://graph.microsoft.com/v1.0')
      .get(`/me/drive/root/delta?token=${token}`)
      .reply(200, {
        value: [],
      });
    const od = new OneDrive({
      auth: DEFAULT_AUTH(),
    });
    await assert.rejects(
      async () => od.fetchChanges('/me/drive/root', token),
      new StatusCodeError('Received response with neither next nor delta link.', 500),
    );
  });

  it('throws when site URL does not match expected format', async () => {
    const od = new OneDrive({
      auth: DEFAULT_AUTH(),
    });
    await assert.rejects(
      async () => od.getSite('https://www.hlx.live'),
      /Site URL does not match \(\*\.sharepoint.com\/sites\/\.\*\): /,
    );
  });

  it('returns site', async () => {
    const od = new OneDrive({
      auth: DEFAULT_AUTH(),
    });
    const site = await od.getSite('https://hlx-my.sharepoint.com/sites/mysites/site1');
    assert(site);
  });

  it('turns 401 into 404 when authentication fails', async () => {
    const od = new OneDrive({
      auth: {
        authenticate: async () => {
          throw new StatusCodeError('Unauthenticated', 401);
        },
        log: console,
      },
    });
    await assert.rejects(
      async () => od.getSite('https://hlx-my.sharepoint.com/sites/mysites/site1'),
      new StatusCodeError('Unauthenticated', 404),
    );
  });

  it('reports other failures as such', async () => {
    const od = new OneDrive({
      auth: {
        authenticate: async () => {
          throw new StatusCodeError('Forbidden', 403);
        },
        log: console,
      },
    });
    await assert.rejects(
      async () => od.getSite('https://hlx-my.sharepoint.com/sites/mysites/site1'),
      new StatusCodeError('Forbidden', 403),
    );
  });
});
