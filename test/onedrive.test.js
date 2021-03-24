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

'use strict';

process.env.HELIX_FETCH_FORCE_HTTP1 = 'true';

const assert = require('assert');
const nock = require('nock');
const OneDrive = require('../src/OneDrive.js');
const MockDrive = require('../src/OneDriveMock.js');
const StatusCodeError = require('../src/StatusCodeError');

const AZ_AUTHORITY_HOST_URL = 'https://login.windows.net';

const { TEST_CLIENT_ID, TEST_USER, TEST_PASSWORD } = process.env;

describe('OneDrive Tests', () => {
  it('throws when required parameters are not specified.', async () => {
    assert.throws(() => new OneDrive({}));
  });

  it('can be constructed.', async () => {
    const drive = new OneDrive({
      clientId: 'foo', clientSecret: 'bar',
    });
    assert.ok(drive);
  });

  it('can encode a share link', () => {
    assert.deepEqual(OneDrive.encodeSharingUrl('https://onedrive.com/a/b/c/d'),
      'u!aHR0cHM6Ly9vbmVkcml2ZS5jb20vYS9iL2MvZA=');
  });

  it('can encode a share link as url', () => {
    assert.deepEqual(OneDrive.encodeSharingUrl(new URL('https://onedrive.com/a/b/c/d')),
      'u!aHR0cHM6Ly9vbmVkcml2ZS5jb20vYS9iL2MvZA=');
  });

  it('can convert a drive item to a uri', () => {
    assert.deepEqual(OneDrive.driveItemToURL({
      id: 'item-id',
      parentReference: {
        driveId: 'drive-id',
      },
    }), new URL('onedrive:/drives/drive-id/items/item-id'));
  });

  it('can convert an url to a drive item', () => {
    assert.deepEqual({
      id: 'item-id',
      parentReference: {
        driveId: 'drive-id',
      },
    }, OneDrive.driveItemFromURL(new URL('onedrive:/drives/drive-id/items/item-id')));
  });

  it('can convert an url string to a drive item', () => {
    assert.deepEqual({
      id: 'item-id',
      parentReference: {
        driveId: 'drive-id',
      },
    }, OneDrive.driveItemFromURL('onedrive:/drives/drive-id/items/item-id'));
  });

  it('returns null for non onedrive urls', () => {
    assert.equal(OneDrive.driveItemFromURL('https://www.example.com/drives/drive-id/items/item-id'), null);
  });

  it('throws an error for onedrive with wrong format (missing drives)', () => {
    assert.throws(() => OneDrive.driveItemFromURL('onedrive:/drive-id/items/item-id'),
      new Error('URI not supported (missing \'drives\' segment): onedrive:/drive-id/items/item-id'));
  });

  it('throws an error for onedrive with wrong format (missing items)', () => {
    assert.throws(() => OneDrive.driveItemFromURL('onedrive:/drives/drive-id/item-id'),
      new Error('URI not supported (missing \'items\' segment): onedrive:/drives/drive-id/item-id'));
  });

  it('fuzzyGetDriveItem returns folder item when relpath is missing', async () => {
    const folderItem = OneDrive.driveItemFromURL('onedrive:/drives/123/items/456');

    const data = {
      value: {
        folder: { childCount: 1 },
        name: 'test',
      },
    };
    const drive = new MockDrive()
      .registerDriveItem('123', '456', data);
    const res = await drive.fuzzyGetDriveItem(folderItem);

    assert.deepEqual(res, [data.value]);
  });

  it('fuzzyGetDriveItem returns exact item', async () => {
    const folderItem = OneDrive.driveItemFromURL('onedrive:/drives/123/items/456');
    const data = {
      value: [{
        file: { mimeType: 'dummy' },
        name: 'document.docx',
      }],
    };

    const drive = new MockDrive()
      .registerDriveItemChildren('123', '456', data);
    const res = await drive.fuzzyGetDriveItem(folderItem, '/document.docx');

    assert.deepEqual(res, [{
      extension: 'docx',
      file: {
        mimeType: 'dummy',
      },
      fuzzyDistance: 0,
      name: 'document.docx',
    }]);
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

    assert.deepEqual(res, [{
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

    assert.deepEqual(res, []);
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
      }],
    };

    const drive = new MockDrive()
      .registerDriveItemChildren('123', '456', data);
    const res = await drive.fuzzyGetDriveItem(folderItem, '/my-1-document.docx');

    assert.deepEqual(res, [
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

    assert.deepEqual(res, [
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

    assert.deepEqual(res, [{
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

    nock(AZ_AUTHORITY_HOST_URL)
      .post('/common/oauth2/token?api-version=1.0')
      .reply(200, {
        token_type: 'Bearer',
        refresh_token: 'dummy',
        access_token: 'dummy',
        expires_in: 81000,
      })
      .get('/common/UserRealm/test-user?api-version=1.0')
      .reply(200, {
        account_type: 'managed',
      });
    nock('https://graph.microsoft.com/v1.0')
      .get('/me')
      .reply(200, expected);

    const od = new OneDrive({
      clientId: TEST_CLIENT_ID || 'foobar',
      username: TEST_USER || 'test-user',
      password: TEST_PASSWORD || 'test-password',
    });
    const me = await od.me();
    assert.deepEqual(me, expected);
  }).timeout(5000);

  it('propagates errors', async () => {
    nock(AZ_AUTHORITY_HOST_URL)
      .post('/common/oauth2/token?api-version=1.0')
      .reply(200, {
        token_type: 'Bearer',
        refresh_token: 'dummy',
        access_token: 'dummy',
        expires_in: 81000,
      })
      .get('/common/UserRealm/test-user?api-version=1.0')
      .reply(200, {
        account_type: 'managed',
      });
    nock('https://graph.microsoft.com/v1.0')
      .get('/me')
      .reply(400, 'wrong input');

    const od = new OneDrive({
      clientId: TEST_CLIENT_ID || 'foobar',
      username: TEST_USER || 'test-user',
      password: TEST_PASSWORD || 'test-password',
    });
    await assert.rejects(od.me(), new StatusCodeError('wrong input', 400));
  }).timeout(5000);
});
