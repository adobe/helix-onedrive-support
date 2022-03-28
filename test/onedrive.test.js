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
const assert = require('assert');
const jose = require('jose');
const { OneDrive } = require('../src/OneDrive.js');
const { OneDriveAuth } = require('../src/OneDriveAuth.js');
const { OneDriveMock: MockDrive } = require('../src/OneDriveMock.js');
const StatusCodeError = require('../src/StatusCodeError');
const { Nock } = require('./utils.js');

/**
 * @param {OneDriveAuthOptions} opts
 * @returns {OneDriveAuth}
 */
const DEFAULT_AUTH = (opts = {}) => new OneDriveAuth({
  clientId: 'foo',
  localAuthCache: true,
  noTenantCache: true,
  ...opts,
}).setAccessToken(new jose.UnsecuredJWT({
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

    assert.deepStrictEqual(res, [data.value]);
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

    assert.deepStrictEqual(res, [{
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

  it('propagates errors', async () => {
    nock('https://graph.microsoft.com/v1.0')
      .get('/me')
      .reply(400, 'wrong input');

    const od = new OneDrive({
      auth: DEFAULT_AUTH(),
    });
    await assert.rejects(od.me(), new StatusCodeError('wrong input', 400));
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
});
