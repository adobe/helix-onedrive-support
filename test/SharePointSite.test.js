/*
 * Copyright 2021 Adobe. All rights reserved.
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
import crypto from 'crypto';
import { SharePointSite } from '../src/SharePointSite.js';
import { Nock } from './utils.js';

describe('SharePointSite Tests', () => {
  let nock;
  beforeEach(() => {
    nock = new Nock();
  });

  afterEach(() => {
    nock.done();
  });

  it('can be constructed', () => {
    assert.doesNotThrow(() => new SharePointSite({}));
  });

  it('retrieves access token', async () => {
    nock('https://login.microsoftonline.com')
      .post('/tenantId/oauth2/v2.0/token')
      .reply(200, () => JSON.stringify({
        access_token: crypto.randomUUID(),
        expiresIn: 3600,
      }));

    const site = new SharePointSite({
      owner: 'owner',
      site: 'site',
      clientId: 'clientId',
      tenantId: 'tenantId',
      refreshToken: 'REFRESH_TOKEN',
    });
    const token = await site.getAccessToken();
    assert(token);
  });

  it('retrieves access token twice within expiry', async () => {
    nock('https://login.microsoftonline.com')
      .post('/tenantId/oauth2/v2.0/token')
      .reply(200, () => JSON.stringify({
        access_token: crypto.randomUUID(),
        expires_in: 3600,
      }));

    const site = new SharePointSite({
      owner: 'owner',
      site: 'site',
      clientId: 'clientId',
      tenantId: 'tenantId',
      refreshToken: 'REFRESH_TOKEN',
    });
    const token1 = await site.getAccessToken();
    const token2 = await site.getAccessToken();
    assert.strictEqual(token1, token2);
  });

  it('retrieves access token twice outside expiry', async () => {
    nock('https://login.microsoftonline.com')
      .post('/tenantId/oauth2/v2.0/token')
      .twice()
      .reply(200, () => JSON.stringify({
        access_token: crypto.randomUUID(),
        expires_in: 0, // <== immediately expires
      }));

    const site = new SharePointSite({
      owner: 'owner',
      site: 'site',
      clientId: 'clientId',
      tenantId: 'tenantId',
      refreshToken: 'REFRESH_TOKEN',
    });
    const token1 = await site.getAccessToken();
    const token2 = await site.getAccessToken();
    assert.notStrictEqual(token1, token2);
  });

  it('rejects when getting access token fails', async () => {
    nock('https://login.microsoftonline.com')
      .post('/tenantId/oauth2/v2.0/token')
      .reply(401);

    const site = new SharePointSite({
      owner: 'owner',
      site: 'site',
      clientId: 'clientId',
      tenantId: 'tenantId',
      refreshToken: 'REFRESH_TOKEN',
    });
    await assert.rejects(() => site.getAccessToken());
  });

  it('getFolder succeeds when folder exists', async () => {
    nock('https://login.microsoftonline.com')
      .post('/tenantId/oauth2/v2.0/token')
      .reply(200, () => JSON.stringify({
        access_token: crypto.randomUUID(),
        expiresIn: 3600,
      }));
    nock('https://owner.sharepoint.com')
      .get('/sites/site/_api/web/GetFolderByServerRelativeUrl(\'\')')
      .reply(200, {
        d: {
          Exists: true,
          Name: 'root',
        },
      });
    const site = new SharePointSite({
      owner: 'owner',
      site: 'site',
      clientId: 'clientId',
      tenantId: 'tenantId',
      refreshToken: 'REFRESH_TOKEN',
    });
    const result = await site.getFolder();
    assert.strictEqual(result.d.Exists, true);
  });

  it('getFolder fails when folder does not exist', async () => {
    nock('https://login.microsoftonline.com')
      .post('/tenantId/oauth2/v2.0/token')
      .reply(200, () => JSON.stringify({
        access_token: crypto.randomUUID(),
        expiresIn: 3600,
      }));
    nock('https://owner.sharepoint.com')
      .get('/sites/site/_api/web/GetFolderByServerRelativeUrl(\'/foo\')')
      .reply(404);
    const site = new SharePointSite({
      owner: 'owner',
      site: 'site',
      clientId: 'clientId',
      tenantId: 'tenantId',
      refreshToken: 'REFRESH_TOKEN',
    });
    await assert.rejects(() => site.getFolder('foo'));
  });

  it('getFile succeeds when file exists', async () => {
    nock('https://login.microsoftonline.com')
      .post('/tenantId/oauth2/v2.0/token')
      .reply(200, () => JSON.stringify({
        access_token: crypto.randomUUID(),
        expiresIn: 3600,
      }));
    nock('https://owner.sharepoint.com')
      .get('/sites/site/_api/web/GetFolderByServerRelativeUrl(\'/parent\')/Files(\'file\')?$expand=ModifiedBy')
      .reply(200, {
        d: {
          Exists: true,
          Name: 'file',
        },
      });
    const site = new SharePointSite({
      owner: 'owner',
      site: 'site',
      clientId: 'clientId',
      tenantId: 'tenantId',
      refreshToken: 'REFRESH_TOKEN',
    });
    const result = await site.getFile('parent/file');
    assert.strictEqual(result.d.Exists, true);
  });

  it('getFile fails when file does not exist', async () => {
    nock('https://login.microsoftonline.com')
      .post('/tenantId/oauth2/v2.0/token')
      .reply(200, () => JSON.stringify({
        access_token: crypto.randomUUID(),
        expiresIn: 3600,
      }));
    nock('https://owner.sharepoint.com')
      .get('/sites/site/_api/web/GetFolderByServerRelativeUrl(\'\')/Files(\'file\')?$expand=ModifiedBy')
      .reply(404);
    const site = new SharePointSite({
      owner: 'owner',
      site: 'site',
      clientId: 'clientId',
      tenantId: 'tenantId',
      refreshToken: 'REFRESH_TOKEN',
    });
    await assert.rejects(() => site.getFile('file'));
  });

  it('getFileContents succeeds when file exists', async () => {
    nock('https://login.microsoftonline.com')
      .post('/tenantId/oauth2/v2.0/token')
      .reply(200, () => JSON.stringify({
        access_token: crypto.randomUUID(),
        expiresIn: 3600,
      }));
    nock('https://owner.sharepoint.com')
      .get('/sites/site/_api/web/GetFolderByServerRelativeUrl(\'/parent\')/Files(\'file\')/$value')
      .reply(200, Buffer.from('file-contents'));
    const site = new SharePointSite({
      owner: 'owner',
      site: 'site',
      clientId: 'clientId',
      tenantId: 'tenantId',
      refreshToken: 'REFRESH_TOKEN',
    });
    const result = await site.getFileContents('parent/file');
    assert.strictEqual(result.toString(), 'file-contents');
  });

  it('getFileAndFolders succeeds when folder exists', async () => {
    nock('https://login.microsoftonline.com')
      .post('/tenantId/oauth2/v2.0/token')
      .reply(200, () => JSON.stringify({
        access_token: crypto.randomUUID(),
        expiresIn: 3600,
      }));
    nock('https://owner.sharepoint.com')
      .get('/sites/site/_api/web/GetFolderByServerRelativeUrl(\'/parent\')?$expand=Files/ModifiedBy,Folders')
      .reply(200, {
        d: {
          Folders: [],
          Files: [{
            Exists: true,
            Name: 'file',
          }],
        },
      });
    const site = new SharePointSite({
      owner: 'owner',
      site: 'site',
      clientId: 'clientId',
      tenantId: 'tenantId',
      refreshToken: 'REFRESH_TOKEN',
    });
    const result = await site.getFilesAndFolders('parent');
    assert(result);
  });
});
