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

const assert = require('assert');
const jose = require('jose');
const { OneDriveAuth } = require('../src/OneDriveAuth.js');
const { Nock } = require('./utils.js');

const AZ_AUTHORITY_HOST_URL = 'https://login.windows.net';

const { TEST_CLIENT_ID, TEST_USER, TEST_PASSWORD } = process.env;

const DEFAULT_AUTH = {
  token_type: 'Bearer',
  refresh_token: 'dummy',
  access_token: 'dummy',
  expires_in: 81000,
};

describe('OneDriveAuth Tests', () => {
  let nock;
  beforeEach(() => {
    nock = new Nock();
    delete process.env.HELIX_ONEDRIVE_NO_SHARE_LINK_CACHE;
  });

  afterEach(() => {
    nock.done();
  });

  it('throws when required parameters are not specified.', async () => {
    assert.throws(() => new OneDriveAuth({}), Error('Missing clientId.'));
  });

  it('can be constructed.', async () => {
    const auth = new OneDriveAuth({
      clientId: 'foo',
      clientSecret: 'bar',
    });
    assert.ok(auth);
  });

  it('can authenticate against a resource', async () => {
    nock('https://login.microsoftonline.com')
      .get('/common/discovery/instance?api-version=1.1&authorization_endpoint=https://login.windows.net/common/oauth2/v2.0/authorize')
      .reply(200, {
        tenant_discovery_endpoint: 'https://login.windows.net/common/v2.0/.well-known/openid-configuration',
        'api-version': '1.1',
        metadata: [
          {
            preferred_network: 'login.microsoftonline.com',
            preferred_cache: 'login.windows.net',
            aliases: [
              'login.microsoftonline.com',
              'login.windows.net',
            ],
          },
        ],
      })
      .get('/common/v2.0/.well-known/openid-configuration')
      .reply(200, {
        token_endpoint: 'https://login.microsoftonline.com/common/oauth2/v2.0/token',
        issuer: 'https://login.microsoftonline.com/{tenantid}/v2.0',
        authorization_endpoint: 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize',
      })
      .post('/common/oauth2/v2.0/token')
      .reply(200, {
        token_type: 'Bearer',
        refresh_token: 'dummy',
        access_token: 'dummy',
        expires_in: 81000,
      });

    const od = new OneDriveAuth({
      clientId: '83ab2922-5f11-4e4d-96f3-d1e0ff152856',
      clientSecret: 'test-client-secret',
      resource: 'test-resource',
      tenant: 'common',
    });
    const resp = await od.getAccessToken();
    delete resp.expiresOn;
    delete resp.extExpiresOn;
    delete resp.correlationId;
    assert.deepStrictEqual(resp, {
      accessToken: 'dummy',
      account: null,
      authority: 'https://login.microsoftonline.com/common/',
      cloudGraphHostName: '',
      code: undefined,
      familyId: '',
      fromCache: false,
      idToken: '',
      idTokenClaims: {},
      msGraphHost: '',
      scopes: [
        'https://graph.microsoft.com/.default',
        'openid',
        'profile',
        'offline_access',
      ],
      state: '',
      tenantId: '',
      tokenType: 'Bearer',
      uniqueId: '',
    });
  });

  it('uses global share link cache by default', async () => {
    nock.loginWindowsNet();
    nock('https://graph.microsoft.com/v1.0')
      .get('/shares/u!aHR0cHM6Ly9vbmVkcml2ZS5jb20vYS9iL2MvZDA/driveItem')
      .reply(200, {
        id: 'some-id',
      });

    const od = new OneDrive({
      clientId: 'foobar',
      refreshToken: 'dummy',
      localAuthCache: true,
      tenant: 'common',
    });
    const item1 = await od.getDriveItemFromShareLink('https://onedrive.com/a/b/c/d0');
    assert.deepStrictEqual(item1, {
      id: 'some-id',
    });
    const item2 = await od.getDriveItemFromShareLink('https://onedrive.com/a/b/c/d0');
    assert.strictEqual(item1, item2);
  });

  it('share link cache can be disabled via option', async () => {
    nock(AZ_AUTHORITY_HOST_URL)
      .post('/common/oauth2/token?api-version=1.0')
      .reply(200, DEFAULT_AUTH);
    nock('https://graph.microsoft.com/v1.0')
      .get('/shares/u!aHR0cHM6Ly9vbmVkcml2ZS5jb20vYS9iL2MvZDAx/driveItem')
      .twice()
      .reply(200, {
        id: 'some-id',
      });

    const od = new OneDrive({
      clientId: 'foobar',
      refreshToken: 'dummy',
      localAuthCache: true,
      noShareLinkCache: true,
      tenant: 'common',
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
    nock(AZ_AUTHORITY_HOST_URL)
      .post('/common/oauth2/token?api-version=1.0')
      .reply(200, DEFAULT_AUTH);
    nock('https://graph.microsoft.com/v1.0')
      .get('/shares/u!aHR0cHM6Ly9vbmVkcml2ZS5jb20vYS9iL2MvZDE/driveItem')
      .twice()
      .reply(200, {
        id: 'some-id',
      });

    const od = new OneDrive({
      clientId: 'foobar',
      refreshToken: 'dummy',
      localAuthCache: true,
      tenant: 'common',
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
    nock.loginWindowsNet();
    nock('https://graph.microsoft.com/v1.0')
      .get('/shares/u!aHR0cHM6Ly9vbmVkcml2ZS5jb20vYS9iL2MvZDI/driveItem')
      .reply(200, {
        id: 'some-id',
      });

    const shareLinkCache = new Map();

    const od = new OneDrive({
      clientId: 'foobar',
      refreshToken: 'dummy',
      localAuthCache: true,
      shareLinkCache,
      tenant: 'common',
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

  it('resolves the tenant from a share link and caches it', async () => {
    nock(AZ_AUTHORITY_HOST_URL)
      .get('/onedrive.onmicrosoft.com/.well-known/openid-configuration')
      .reply(200, {
        issuer: 'https://sts.windows.net/c0452eed-9384-4001-b1b1-71b3d5cf28ad/',
      })
      .post('/c0452eed-9384-4001-b1b1-71b3d5cf28ad/oauth2/token?api-version=1.0')
      .twice()
      .reply(200, {
        token_type: 'Bearer',
        refresh_token: 'dummy',
        access_token: 'dummy',
        expires_in: 81000,
      });
    nock('https://graph.microsoft.com/v1.0')
      .get('/shares/u!aHR0cHM6Ly9vbmVkcml2ZS5jb20vYS9iL2MvZDI/driveItem')
      .twice()
      .reply(200, {
        id: 'some-id',
      });

    const tenantCache = new Map();
    const od1 = new OneDrive({
      clientId: 'foobar',
      refreshToken: 'dummy',
      localAuthCache: true,
      tenantCache,
      noShareLinkCache: true,
    });
    const item1 = await od1.getDriveItemFromShareLink('https://onedrive.com/a/b/c/d2');
    assert.deepStrictEqual(item1, {
      id: 'some-id',
    });

    const od2 = new OneDrive({
      clientId: 'foobar',
      refreshToken: 'dummy',
      localAuthCache: true,
      tenantCache,
      noShareLinkCache: true,
    });
    const item2 = await od2.getDriveItemFromShareLink('https://onedrive.com/a/b/c/d2');
    assert.deepStrictEqual(item1, item2);

    assert.deepStrictEqual(Object.fromEntries(tenantCache.entries()), {
      onedrive: 'c0452eed-9384-4001-b1b1-71b3d5cf28ad',
    });
  });

  it('resolves the tenant from a sharepoint share link and caches it', async () => {
    nock(AZ_AUTHORITY_HOST_URL)
      .get('/adobe.onmicrosoft.com/.well-known/openid-configuration')
      .reply(200, {
        issuer: 'https://sts.windows.net/c0452eed-9384-4001-b1b1-71b3d5cf28ad/',
      })
      .post('/c0452eed-9384-4001-b1b1-71b3d5cf28ad/oauth2/token?api-version=1.0')
      .reply(200, {
        token_type: 'Bearer',
        refresh_token: 'dummy',
        access_token: 'dummy',
        expires_in: 81000,
      });
    nock('https://graph.microsoft.com/v1.0')
      .get('/shares/u!aHR0cHM6Ly9hZG9iZS1teS5zaGFyZXBvaW50LmNvbS9hL2IvYy9kMg=/driveItem')
      .reply(200, {
        id: 'some-id',
      });

    const tenantCache = new Map();
    const od1 = new OneDrive({
      clientId: 'foobar',
      refreshToken: 'dummy',
      localAuthCache: true,
      tenantCache,
      noShareLinkCache: true,
    });
    const item1 = await od1.getDriveItemFromShareLink(new URL('https://adobe-my.sharepoint.com/a/b/c/d2'));
    assert.deepStrictEqual(item1, {
      id: 'some-id',
    });
    assert.deepStrictEqual(Object.fromEntries(tenantCache.entries()), {
      adobe: 'c0452eed-9384-4001-b1b1-71b3d5cf28ad',
    });
  });

  it('resolves the tenant from a share link and ignores cache', async () => {
    nock(AZ_AUTHORITY_HOST_URL)
      .get('/onedrive.onmicrosoft.com/.well-known/openid-configuration')
      .twice()
      .reply(200, {
        issuer: 'https://sts.windows.net/c0452eed-9384-4001-b1b1-71b3d5cf28ad/',
      })
      .post('/c0452eed-9384-4001-b1b1-71b3d5cf28ad/oauth2/token?api-version=1.0')
      .twice()
      .reply(200, {
        token_type: 'Bearer',
        refresh_token: 'dummy',
        access_token: 'dummy',
        expires_in: 81000,
      });
    nock('https://graph.microsoft.com/v1.0')
      .get('/shares/u!aHR0cHM6Ly9vbmVkcml2ZS5jb20vYS9iL2MvZDI/driveItem')
      .twice()
      .reply(200, {
        id: 'some-id',
      });

    const od1 = new OneDrive({
      clientId: 'foobar',
      refreshToken: 'dummy',
      localAuthCache: true,
      noTenantCache: true,
      noShareLinkCache: true,
    });
    const item1 = await od1.getDriveItemFromShareLink('https://onedrive.com/a/b/c/d2');
    assert.deepStrictEqual(item1, {
      id: 'some-id',
    });
    const od2 = new OneDrive({
      clientId: 'foobar',
      refreshToken: 'dummy',
      localAuthCache: true,
      noTenantCache: true,
      noShareLinkCache: true,
    });
    const item2 = await od2.getDriveItemFromShareLink('https://onedrive.com/a/b/c/d2');
    assert.deepStrictEqual(item1, item2);
  });

  it('sets the access token an extract the tenant', async () => {
    const keyPair = await jose.generateKeyPair('RS256');
    const bearerToken = await new jose.SignJWT({
      email: 'bob',
      name: 'Bob',
      userId: '112233',
      tid: 'test-tenantid',
    })
      .setProtectedHeader({ alg: 'RS256' })
      .setIssuedAt()
      .setIssuer('urn:example:issuer')
      .setAudience('dummy-clientid')
      .setExpirationTime('2h')
      .sign(keyPair.privateKey);

    const od = new OneDrive({
      clientId: 'foobar',
      refreshToken: 'dummy',
      localAuthCache: true,
      noTenantCache: true,
      noShareLinkCache: true,
    });
    od.setAccessToken(bearerToken);

    const accessToken = await od.getAccessToken();
    assert.strictEqual(accessToken.accessToken, bearerToken);
    assert.strictEqual(accessToken.tenantId, 'test-tenantid');
    assert.strictEqual(od.tenant, 'test-tenantid');
  });

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
      tenant: 'common',
    });
    await assert.rejects(od.me(), new StatusCodeError('wrong input', 400));
  }).timeout(5000);

  it('uploadDriveItem fails with bad conflict behaviour', async () => {
    const drive = new OneDrive({
      clientId: 'foo', clientSecret: 'bar',
    });
    await assert.rejects(
      async () => drive.uploadDriveItem(Buffer.alloc(0), 'item', '', 'guess'),
      /Error: Bad confict behaviour/,
    );
  });
});
