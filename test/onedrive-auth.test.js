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

  it('resolves the tenant from a share link and caches it', async () => {
    nock(AZ_AUTHORITY_HOST_URL)
      .get('/onedrive.onmicrosoft.com/.well-known/openid-configuration')
      .reply(200, {
        issuer: 'https://sts.windows.net/c0452eed-9384-4001-b1b1-71b3d5cf28ad/',
      });

    const tenantCache = new Map();
    const od1 = new OneDriveAuth({
      clientId: 'foobar',
      refreshToken: 'dummy',
      localAuthCache: true,
      tenantCache,
    });
    await od1.initTenantFromUrl('https://onedrive.com/a/b/c/d2');

    const od2 = new OneDriveAuth({
      clientId: 'foobar',
      refreshToken: 'dummy',
      localAuthCache: true,
      tenantCache,
    });
    await od2.initTenantFromUrl('https://onedrive.com/a/b/c/d2');

    assert.deepStrictEqual(Object.fromEntries(tenantCache.entries()), {
      onedrive: 'c0452eed-9384-4001-b1b1-71b3d5cf28ad',
    });
  });

  it('resolves the tenant from a sharepoint share link and caches it', async () => {
    nock(AZ_AUTHORITY_HOST_URL)
      .get('/adobe.onmicrosoft.com/.well-known/openid-configuration')
      .reply(200, {
        issuer: 'https://sts.windows.net/c0452eed-9384-4001-b1b1-71b3d5cf28ad/',
      });

    const tenantCache = new Map();
    const od1 = new OneDriveAuth({
      clientId: 'foobar',
      refreshToken: 'dummy',
      localAuthCache: true,
      tenantCache,
    });
    await od1.initTenantFromUrl(new URL('https://adobe-my.sharepoint.com/a/b/c/d2'));
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
      });

    const od1 = new OneDriveAuth({
      clientId: 'foobar',
      refreshToken: 'dummy',
      localAuthCache: true,
      noTenantCache: true,
    });
    await od1.initTenantFromUrl('https://onedrive.com/a/b/c/d2');
    delete od1.tenant;
    await od1.initTenantFromUrl('https://onedrive.com/a/b/c/d2');
  });

  it('sets the access token an extract the tenant', async () => {
    const bearerToken = new jose.UnsecuredJWT({
      email: 'bob',
      tid: 'test-tenantid',
    })
      .encode();

    const od = new OneDriveAuth({
      clientId: 'foobar',
      refreshToken: 'dummy',
      localAuthCache: true,
      noTenantCache: true,
    });
    od.setAccessToken(bearerToken);

    const accessToken = await od.getAccessToken();
    assert.strictEqual(accessToken.accessToken, bearerToken);
    assert.strictEqual(accessToken.tenantId, 'test-tenantid');
    assert.strictEqual(od.tenant, 'test-tenantid');
  });
});
