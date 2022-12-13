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
import { OneDriveAuth, ACQUIRE_METHODS } from '../src/OneDriveAuth.js';
import { MemCachePlugin } from '../src/cache/MemCachePlugin.js';
import { Nock } from './utils.js';

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

  it('can be disposed.', async () => {
    const auth = new OneDriveAuth({
      clientId: 'foo',
      clientSecret: 'bar',
    });
    await assert.doesNotReject(async () => auth.dispose());
  });

  it('throws when username/password are specified', async () => {
    assert.throws(() => new OneDriveAuth({
      clientId: 'foo',
      clientSecret: 'bar',
      username: 'bob',
      password: 'secret',
    }), Error('Username/password authentication no longer supported.'));
  });

  it('throws when refresh token is specified', async () => {
    assert.throws(() => new OneDriveAuth({
      clientId: 'foo',
      clientSecret: 'bar',
      refreshToken: 'dummy',
    }), Error('Refresh token no longer supported.'));
  });

  it('throws when unknown acquire method is specified', async () => {
    assert.throws(() => new OneDriveAuth({
      clientId: 'foo',
      clientSecret: 'bar',
      acquireMethod: 'magic',
    }), Error('Authentication method unknown: magic, should be none or one of: byDeviceCode,byClientCredential'));
  });

  it('throws when onCode acquire method is specified without onCode parameter', async () => {
    assert.throws(() => new OneDriveAuth({
      clientId: 'foo',
      clientSecret: 'bar',
      acquireMethod: ACQUIRE_METHODS.BY_DEVICE_CODE,
    }), Error('Authontication method byDeviceCode requires \'onCode\' parameter'));
  });

  it('can authenticate against a resource', async () => {
    nock.discovery();
    nock.token({
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
      acquireMethod: ACQUIRE_METHODS.BY_CLIENT_CREDENTIAL,
    });
    const resp = await od.authenticate();
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
      fromNativeBroker: false,
      idToken: '',
      idTokenClaims: {},
      msGraphHost: '',
      requestId: '',
      scopes: [
        'https://graph.microsoft.com/.default',
      ],
      state: '',
      tenantId: '',
      tokenType: 'Bearer',
      uniqueId: '',
    });
  });

  it('can authenticate with device code', async () => {
    nock.discovery();

    nock('https://login.microsoftonline.com')
      .post('/common/oauth2/v2.0/devicecode')
      .reply(200, {
        device_code: 'DAQABAAEAAAD',
        expires_in: 900,
        interval: 5,
        message: 'To sign in, use a web browser to open the page https://microsoft.com/devicelogin and enter the code DTSWBVY27 to authenticate.',
        user_code: 'DTSWBVY27',
        verification_uri: 'https://microsoft.com/devicelogin',
      });

    nock.token({
      token_type: 'Bearer',
      refresh_token: 'dummy',
      access_token: 'dummy',
      expires_in: 81000,
      id_token: new UnsecuredJWT({ sub: 'test' }).encode(),
      client_info: Buffer.from(JSON.stringify({ uid: 'Bob', utid: 'common' })).toString('base64'),
    });

    const od = new OneDriveAuth({
      clientId: '83ab2922-5f11-4e4d-96f3-d1e0ff152856',
      clientSecret: 'test-client-secret',
      resource: 'test-resource',
      tenant: 'common',
      onCode: async (code) => {
        assert.strictEqual(code.userCode, 'DTSWBVY27');
      },
      acquireMethod: ACQUIRE_METHODS.BY_DEVICE_CODE,
    });
    await od.authenticate();
    assert.strictEqual(await od.isAuthenticated(), true);
  });

  it('catches 401 errors when authenticating', async () => {
    nock.discovery();
    nock.unauthenticated();

    const od = new OneDriveAuth({
      clientId: '83ab2922-5f11-4e4d-96f3-d1e0ff152856',
      clientSecret: 'test-client-secret',
      resource: 'test-resource',
      tenant: 'common',
      acquireMethod: ACQUIRE_METHODS.BY_CLIENT_CREDENTIAL,
    });
    await assert.rejects(
      od.authenticate(false),
      /Invalid client secret/,
    );
  });

  it('returns null when silent acquire fails', async () => {
    nock.discovery();
    nock('https://login.microsoftonline.com')
      .post('/common/oauth2/v2.0/devicecode')
      .reply(200, {
        device_code: 'DAQABAAEAAAD',
        expires_in: 900,
        interval: 5,
        message: 'To sign in, use a web browser to open the page https://microsoft.com/devicelogin and enter the code DTSWBVY27 to authenticate.',
        user_code: 'DTSWBVY27',
        verification_uri: 'https://microsoft.com/devicelogin',
      });
    nock.token({
      token_type: 'Bearer',
      refresh_token: 'dummy',
      access_token: 'dummy',
      expires_in: 81000,
      id_token: new UnsecuredJWT({ sub: 'test' }).encode(),
      client_info: Buffer.from(JSON.stringify({ uid: 'Bob', utid: 'common' })).toString('base64'),
    });

    const caches = new Map();
    const od1 = new OneDriveAuth({
      clientId: '83ab2922-5f11-4e4d-96f3-d1e0ff152856',
      clientSecret: 'test-client-secret',
      resource: 'test-resource',
      tenant: 'common',
      onCode: async (code) => {
        assert.strictEqual(code.userCode, 'DTSWBVY27');
      },
      cachePlugin: new MemCachePlugin({
        key: 'default',
        caches,
      }),
      acquireMethod: ACQUIRE_METHODS.BY_DEVICE_CODE,
    });
    await od1.authenticate();

    // make the cached access token expire
    const entry = JSON.parse(caches.get('default'));
    entry.AccessToken[Object.keys(entry.AccessToken)[0]].expires_on = Math.floor(Date.now() / 1000);
    caches.set('default', JSON.stringify(entry));

    nock.discovery();
    nock.unauthenticated();

    const od2 = new OneDriveAuth({
      clientId: '83ab2922-5f11-4e4d-96f3-d1e0ff152856',
      clientSecret: 'test-client-secret',
      resource: 'test-resource',
      tenant: 'common',
      cachePlugin: new MemCachePlugin({
        key: 'default',
        caches,
      }),
    });
    assert.strictEqual(await od2.authenticate(true), null);

    nock.discovery();
    nock.unauthenticated();

    const od3 = new OneDriveAuth({
      clientId: '83ab2922-5f11-4e4d-96f3-d1e0ff152856',
      clientSecret: 'test-client-secret',
      resource: 'test-resource',
      tenant: 'common',
      cachePlugin: new MemCachePlugin({
        key: 'default',
        caches,
      }),
    });
    await assert.rejects(
      async () => od3.authenticate(false),
      Error('Unable to acquire token silently and no other acquire method supplied'),
    );
  });

  it('uses the tenant from a mountpoint', async () => {
    const od = new OneDriveAuth({
      clientId: 'foobar',
      localAuthCache: true,
    });
    await od.initTenantFromMountPoint({
      tenantId: 'c0452eed-9384-4001-b1b1-71b3d5cf28ad',
    });
    assert.deepStrictEqual(od.tenant, 'c0452eed-9384-4001-b1b1-71b3d5cf28ad');
  });

  it('resolves the tenant from a share link and caches it', async () => {
    nock(AZ_AUTHORITY_HOST_URL)
      .get('/somedrive.onmicrosoft.com/.well-known/openid-configuration')
      .reply(200, {
        issuer: 'https://sts.windows.net/c0452eed-9384-4001-b1b1-71b3d5cf28ad/',
      });

    const tenantCache = new Map();
    const od1 = new OneDriveAuth({
      clientId: 'foobar',
      localAuthCache: true,
      tenantCache,
    });
    await od1.initTenantFromMountPoint({
      url: 'https://somedrive.com/a/b/c/d2',
    });

    const od2 = new OneDriveAuth({
      clientId: 'foobar',
      localAuthCache: true,
      tenantCache,
    });
    await od2.initTenantFromUrl('https://somedrive.com/a/b/c/d2');

    assert.deepStrictEqual(Object.fromEntries(tenantCache.entries()), {
      somedrive: 'c0452eed-9384-4001-b1b1-71b3d5cf28ad',
    });
  });

  it('returns common tenant if resolving the tenant fails', async () => {
    nock(AZ_AUTHORITY_HOST_URL)
      .get('/somedrive.onmicrosoft.com/.well-known/openid-configuration')
      .reply(404);

    const tenantCache = new Map();
    const od = new OneDriveAuth({
      clientId: 'foobar',
      localAuthCache: true,
      tenantCache,
    });
    await od.initTenantFromMountPoint({
      url: 'https://somedrive.com/a/b/c/d2',
    });

    assert.deepStrictEqual(Object.fromEntries(tenantCache.entries()), {
      somedrive: 'common',
    });
  });

  it('returns common tenant if resolving the tenant returns no issuer', async () => {
    nock(AZ_AUTHORITY_HOST_URL)
      .get('/somedrive.onmicrosoft.com/.well-known/openid-configuration')
      .reply(200, {});

    const tenantCache = new Map();
    const od = new OneDriveAuth({
      clientId: 'foobar',
      localAuthCache: true,
      tenantCache,
    });
    await od.initTenantFromMountPoint({
      url: 'https://somedrive.com/a/b/c/d2',
    });

    assert.deepStrictEqual(Object.fromEntries(tenantCache.entries()), {
      somedrive: 'common',
    });
  });

  it('resolves the onedrive.live.com urls', async () => {
    const od = new OneDriveAuth({
      clientId: 'foobar',
      localAuthCache: true,
    });
    await od.initTenantFromMountPoint({
      url: 'https://onedrive.live.com/a/b/c/d2',
    });
    assert.strictEqual(od.tenant, 'common');
  });

  it('resolves the 1drv.ms urls', async () => {
    const od = new OneDriveAuth({
      clientId: 'foobar',
      localAuthCache: true,
    });
    await od.initTenantFromMountPoint({
      url: 'https://1drv.ms/a/b/c/d2',
    });
    assert.strictEqual(od.tenant, 'common');
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
      localAuthCache: true,
      tenantCache,
    });
    await od1.initTenantFromUrl(new URL('https://adobe-my.sharepoint.com/a/b/c/d2'));
    await od1.initTenantFromUrl(new URL('https://adobe-my.sharepoint.com/a/b/c/d2'));

    assert.deepStrictEqual(Object.fromEntries(tenantCache.entries()), {
      adobe: 'c0452eed-9384-4001-b1b1-71b3d5cf28ad',
    });
  });

  it('resolves the tenant from a share link and ignores cache', async () => {
    nock(AZ_AUTHORITY_HOST_URL)
      .get('/sonedrive.onmicrosoft.com/.well-known/openid-configuration')
      .twice()
      .reply(200, {
        issuer: 'https://sts.windows.net/c0452eed-9384-4001-b1b1-71b3d5cf28ad/',
      });

    const od1 = new OneDriveAuth({
      clientId: 'foobar',
      localAuthCache: true,
      noTenantCache: true,
    });
    await od1.initTenantFromUrl('https://sonedrive.com/a/b/c/d2');
    delete od1.tenant;
    await od1.initTenantFromUrl('https://sonedrive.com/a/b/c/d2');

    // this should not fetch it again
    await od1.initTenantFromMountPoint({
      url: 'https://onedrive.com/a/b/c/d2',
    });
  });

  it('sets the access token an extract the tenant', async () => {
    const bearerToken = new UnsecuredJWT({
      email: 'bob',
      tid: 'test-tenantid',
    })
      .encode();

    const od = new OneDriveAuth({
      clientId: 'foobar',
      localAuthCache: true,
      noTenantCache: true,
    });
    od.setAccessToken(bearerToken);

    const accessToken = await od.authenticate();
    assert.strictEqual(accessToken.accessToken, bearerToken);
    assert.strictEqual(accessToken.tenantId, 'test-tenantid');
    assert.strictEqual(od.tenant, 'test-tenantid');
  });

  it('getAuthorityUrl without tenant resolution throws', async () => {
    const od = new OneDriveAuth({
      clientId: 'foobar',
      localAuthCache: true,
      noTenantCache: true,
    });
    assert.throws(() => od.getAuthorityUrl());
  });

  it('setAccessToken warns when token is invalid', async () => {
    const od = new OneDriveAuth({
      clientId: 'foobar',
      localAuthCache: true,
      noTenantCache: true,
    });
    od.setAccessToken('test');
  });

  it('authenticate returns null when no accounts are there in silent mode', async () => {
    const od = new OneDriveAuth({
      clientId: '83ab2922-5f11-4e4d-96f3-d1e0ff152856',
      clientSecret: 'test-client-secret',
      resource: 'test-resource',
      tenant: 'common',
    });

    const token = await od.authenticate(true);
    assert.strictEqual(token, null);
  });
});
