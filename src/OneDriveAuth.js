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

// eslint-disable-next-line max-classes-per-file
import { keepAliveNoCache } from '@adobe/fetch';
import { ConfidentialClientApplication, LogLevel, PublicClientApplication } from '@azure/msal-node';
import { MemCachePlugin } from '@adobe/helix-shared-tokencache';
import { decodeJwt } from 'jose';

const AZ_AUTHORITY_HOST_URL = 'https://login.windows.net';
const AZ_COMMON_TENANT = 'common';

const DEFAULT_SCOPES = ['https://graph.microsoft.com/.default', 'openid', 'profile', 'offline_access'];

const MSAL_LOG_LEVELS = [
  'error',
  'warn',
  'info',
  'debug',
  'trace',
];

export const AcquireMethod = {
  BY_DEVICE_CODE: 'byDeviceCode',
  BY_CLIENT_CREDENTIAL: 'byClientCredential',
};

/**
 * aliases
 * @typedef {import('@azure/msal-node').AuthenticationResult} AuthenticationResult
*/

/**
 * map that caches the tenant ids
 * @type {Map<string, string>}
 */
const globalTenantCache = new Map();

/**
 * Helper class that facilitates accessing one drive.
 *
 * @class
 * @field {ConfidentialClientApplication|PublicClientApplication} app
 */
export class OneDriveAuth {
  /**
   * @param {OneDriveAuthOptions} opts Options
   */
  constructor(opts) {
    this.fetchContext = keepAliveNoCache({ userAgent: 'adobe-fetch' });

    if (!opts.clientId && !opts.accessToken) {
      throw new Error('Either clientId or accessToken must not be null.');
    }
    if (opts.username || opts.password) {
      throw new Error('Username/password authentication no longer supported.');
    }
    if (opts.refreshToken) {
      throw new Error('Refresh token no longer supported.');
    }

    this.clientId = opts.clientId;
    this.clientSecret = opts.clientSecret || '';
    this._log = opts.log || console;
    this.tenant = opts.tenant;
    this.cachePlugin = opts.cachePlugin;
    this.scopes = opts.scopes || DEFAULT_SCOPES;
    this.onCode = opts.onCode;
    this.acquireMethod = opts.acquireMethod || '';
    this.logFields = Object.entries(opts.logFields || {})
      .map(([key, value]) => `[${key}:${value}]`).join(' ');

    const validAcquireMethods = Array.from(Object.values(AcquireMethod));
    if (this.acquireMethod && !validAcquireMethods.includes(this.acquireMethod)) {
      throw new Error(`Authentication method unknown: ${this.acquireMethod}, should be none or one of: ${validAcquireMethods}`);
    }
    if (this.acquireMethod === AcquireMethod.BY_DEVICE_CODE && !this.onCode) {
      throw new Error(`Authentication method ${AcquireMethod.BY_DEVICE_CODE} requires 'onCode' parameter`);
    }
    if (!this.acquireMethod && this.onCode) {
      this.acquireMethod = AcquireMethod.BY_DEVICE_CODE;
    }

    if (!opts.noTenantCache && !process.env.HELIX_ONEDRIVE_NO_TENANT_CACHE) {
      /** @type {Map<string, string>} */
      this.tenantCache = opts.tenantCache || globalTenantCache;
    }

    if ((opts.localAuthCache || process.env.HELIX_ONEDRIVE_LOCAL_AUTH_CACHE) && !this.cachePlugin) {
      this.cachePlugin = new MemCachePlugin({
        log: this._log,
        key: 'default',
        caches: new Map(),
      });
    }

    if (opts.accessToken) {
      this.setAccessToken(opts.accessToken);
    }
  }

  get app() {
    if (!this._app) {
      const {
        log,
        cachePlugin,
      } = this;
      const msalConfig = {
        auth: {
          clientId: this.clientId,
          clientSecret: this.clientSecret,
          authority: this.getAuthorityUrl(),
        },
        system: {
          loggerOptions: {
            loggerCallback(loglevel, message) {
              log[MSAL_LOG_LEVELS[loglevel]](message);
            },
            piiLoggingEnabled: false,
            logLevel: LogLevel.Info,
          },
        },
      };
      if (cachePlugin) {
        msalConfig.cache = {
          cachePlugin,
        };
      }
      this._app = this.acquireMethod === AcquireMethod.BY_DEVICE_CODE
        ? new PublicClientApplication(msalConfig)
        : new ConfidentialClientApplication(msalConfig);
    }
    return this._app;
  }

  async resolveTenant(tenantHostName) {
    const { log } = this;
    if (tenantHostName === 'onedrive' || tenantHostName === '1drv') {
      log.info(`forcing 'common' tenant for '${tenantHostName}'.`);
      return AZ_COMMON_TENANT;
    }
    const configUrl = `https://login.windows.net/${tenantHostName}.onmicrosoft.com/.well-known/openid-configuration`;
    const { fetch } = this.fetchContext;
    const res = await fetch(configUrl);
    if (!res.ok) {
      log.info(`error fetching openid-configuration for ${tenantHostName}: ${res.status}. Fallback to 'common'`);
      return AZ_COMMON_TENANT;
    }

    const { issuer } = await res.json();
    if (!issuer) {
      log.info(`unable to extract tenant from openid-configuration for ${tenantHostName}: no 'issuer'. Fallback to 'common'`);
      return AZ_COMMON_TENANT;
    }

    // eslint-disable-next-line prefer-destructuring
    const tenant = new URL(issuer).pathname.split('/')[1];
    log.info(`fetched tenant information from for ${tenantHostName}: ${tenant}`);
    return tenant;
  }

  static getTenantHostFromUrl(sharingUrl) {
    const url = sharingUrl instanceof URL
      ? sharingUrl
      : new URL(sharingUrl);
    let [tenantHost] = url.hostname.split('.');
    // special case: `xxxx-my.sharepoint.com`
    if (url.hostname.endsWith('-my.sharepoint.com')) {
      tenantHost = tenantHost.substring(0, tenantHost.length - 3);
    }
    return tenantHost;
  }

  async initTenantFromUrl(sharingUrl) {
    if (this.tenant) {
      return this.tenant;
    }
    const { log } = this;
    const tenantHost = OneDriveAuth.getTenantHostFromUrl(sharingUrl);

    if (this.tenantCache) {
      this.tenant = this.tenantCache.get(tenantHost);
    }
    if (!this.tenant) {
      this.tenant = await this.resolveTenant(tenantHost);
      if (this.tenantCache) {
        this.tenantCache.set(tenantHost, this.tenant);
      }
    }
    log.info(`using tenant ${this.tenant} for ${tenantHost} from ${sharingUrl}`);
    return this.tenant;
  }

  async initTenantFromMountPoint(mp) {
    const { log } = this;
    if (this.tenant) {
      return this.tenant;
    }
    if (mp.tenantId) {
      this.tenant = mp.tenantId;
      log.info(`using tenant ${this.tenant} from fstab.`);
      return this.tenant;
    }
    return this.initTenantFromUrl(mp.url);
  }

  /**
   */
  // eslint-disable-next-line class-methods-use-this
  async dispose() {
    // TODO: clear other state?
    const { reset } = this.fetchContext;
    return reset();
  }

  /**
   */
  get log() {
    return this._log;
  }

  getAuthorityUrl() {
    if (!this.tenant) {
      throw new Error('unable to compute authority url. no tenant.');
    }
    return `${AZ_AUTHORITY_HOST_URL}/${this.tenant}`;
  }

  /**
   * @returns {boolean}
   */
  async isAuthenticated() {
    const accounts = await this.app.getTokenCache().getAllAccounts();
    return accounts.length > 0;
  }

  /**
   * Sets the access token to use for all requests. if the token is a valid JWT token,
   * its `tid` claim is used a tenant (if no tenant is already set).
   *
   * @param {string} bearerToken
   * @returns {OneDriveAuth} this
   */
  setAccessToken(bearerToken) {
    const { log } = this;
    /** @type AuthenticationResult */
    this.authResult = {
      accessToken: bearerToken,
    };
    if (!this.tenant) {
      try {
        const { tid } = decodeJwt(bearerToken);
        if (tid) {
          log.info(`using tenant from access token: ${tid}`);
          this.tenant = tid;
        }
      } catch (e) {
        log.warn(`unable to decode access token: ${e.message}`);
      }
    }
    this.authResult.tenantId = this.tenant;
    return this;
  }

  handleAcquireError(account, e, forced = false) {
    const { log } = this;
    const msg = `Error while reacquiring token from cache${forced ? ' (forced)' : ''}.`;

    log.warn(`${msg}\nUsername: ${account.username}\nAuth-Location: ${this.cachePlugin.location}\nMessage: ${e.message}`);
  }

  /**
   * Authenticates against MS
   * @param {boolean} silentOnly
   * @returns {Promise<null|AuthenticationResult>}
   */
  async doAuthenticate(silentOnly) {
    const { log, app } = this;

    let accounts = await app.getTokenCache().getAllAccounts();
    if (accounts.length > 0) {
      let account = accounts[0];

      try {
        return await app.acquireTokenSilent({ account });
      } catch (e) {
        this.handleAcquireError(account, e);
      }

      // try again with fresh mem cache
      if (this.cachePlugin instanceof MemCachePlugin) {
        this.cachePlugin.clear();

        accounts = await app.getTokenCache().getAllAccounts();
        if (accounts.length > 0) {
          [account] = accounts;

          try {
            return await app.acquireTokenSilent({
              forceRefresh: true,
              account,
            });
          } catch (e) {
            this.handleAcquireError(account, e, true);
          }
        }
      }
    }

    if (silentOnly) {
      return null;
    }

    try {
      if (this.acquireMethod === AcquireMethod.BY_DEVICE_CODE) {
        log.debug('acquire token with device.');
        return await app.acquireTokenByDeviceCode({
          deviceCodeCallback: async (code) => {
            await this.onCode(code);
          },
          scopes: this.scopes,
        });
      }
      if (this.acquireMethod === AcquireMethod.BY_CLIENT_CREDENTIAL
          // check if plugin wants us to use client credentials
          || (await this.cachePlugin.getPluginMetadata() || {}).useClientCredentials) {
        log.debug('acquire token with client credentials.');
        return await app.acquireTokenByClientCredential({
          scopes: this.scopes,
        });
      }
    } catch (e) {
      log.error(`Error while acquiring access token (${this.acquireMethod}).\nMessage: ${e.message}`);
      throw e;
    }

    throw new Error('Unable to acquire token silently and no other acquire method supplied');
  }

  /**
   * Authenticates by either using the cached result or talking to MS
   * @param {boolean} silentOnly
   * @returns {Promise<AuthenticationResult>}
   */
  async authenticate(silentOnly) {
    if (!this.authResult) {
      this.authResult = await this.doAuthenticate(silentOnly);
    }
    return this.authResult;
  }
}
