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
import { keepAliveNoCache } from '@adobe/helix-fetch';
import { ConfidentialClientApplication, LogLevel } from '@azure/msal-node';
import { decodeJwt } from 'jose';
import { MemCachePlugin } from './cache/MemCachePlugin.js';
import { StatusCodeError } from './StatusCodeError.js';

const { fetch, reset } = keepAliveNoCache({ userAgent: 'helix-fetch' });

const AZ_AUTHORITY_HOST_URL = 'https://login.windows.net';
const AZ_DEFAULT_RESOURCE = 'https://graph.microsoft.com'; // '00000002-0000-0000-c000-000000000000'; ??
const AZ_COMMON_TENANT = 'common';

const DEFAULT_SCOPES = ['https://graph.microsoft.com/.default', 'openid', 'profile', 'offline_access'];

const MSAL_LOG_LEVELS = [
  'error',
  'warn',
  'info',
  'debug',
  'trace',
];

/**
 * aliases
 * @typedef {import("@azure/msal-node").AuthenticationResult} AuthenticationResult
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
    if (!opts.clientId) {
      throw new Error('Missing clientId.');
    }
    if (opts.username || opts.password) {
      throw new Error('Username/password authentication no longer support.');
    }

    this.clientId = opts.clientId;
    this.clientSecret = opts.clientSecret || '';
    this.refreshToken = opts.refreshToken || '';
    this._log = opts.log || console;
    this.tenant = opts.tenant;
    this.resource = opts.resource || AZ_DEFAULT_RESOURCE;
    this.cachePlugin = opts.cachePlugin;
    this.scopes = opts.scopes || DEFAULT_SCOPES;

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
            logLevel: LogLevel.Verbose,
          },
        },
      };

      if (cachePlugin) {
        msalConfig.cache = {
          cachePlugin,
        };
      }
      this._app = new ConfidentialClientApplication(msalConfig);
    }
    return this._app;
  }

  async resolveTenant(tenantHostName) {
    const { log } = this;
    const configUrl = `https://login.windows.net/${tenantHostName}.onmicrosoft.com/.well-known/openid-configuration`;
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
      return;
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
  }

  /**
   */
  // eslint-disable-next-line class-methods-use-this
  async dispose() {
    // TODO: clear other state?
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
   * Performs a login using an interactive flow which prompts the user to open a browser window and
   * enter the authorization code.
   * @params {function} [onCode] - optional function that gets invoked after code was retrieved.
   * @returns {Promise<AuthenticationResult>}
   */
  async acquireTokenByDeviceCode(onCode) {
    const { log, app } = this;
    try {
      return await app.acquireTokenByDeviceCode({
        deviceCodeCallback: async (code) => {
          log.info(code.message);
          if (typeof onCode === 'function') {
            await onCode(code);
          }
        },
        scopes: this.scopes,
      });
    } catch (e) {
      log.error('Error while requesting access token with device code', e);
      throw e;
    }
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

  /**
   * Authenticates against MS
   * @param {boolean} silentOnly
   * @returns {Promise<null|AuthenticationResult>}
   */
  async doAuthenticate(silentOnly) {
    const { log, app } = this;
    const accounts = await app.getTokenCache().getAllAccounts();
    if (accounts.length > 0) {
      try {
        return await app.acquireTokenSilent({
          account: accounts[0],
        });
      } catch (e) {
        if (e.message !== 'Entry not found in cache.') {
          log.warn(`Unable to acquire token from cache: ${e}`);
        } else {
          log.debug(`Unable to acquire token from cache: ${e}`);
        }
      }
    }
    if (silentOnly) {
      return null;
    }

    try {
      if (this.refreshToken) {
        log.debug('acquire token with refresh token.');
        return await app.acquireTokenByRefreshToken({
          refreshToken: this.refreshToken,
        });
      } else if (this.clientSecret) {
        log.debug('acquire token with client credentials.');
        return await app.acquireTokenByClientCredential({
          scopes: this.scopes,
        });
      } else {
        const err = new StatusCodeError('No valid authentication credentials supplied.');
        err.statusCode = 401;
        throw err;
      }
    } catch (e) {
      log.error(`Error while acquiring access token ${e}`);
      throw e;
    }
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
