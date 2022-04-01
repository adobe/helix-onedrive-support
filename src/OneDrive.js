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
const EventEmitter = require('events');
const { promisify } = require('util');
const jose = require('jose');
const { AuthenticationContext, MemoryCache } = require('adal-node');
const { fetch, reset } = require('@adobe/helix-fetch').keepAliveNoCache({ userAgent: 'helix-fetch' });

const Workbook = require('./Workbook.js');
const StatusCodeError = require('./StatusCodeError.js');
const { driveItemFromURL, driveItemToURL } = require('./utils.js');
const { splitByExtension, sanitize, editDistance } = require('./fuzzy-helper.js');
const SharePointSite = require('./SharePointSite.js');

const AZ_AUTHORITY_HOST_URL = 'https://login.windows.net';
const AZ_DEFAULT_RESOURCE = 'https://graph.microsoft.com'; // '00000002-0000-0000-c000-000000000000'; ??
const AZ_COMMON_TENANT = 'common';

/**
 * the maximum subscription time in milliseconds
 * @see https://docs.microsoft.com/en-us/graph/api/resources/subscription?view=graph-rest-1.0#maximum-length-of-subscription-per-resource-type
 *
 * @static
 * @memberOf OneDrive
 */
const MAX_SUBSCRIPTION_EXPIRATION_TIME = 4230 * 60 * 1000;

/**
 * map that caches share item data. key is a sharing url, the value a drive item.
 * @type {Map<string, string>}
 * @private
 */
const globalShareLinkCache = new Map();

/**
 * map that caches the tenant ids
 * @type {Map<string, string>}
 */
const globalTenantCache = new Map();

/**
 * Helper class that facilitates accessing one drive.
 */
class OneDrive extends EventEmitter {
  /**
   * @param {OneDriveOptions} opts Options
   */
  constructor(opts) {
    super(opts);
    this.clientId = opts.clientId;
    this.clientSecret = opts.clientSecret || '';
    this.refreshToken = opts.refreshToken || '';
    this.username = opts.username || '';
    this.password = opts.password || '';
    this._log = opts.log || console;
    this.tenant = opts.tenant;
    this.resource = opts.resource || AZ_DEFAULT_RESOURCE;
    this.localAuthCache = opts.localAuthCache;

    if (!opts.noShareLinkCache && !process.env.HELIX_ONEDRIVE_NO_SHARE_LINK_CACHE) {
      /** @type {Map<string, string>} */
      this.shareLinkCache = opts.shareLinkCache || globalShareLinkCache;
    }
    if (!opts.noTenantCache && !process.env.HELIX_ONEDRIVE_NO_TENANT_CACHE) {
      /** @type {Map<string, string>} */
      this.tenantCache = opts.tenantCache || globalTenantCache;
    }

    if (!this.clientId) {
      throw new Error('Missing clientId.');
    }
  }

  /**
   * Return the auth context
   * @returns {AuthenticationContext}
   */
  async getAuthContext() {
    if (!this.authContext) {
      this.authContext = new AuthenticationContext(
        this.getAuthorityUrl(),
        undefined,
        this.localAuthCache ? new MemoryCache() : undefined,
      );
      [
        'acquireUserCode',
        'acquireToken',
        'acquireTokenWithDeviceCode',
        'acquireTokenWithRefreshToken',
        'acquireTokenWithUsernamePassword',
        'acquireTokenWithClientCredentials',
      ].forEach((m) => {
        this.authContext[m] = promisify(this.authContext[m].bind(this.authContext));
      });
      const { cache } = this.authContext;
      if (this.localAuthCache) {
        const originalAdd = cache.add;
        cache.add = (entries, cb) => {
          originalAdd.call(cache, entries, (...args) => {
            // eslint-disable-next-line no-underscore-dangle
            this.emit('tokens', cache._entries);
            cb(...args);
          });
        };
        const originalRemove = cache.remove;
        cache.remove = (entries, cb) => {
          originalRemove.call(cache, entries, (...args) => {
            // eslint-disable-next-line no-underscore-dangle
            this.emit('tokens', cache._entries);
            cb(...args);
          });
        };
      }
      cache.add.promise = promisify(cache.add.bind(cache));
      cache.remove.promise = promisify(cache.remove.bind(cache));
      cache.find.promise = promisify(cache.find.bind(cache));
    }
    return this.authContext;
  }

  async resolveTenant(tenantHost) {
    const { log } = this;
    const configUrl = `https://login.windows.net/${tenantHost}.onmicrosoft.com/.well-known/openid-configuration`;
    const res = await fetch(configUrl);
    if (!res.ok) {
      log.info(`error fetching openid-configuration for ${tenantHost}: ${res.status}. Fallback to 'common'`);
      return AZ_COMMON_TENANT;
    }

    const { issuer } = await res.json();
    if (!issuer) {
      log.info(`unable to extract tenant from openid-configuration for ${tenantHost}: no 'issuer'. Fallback to 'common'`);
      return AZ_COMMON_TENANT;
    }

    // eslint-disable-next-line prefer-destructuring
    const tenant = new URL(issuer).pathname.split('/')[1];
    log.info(`fetched tenant information from for ${tenantHost}: ${tenant}`);
    return tenant;
  }

  async initTenantFromShareLink(sharingUrl) {
    if (this.tenant) {
      return;
    }
    const { log } = this;
    const url = sharingUrl instanceof URL
      ? sharingUrl
      : new URL(sharingUrl);
    let [tenantHost] = url.hostname.split('.');
    // special case: `xxxx-my.sharepoint.com`
    if (url.hostname.endsWith('-my.sharepoint.com')) {
      tenantHost = tenantHost.substring(0, tenantHost.length - 3);
    }

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
  get authenticated() {
    // eslint-disable-next-line no-underscore-dangle
    return this.authContext?.cache._entries.length > 0;
  }

  /**
   * Adds entries to the token cache
   * @param {TokenResponse[]} entries
   * @return this;
   */
  async loadTokenCache(entries) {
    return (await this.getAuthContext()).cache.add.promise(entries);
  }

  /**
   * Performs a login using an interactive flow which prompts the user to open a browser window and
   * enter the authorization code.
   * @params {function} [onCode] - optional function that gets invoked after code was retrieved.
   * @returns {Promise<TokenResponse>}
   */
  async login(onCode) {
    const { log } = this;
    const context = await this.getAuthContext();

    let code;
    try {
      code = await context.acquireUserCode(this.resource, this.clientId, 'en');
    } catch (e) {
      log.error('Error while requesting user code', e);
      throw e;
    }

    log.info(code.message);
    if (typeof onCode === 'function') {
      await onCode(code);
    }

    try {
      return await context.acquireTokenWithDeviceCode(this.resource, this.clientId, code);
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
   */
  setAccessToken(bearerToken) {
    const { log } = this;
    this.accessToken = {
      accessToken: bearerToken,
    };
    if (!this.tenant) {
      try {
        const { tid } = jose.decodeJwt(bearerToken);
        if (tid) {
          log.info(`using tenant from access token: ${tid}`);
          this.tenant = tid;
        }
      } catch (e) {
        log.warn(`unable to decode access token: ${e.message}`);
      }
    }
    this.accessToken.tenantId = this.tenant;
  }

  /**
   */
  async fetchAccessToken() {
    const { log } = this;
    const context = await this.getAuthContext();
    try {
      return await context.acquireToken(this.resource, this.username, this.clientId);
    } catch (e) {
      if (e.message !== 'Entry not found in cache.') {
        log.warn(`Unable to acquire token from cache: ${e}`);
      } else {
        log.debug(`Unable to acquire token from cache: ${e}`);
      }
    }

    try {
      if (this.refreshToken) {
        log.debug('acquire token with refresh token.');
        const resp = await context.acquireTokenWithRefreshToken(
          this.refreshToken,
          this.clientId,
          this.clientSecret,
          this.resource,
        );
        return await this.augmentAndCacheResponse(resp);
      } else if (this.username && this.password) {
        log.debug('acquire token with ROPC.');
        return await context.acquireTokenWithUsernamePassword(
          this.resource,
          this.username,
          this.password,
          this.clientId,
        );
      } else if (this.clientSecret) {
        log.debug('acquire token with client credentials.');
        return await context.acquireTokenWithClientCredentials(
          this.resource,
          this.clientId,
          this.clientSecret,
        );
      } else {
        const err = new StatusCodeError('No valid authentication credentials supplied.');
        err.statusCode = 401;
        throw err;
      }
    } catch (e) {
      log.error(`Error while refreshing access token ${e}`);
      throw e;
    }
  }

  async getAccessToken() {
    if (!this.accessToken) {
      this.accessToken = await this.fetchAccessToken();
    }
    return this.accessToken;
  }

  /**
   */
  createLoginUrl(redirectUri, state) {
    return `${this.getAuthorityUrl()}/oauth2/authorize?response_type=code&scope=/.default&client_id=${this.clientId}&redirect_uri=${redirectUri}&state=${state}&resource=${this.resource}`;
  }

  async augmentAndCacheResponse(response) {
    // somehow adal doesn't add the clientId and authority to response
    // eslint-disable-next-line no-underscore-dangle
    if (!response._clientId) {
      // eslint-disable-next-line no-underscore-dangle
      response._clientId = this.clientId;
      // eslint-disable-next-line no-underscore-dangle
      response._authority = this.authorityUrl;
    }
    const found = await this.authContext.cache.find.promise({
      refreshToken: response.refreshToken,
    });
    if (found.length) {
      await this.authContext.cache.remove.promise(found);
    }
    await this.authContext.cache.add.promise([response]);
    return response;
  }

  /**
   */
  async acquireToken(redirectUri, code) {
    const { log } = this;
    const context = await this.getAuthContext();
    try {
      const resp = await context.acquireTokenWithAuthorizationCode(
        code,
        redirectUri,
        this.resource,
        this.clientId,
        this.clientSecret,
      );
      return await this.augmentAndCacheResponse(resp);
    } catch (e) {
      log.error('Error while getting token with authorization code.', e);
      throw e;
    }
  }

  /**
   */
  async doFetch(relUrl, rawResponseBody = false, options = {}) {
    const opts = { ...options };
    const { accessToken } = await this.getAccessToken();
    if (!opts.headers) {
      opts.headers = {};
    }
    opts.headers.authorization = `Bearer ${accessToken}`;
    const url = `https://graph.microsoft.com/v1.0${relUrl}`;
    try {
      const resp = await fetch(url, opts);
      if (!resp.ok) {
        const text = await resp.text();
        let err;
        try {
          // try to parse json
          err = StatusCodeError.fromErrorResponse(JSON.parse(text), resp.status);
        } catch {
          err = new StatusCodeError(text, resp.status);
        }
        if (err.statusCode === 404) {
          this.log.warn(`${relUrl} : ${err.statusCode} - ${err.message}`);
        } else {
          this.log.error(`${relUrl} : ${err.statusCode} - ${err.message}`, err.details);
        }
        throw err;
      }
      // check content type before trying to parse a response body as JSON
      const contentType = resp.headers.get('content-type');
      const json = contentType && contentType.startsWith('application/json');

      // await result in order to be able to catch any error
      return await (rawResponseBody || !json ? resp.buffer() : resp.json());
    } catch (e) {
      if (e instanceof StatusCodeError) {
        throw e;
      }
      const err = StatusCodeError.fromError(e);
      this.log.error(`${relUrl} : ${err.statusCode} - ${err.message} - ${err.details}`);
      throw err;
    }
  }

  async me() {
    return this.doFetch('/me');
  }

  /**
   * Encodes the sharing url into a token that can be used to access drive items.
   * @param {string} sharingUrl A sharing URL from OneDrive
   * @see https://docs.microsoft.com/en-us/onedrive/developer/rest-api/api/shares_get?view=odsp-graph-online#encoding-sharing-urls
   * @returns {string} an id for a shared item.
   */
  static encodeSharingUrl(sharingUrl) {
    const base64 = Buffer
      .from(String(sharingUrl), 'utf-8')
      .toString('base64')
      .replace(/=/, '')
      .replace(/\//, '_')
      .replace(/\+/, '-');
    return `u!${base64}`;
  }

  /**
   */
  async resolveShareLink(sharingUrl) {
    await this.initTenantFromShareLink(sharingUrl);
    const link = OneDrive.encodeSharingUrl(sharingUrl);
    this.log.debug(`resolving sharelink ${sharingUrl} (${link})`);
    try {
      return await this.doFetch(`/shares/${link}/driveItem`);
    } catch (e) {
      if (e.statusCode === 401 || e.statusCode === 403) {
        // an inexistent share returns either 401 or 403, we prefer to just say it wasn't found
        throw new StatusCodeError(e.message, 404, e.details);
      }
      throw e;
    }
  }

  /**
   */
  async getDriveRootItem(driveId) {
    return this.doFetch(`/drives/${driveId}/root`);
  }

  /**
   */
  async getDriveItemFromShareLink(sharingUrl) {
    let driveItem = OneDrive.driveItemFromURL(sharingUrl);
    if (driveItem) {
      return driveItem;
    }
    await this.initTenantFromShareLink(sharingUrl);
    if (this.shareLinkCache) {
      driveItem = this.shareLinkCache.get(sharingUrl);
    }
    if (!driveItem) {
      driveItem = await this.resolveShareLink(sharingUrl);
      if (this.shareLinkCache) {
        this.shareLinkCache.set(sharingUrl, driveItem);
      }
    }
    return driveItem;
  }

  /**
   */
  async listChildren(folderItem, relPath = '', query = {}) {
    // eslint-disable-next-line no-param-reassign
    relPath = relPath.replace(/\/+$/, '');
    const rootPath = `/drives/${folderItem.parentReference.driveId}/items/${folderItem.id}`;
    let uri = !relPath ? `${rootPath}/children` : `${rootPath}:${relPath}:/children`;
    const qry = new URLSearchParams(query).toString();
    if (qry) {
      uri = `${uri}?${qry}`;
    }
    return this.doFetch(uri);
  }

  /**
   * Tries to get the drive items for the given folder and relative path, by loading the files of
   * the respective directory and returning the item with the best matching filename. Please note,
   * that only the files are matched 'fuzzily' but not the folders. The rules for transforming the
   * filenames to the name segment of the `relPath` are:
   * - convert to lower case
   * - replace all non-alphanumeric characters with a dash
   * - remove all consecutive dashes
   * - extensions are ignored, if the given path doesn't have one or if ignoreExtension is true
   *
   * The result is an array of drive items that match the given path. They are ordered by the edit
   * distance to the original name and then alphanumerically.
   *
   * @param {DriveItem} folderItem
   * @param {string} [relPath = '']
   * @param {boolean} [ignoreExtension = false]
   * @returns {Promise<DriveItem[]>}
   */
  async fuzzyGetDriveItem(folderItem, relPath = '', ignoreExtension = false) {
    if (relPath && !relPath.startsWith('/')) {
      throw new Error('relPath must be empty or start with /');
    }

    // first try to get item directly
    try {
      const ret = await this.getDriveItem(folderItem, relPath, false);
      if (relPath) {
        // eslint-disable-next-line prefer-destructuring
        ret.extension = splitByExtension(relPath)[1];
      }
      this.log.info(`fetched drive item directly: /drives/${folderItem.parentReference.driveId}/items/${folderItem.id}:${relPath}`);
      return [ret];
    } catch (e) {
      this.log.info(`fetched drive item directly failed: /drives/${folderItem.parentReference.driveId}/items/${folderItem.id}:${relPath} (${e.statusCode})`);
      // if no 404 or no relPath, propagate error
      if (e.statusCode !== 404 || !relPath) {
        throw e;
      }
    }

    const idx = relPath.lastIndexOf('/');
    const folderRelPath = relPath.substring(0, idx);
    const name = relPath.substring(idx + 1);
    const [baseName, ext] = splitByExtension(name);
    const sanitizedName = sanitize(baseName);

    const query = {
      $top: 999,
      $select: 'name,parentReference,file,id,size,webUrl,lastModifiedDateTime',
    };
    let fileList = [];
    do {
      // eslint-disable-next-line no-await-in-loop
      const result = await this.listChildren(folderItem, folderRelPath, query);
      fileList = fileList.concat(result.value);
      if (result['@odata.nextLink']) {
        const nextLink = new URL(result['@odata.nextLink']);
        query.$skiptoken = nextLink.searchParams.get('$skiptoken');
        this.log.debug(`fetching more children with skiptoken ${query.$skiptoken}`);
      } else {
        query.$skiptoken = null;
      }
    } while (query.$skiptoken);

    this.log.info(`loaded ${fileList.length} children from /drives/${folderItem.parentReference.driveId}/items/${folderItem.id}:${relPath}`);
    const items = fileList.filter((item) => {
      if (!item.file) {
        return false;
      }
      const [itemName, itemExt] = splitByExtension(item.name);
      // remember extension
      // eslint-disable-next-line no-param-reassign
      item.extension = itemExt;
      if (ext && ext !== itemExt && !ignoreExtension) {
        // only match extension if given via relPath
        return false;
      }
      const sanitizedItemName = sanitize(itemName);
      if (sanitizedItemName !== sanitizedName) {
        return false;
      }
      // compute edit distance
      // eslint-disable-next-line no-param-reassign
      item.fuzzyDistance = editDistance(baseName, itemName);
      return true;
    });

    // sort items by edit distance first and 2nd by item name
    items.sort((i0, i1) => {
      let c = i0.fuzzyDistance - i1.fuzzyDistance;
      if (c === 0) {
        c = i0.name.localeCompare(i1.name);
      }
      return c;
    });
    return items;
  }

  /**
   */
  async getDriveItem(folderItem, relPath = '', download = false) {
    // eslint-disable-next-line no-param-reassign
    relPath = relPath.replace(/\/+$/, '');
    const uri = relPath
      ? `/drives/${folderItem.parentReference.driveId}/items/${folderItem.id}:${relPath}`
      : `/drives/${folderItem.parentReference.driveId}/items/${folderItem.id}`;
    return download ? this.doFetch(`${uri}:/content`, true) : this.doFetch(uri);
  }

  /**
   */
  async downloadDriveItem(driveItem) {
    const uri = `/drives/${driveItem.parentReference.driveId}/items/${driveItem.id}/content`;
    return this.doFetch(uri, true);
  }

  /**
   * @see https://docs.microsoft.com/en-us/graph/api/driveitem-put-content?view=graph-rest-1.0&tabs=http
   */
  async uploadDriveItem(buffer, driveItem, relPath = '', conflictBehaviour = 'replace') {
    const validConflictBehaviours = [
      'replace',
      'rename',
      'fail',
    ];
    if (!validConflictBehaviours.includes(conflictBehaviour)) {
      throw new Error(`Bad confict behaviour: ${conflictBehaviour}, must be one of: ${validConflictBehaviours.join('/')}`);
    }
    // eslint-disable-next-line no-param-reassign
    relPath = relPath.replace(/\/+$/, '');
    if (relPath) {
      // eslint-disable-next-line no-param-reassign
      relPath = `:${relPath}:`;
    }

    // PUT /drives/{drive-id}/items/{parent-id}:/{filename}:/content
    const uri = `/drives/${driveItem.parentReference.driveId}/items/${driveItem.id}${relPath}/content?@microsoft.graph.conflictBehavior=${conflictBehaviour}`;
    const opts = {
      method: 'PUT',
      body: buffer,
      headers: {
        'Content-Type': 'application/octet-stream',
      },
    };
    return this.doFetch(uri, false, opts);
  }

  /**
   */
  getWorkbook(driveItem) {
    return new Workbook(this, `/drives/${driveItem.parentReference.driveId}/items/${driveItem.id}/workbook`, this.log);
  }

  /**
   */
  async listSubscriptions() {
    return this.doFetch('/subscriptions');
  }

  /**
   */
  async createSubscription({
    resource,
    notificationUrl,
    clientState,
    changeType = 'updated',
    expiresIn = MAX_SUBSCRIPTION_EXPIRATION_TIME,
  }) {
    const opts = {
      method: 'POST',
      body: {
        changeType,
        notificationUrl,
        resource,
        expirationDateTime: new Date(Date.now() + expiresIn).toISOString(),
        clientState,
      },
    };
    return this.doFetch('/subscriptions', false, opts);
  }

  /**
   */
  async refreshSubscription(id, expiresIn = MAX_SUBSCRIPTION_EXPIRATION_TIME) {
    this.log.debug(`refreshing expiration time of subscription ${id} by ${expiresIn} ms`);
    const opts = {
      uri: `/subscriptions/${id}`,
      method: 'PATCH',
      body: {
        expirationDateTime: new Date(Date.now() + expiresIn).toISOString(),
      },
    };
    return this.doFetch(`/subscriptions/${id}`, false, opts);
  }

  /**
   */
  async deleteSubscription(id) {
    this.log.debug(`deleting subscription ${id}`);
    const opts = {
      method: 'DELETE',
    };
    return this.doFetch(`/subscriptions/${id}`, true, opts);
  }

  /**
   * Fetches the changes from the respective resource using the provided delta token.
   * Use an empty token to fetch the initial state or `latest` to fetch the latest state.
   * @param {string} resource OneDrive resource path.
   * @param {string} [token] Delta token.
   * @returns {Promise<Array>} An object with an array of changes and a delta token.
   */
  async fetchChanges(resource, token) {
    let next = token ? `${resource}/delta?token=${token}` : `${resource}/delta`;
    let items = [];

    for (; ;) {
      const {
        value,
        '@odata.nextLink': nextLink,
        '@odata.deltaLink': deltaLink,
        // eslint-disable-next-line no-await-in-loop
      } = await this.doFetch(next);
      items = items.concat(value);
      if (nextLink) {
        // not the last page, we have a next link
        const nextToken = new URL(nextLink).searchParams.get('token');
        next = `${resource}/delta?token=${nextToken}`;
      } else if (deltaLink) {
        // last page, we have a next link
        return {
          changes: items,
          token: new URL(deltaLink).searchParams.get('token'),
        };
      } else {
        const error = new StatusCodeError('Received response with neither next nor delta link.', 500);
        this.log.error(error);
        throw error;
      }
    }
  }

  async getSite(siteURL) {
    this.log.debug(`getting site: (${siteURL})`);

    const match = siteURL.match(/^https:\/\/(\S+).sharepoint.com\/sites\/([^/]+)\/(\S+)$/);
    if (!match) {
      throw new Error(`Site URL does not match (*.sharepoint.com/sites/.*): ${match}`);
    }
    const [, owner, site, root] = match;

    try {
      const accessToken = await this.getAccessToken();
      return new SharePointSite({
        owner,
        site,
        root,
        clientId: this.clientId,
        tenantId: accessToken.tenantId,
        refreshToken: accessToken.refreshToken,
        log: this.log,
      });
    } catch (e) {
      if (e.statusCode === 401) {
        // an inexistant share returns 401, we prefer to just say it wasn't found
        throw new StatusCodeError(e.message, 404, e.details);
      }
      throw e;
    }
  }
}

module.exports = Object.assign(OneDrive, {
  MAX_SUBSCRIPTION_EXPIRATION_TIME,
  driveItemToURL,
  driveItemFromURL,
});
