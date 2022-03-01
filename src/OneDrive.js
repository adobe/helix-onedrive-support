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
const { AuthenticationContext, MemoryCache } = require('adal-node');
const fetchAPI = require('@adobe/helix-fetch');

const Workbook = require('./Workbook.js');
const StatusCodeError = require('./StatusCodeError.js');
const { driveItemFromURL, driveItemToURL } = require('./utils.js');
const { splitByExtension, sanitize, editDistance } = require('./fuzzy-helper.js');

const { fetch, reset } = process.env.HELIX_FETCH_FORCE_HTTP1
  ? fetchAPI.context({
    alpnProtocols: [fetchAPI.ALPN_HTTP1_1],
    userAgent: 'helix-fetch', // static user agent for test recordings
  })
  /* istanbul ignore next */
  : fetchAPI;

const AZ_AUTHORITY_HOST_URL = 'https://login.windows.net';
const AZ_DEFAULT_RESOURCE = 'https://graph.microsoft.com'; // '00000002-0000-0000-c000-000000000000'; ??
const AZ_DEFAULT_TENANT = 'common';

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
 * @type {Map<string, *>}
 * @private
 */
const globalShareLinkCache = new Map();

/**
 * Helper class that facilitates accessing one drive.
 */
class OneDrive extends EventEmitter {
  /**
   * @param {OneDriveOptions} opts Options
   * @param {string}  opts.clientId The client id of the app
   * @param {string}  [opts.clientSecret] The client secret of the app
   * @param {string}  [opts.refreshToken] The refresh token.
   * @param {string}  [opts.accessToken] The access token.
   * @param {string}  [opts.username] Username for username/password authentication.
   * @param {string}  [opts.password] Password for username/password authentication.
   * @param {number}  [opts.expiresOn] Expiration time.
   * @param {Logger}  [opts.log] A logger.
   * @param {boolean} [opts.localAuthCache] Whether to use local auth cache
   * @param {string}  [opts.resource] Azure resource to authenticate against. defaults to MS Graph.
   */
  constructor(opts) {
    super(opts);
    this.clientId = opts.clientId;
    this.clientSecret = opts.clientSecret || '';
    this.refreshToken = opts.refreshToken || '';
    this.username = opts.username || '';
    this.password = opts.password || '';
    this._log = opts.log || console;
    this.tenant = opts.tenant || AZ_DEFAULT_TENANT;
    this.resource = opts.resource || AZ_DEFAULT_RESOURCE;

    if (!opts.noShareLinkCache && !process.env.HELIX_ONEDRIVE_NO_SHARE_LINK_CACHE) {
      this.shareLinkCache = opts.shareLinkCache || globalShareLinkCache;
    }

    if (!this.clientId) {
      throw new Error('Missing clientId.');
    }
    this.authContext = new AuthenticationContext(
      this.authorityUrl,
      undefined,
      opts.localAuthCache ? new MemoryCache() : undefined,
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
    if (opts.localAuthCache) {
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

  get authorityUrl() {
    return `${AZ_AUTHORITY_HOST_URL}/${this.tenant}`;
  }

  /**
   * @returns {boolean}
   */
  get authenticated() {
    // eslint-disable-next-line no-underscore-dangle
    return this.authContext.cache._entries.length > 0;
  }

  /**
   * Adds entries to the token cache
   * @param {TokenResponse[]} entries
   * @return this;
   */
  async loadTokenCache(entries) {
    return this.authContext.cache.add.promise(entries);
  }

  /**
   * Performs a login using an interactive flow which prompts the user to open a browser window and
   * enter the authorization code.
   * @params {function} [onCode] - optional function that gets invoked after code was retrieved.
   * @returns {Promise<TokenResponse>}
   */
  async login(onCode) {
    const { log, authContext: context } = this;

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
   */
  async getAccessToken() {
    const { log, authContext: context } = this;
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
          this.refreshToken, this.clientId, this.clientSecret, this.resource,
        );
        return await this.augmentAndCacheResponse(resp);
      } else if (this.username && this.password) {
        log.debug('acquire token with ROPC.');
        return await context.acquireTokenWithUsernamePassword(
          this.resource, this.username, this.password, this.clientId,
        );
      } else if (this.clientSecret) {
        log.debug('acquire token with client credentials.');
        return await context.acquireTokenWithClientCredentials(
          this.resource, this.clientId, this.clientSecret,
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

  /**
   */
  createLoginUrl(redirectUri, state) {
    return `${this.authorityUrl}/oauth2/authorize?response_type=code&scope=/.default&client_id=${this.clientId}&redirect_uri=${redirectUri}&state=${state}&resource=${this.resource}`;
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
    const { log, authContext: context } = this;
    try {
      const resp = await context.acquireTokenWithAuthorizationCode(
        code, redirectUri, this.resource, this.clientId, this.clientSecret,
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
    const link = OneDrive.encodeSharingUrl(sharingUrl);
    this.log.debug(`resolving sharelink ${sharingUrl} (${link})`);
    try {
      return await this.doFetch(`/shares/${link}/driveItem`);
    } catch (e) {
      if (e.statusCode === 401) {
        // an inexistant share returns 401, we prefer to just say it wasn't found
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
   * - extensions are ignored, if the given path doesn't have one
   *
   * The result is an array of drive items that match the given path. They are ordered by the edit
   * distance to the original name and then alphanumerically.
   *
   * @param {DriveItem} folderItem
   * @param {string} relPath
   * @returns {Promise<DriveItem[]>}
   */
  async fuzzyGetDriveItem(folderItem, relPath = '') {
    const idx = relPath.lastIndexOf('/');
    if (idx < 0) {
      const ret = await this.getDriveItem(folderItem, '', false);
      // todo: add extra extension
      return [ret.value];
    }
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

    this.log.debug(`loaded ${fileList.length} children from ${relPath}`);
    const items = fileList.filter((item) => {
      if (!item.file) {
        return false;
      }
      const [itemName, itemExt] = splitByExtension(item.name);
      // remember extension
      // eslint-disable-next-line no-param-reassign
      item.extension = itemExt;
      if (ext && ext !== itemExt) {
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
}

module.exports = Object.assign(OneDrive, {
  MAX_SUBSCRIPTION_EXPIRATION_TIME,
  driveItemToURL,
  driveItemFromURL,
});
