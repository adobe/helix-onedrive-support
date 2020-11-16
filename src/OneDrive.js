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
const util = require('util');
const { AuthenticationContext } = require('adal-node');
const rp = require('request-promise-native');

const Workbook = require('./Workbook.js');
const StatusCodeError = require('./StatusCodeError.js');
const { driveItemFromURL, driveItemToURL } = require('./utils.js');
const { splitByExtension, sanitize, editDistance } = require('./fuzzy-helper.js');

const AZ_AUTHORITY_HOST_URL = 'https://login.windows.net';
const AZ_RESOURCE = 'https://graph.microsoft.com'; // '00000002-0000-0000-c000-000000000000'; ??
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
const shareItemCache = new Map();

/**
 * Helper class that facilitates accessing one drive.
 */
class OneDrive extends EventEmitter {
  /**
   * @param {OneDriveOptions} opts Options
   * @param {string} opts.clientId The client id of the app
   * @param {string} [opts.clientSecret] The client secret of the app
   * @param {string} [opts.refreshToken] The refresh token.
   * @param {string} [opts.accessToken] The access token.
   * @param {string} [opts.username] Username for username/password authentication.
   * @param {string} [opts.password] Password for username/password authentication.
   * @param {number} [opts.expiresOn] Expiration time.
   * @param {Logger} [opts.log] A logger.
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

    if (!this.clientId) {
      throw new Error('Missing clientId.');
    }
    this.authContext = new AuthenticationContext(this.authorityUrl);
    const { cache } = this.authContext;
    cache.find.promise = util.promisify(cache.find.bind(cache));
    cache.add.promise = util.promisify(cache.add.bind(cache));
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
    const code = await new Promise((resolve, reject) => {
      context.acquireUserCode(AZ_RESOURCE, this.clientId, 'en', (err, response) => {
        if (err) {
          log.error('Error while requesting user code', err);
          reject(err);
        } else {
          resolve(response);
        }
      });
    });

    log.info(code.message);
    if (typeof onCode === 'function') {
      await onCode(code);
    }

    return new Promise((resolve, reject) => {
      context.acquireTokenWithDeviceCode(AZ_RESOURCE, this.clientId, code,
        (err, response) => {
          if (err) {
            log.error('Error while requesting access token with device code', err);
            reject(err);
          } else {
            // eslint-disable-next-line no-underscore-dangle
            this.emit('tokens', context.cache._entries);
            this.refreshToken = response.refreshToken;
            resolve(response);
          }
        });
    });
  }

  /**
   */
  async getAccessToken() {
    const { log, authContext: context } = this;
    return new Promise((resolve, reject) => {
      const callback = (err, response) => {
        if (err) {
          log.error('Error while refreshing access token', err);
          reject(err);
        } else {
          log.debug('Token acquired.');
          // eslint-disable-next-line no-underscore-dangle
          this.emit('tokens', context.cache._entries);
          resolve(response);
        }
      };
      if (this.refreshToken) {
        log.debug('acquire token with refresh token.');
        context.acquireTokenWithRefreshToken(this.refreshToken, this.clientId,
          this.clientSecret, AZ_RESOURCE, callback);
      } else if (this.username && this.password) {
        log.debug('acquire token with ROPC.');
        context.acquireTokenWithUsernamePassword(AZ_RESOURCE, this.username, this.password,
          this.clientId, callback);
      } else {
        log.debug('acquire token with client credentials.');
        context.acquireTokenWithClientCredentials(AZ_RESOURCE, this.clientId, this.clientSecret,
          callback);
      }
    });
  }

  /**
   */
  createLoginUrl(redirectUri, state) {
    return `${this.authorityUrl}/oauth2/authorize?response_type=code&scope=/.default&client_id=${this.clientId}&redirect_uri=${redirectUri}&state=${state}&resource=${AZ_RESOURCE}`;
  }

  /**
   */
  async acquireToken(redirectUri, code) {
    const { log, authContext: context } = this;
    return new Promise((resolve, reject) => {
      context.acquireTokenWithAuthorizationCode(
        code,
        redirectUri,
        AZ_RESOURCE,
        this.clientId,
        this.clientSecret,
        async (err, response) => {
          if (err) {
            log.error('Error while getting token with authorization code.', err);
            reject(err);
          } else {
            // somehow adal doesn't add the clientId and authority to the this
            // eslint-disable-next-line no-underscore-dangle
            if (!response._clientId) {
              // eslint-disable-next-line no-underscore-dangle
              response._clientId = this.clientId;
              // eslint-disable-next-line no-underscore-dangle
              response._authority = this.authorityUrl;
            }
            const { cache } = context;
            const found = await cache.find.promise({
              refreshToken: response.refreshToken,
            });
            if (!found.length) {
              await cache.add.promise([response]);
            }
            // eslint-disable-next-line no-underscore-dangle
            this.emit('tokens', context.cache._entries);
            resolve(response);
          }
        },
      );
    });
  }

  /**
   */
  async getClient(raw = false) {
    const { accessToken } = await this.getAccessToken();
    const opts = {
      baseUrl: 'https://graph.microsoft.com/v1.0',
      json: true,
      auth: {
        bearer: accessToken,
      },
    };
    if (raw) {
      delete opts.json;
      opts.encoding = null;
    }
    return rp.defaults(opts);
  }

  async me() {
    try {
      return await (await this.getClient()).get('/me');
    } catch (e) {
      const error = StatusCodeError.fromError(e);
      this.log[(error.statusCode === 404) ? 'warn' : 'error'](error);
      throw error;
    }
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
      return await (await this.getClient()).get(`/shares/${link}/driveItem`);
    } catch (e) {
      const error = StatusCodeError.fromError(e);
      this.log[(error.statusCode === 404) ? 'warn' : 'error'](error);
      throw error;
    }
  }

  /**
   */
  async getDriveRootItem(driveId) {
    const uri = `/drives/${driveId}/root`;
    try {
      return await (await this.getClient()).get(uri);
    } catch (e) {
      const error = StatusCodeError.fromError(e);
      this.log[(error.statusCode === 404) ? 'warn' : 'error'](error);
      throw error;
    }
  }

  /**
   */
  async getDriveItemFromShareLink(sharingUrl) {
    let driveItem = OneDrive.driveItemFromURL(sharingUrl);
    if (driveItem) {
      return driveItem;
    }
    driveItem = shareItemCache.get(sharingUrl);
    if (!driveItem) {
      driveItem = await this.resolveShareLink(sharingUrl);
      shareItemCache.set(sharingUrl, driveItem);
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
    try {
      return await (await this.getClient()).get(uri);
    } catch (e) {
      const error = StatusCodeError.fromError(e);
      this.log[(error.statusCode === 404) ? 'warn' : 'error'](error);
      throw error;
    }
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
      $select: 'name,parentReference,file,id,size',
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
    try {
      if (download) {
        return (await this.getClient(true))
          .get(`${uri}:/content`);
      }
      return await (await this.getClient()).get(uri);
    } catch (e) {
      const error = StatusCodeError.fromError(e);
      this.log[(error.statusCode === 404) ? 'warn' : 'error'](error);
      throw error;
    }
  }

  /**
   */
  async downloadDriveItem(driveItem) {
    const uri = `/drives/${driveItem.parentReference.driveId}/items/${driveItem.id}/content`;
    try {
      return await (await this.getClient(true)).get(uri);
    } catch (e) {
      const error = StatusCodeError.fromError(e);
      this.log[(error.statusCode === 404) ? 'warn' : 'error'](error);
      throw error;
    }
  }

  /**
   * @see https://docs.microsoft.com/en-us/graph/api/driveitem-put-content?view=graph-rest-1.0&tabs=http
   */
  async uploadDriveItem(buffer, driveItem, relPath = '') {
    // eslint-disable-next-line no-param-reassign
    relPath = relPath.replace(/\/+$/, '');
    if (relPath) {
      // eslint-disable-next-line no-param-reassign
      relPath = `:${relPath}:`;
    }

    // PUT /drives/{drive-id}/items/{parent-id}:/{filename}:/content
    const uri = `/drives/${driveItem.parentReference.driveId}/items/${driveItem.id}${relPath}/content`;
    try {
      const client = await this.getClient(true);
      return await client({
        uri,
        method: 'PUT',
        body: buffer,
        headers: {
          'Content-Type': 'application/octet-stream',
        },
      });
    } catch (e) {
      const error = StatusCodeError.fromError(e);
      this.log[(error.statusCode === 404) ? 'warn' : 'error'](error);
      throw error;
    }
  }

  /**
   */
  getWorkbook(driveItem) {
    return new Workbook(this, `/drives/${driveItem.parentReference.driveId}/items/${driveItem.id}/workbook`, this.log);
  }

  /**
   */
  async listSubscriptions() {
    try {
      return await (await this.getClient()).get('/subscriptions');
    } catch (e) {
      const error = StatusCodeError.fromError(e);
      this.log[(error.statusCode === 404) ? 'warn' : 'error'](error);
      throw error;
    }
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
    try {
      return await (await this.getClient())({
        uri: '/subscriptions',
        method: 'POST',
        body: {
          changeType,
          notificationUrl,
          resource,
          expirationDateTime: new Date(Date.now() + expiresIn).toISOString(),
          clientState,
        },
        json: true,
        headers: {
          'content-type': 'application/json',
        },
      });
    } catch (e) {
      const error = StatusCodeError.fromError(e);
      this.log[(error.statusCode === 404) ? 'warn' : 'error'](error);
      throw error;
    }
  }

  /**
   */
  async refreshSubscription(id, expiresIn = MAX_SUBSCRIPTION_EXPIRATION_TIME) {
    this.log.debug(`refreshing expiration time of subscription ${id} by ${expiresIn} ms`);
    try {
      return await (await this.getClient())({
        uri: `/subscriptions/${id}`,
        method: 'PATCH',
        body: {
          expirationDateTime: new Date(Date.now() + expiresIn).toISOString(),
        },
        json: true,
        headers: {
          'content-type': 'application/json',
        },
      });
    } catch (e) {
      const error = StatusCodeError.fromError(e);
      this.log[(error.statusCode === 404) ? 'warn' : 'error'](error);
      throw error;
    }
  }

  /**
   */
  async deleteSubscription(id) {
    this.log.debug(`deleting subscription ${id}`);
    try {
      return await (await this.getClient())({
        uri: `/subscriptions/${id}`,
        method: 'DELETE',
      });
    } catch (e) {
      const error = StatusCodeError.fromError(e);
      this.log[(error.statusCode === 404) ? 'warn' : 'error'](error);
      throw error;
    }
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

    try {
      const client = await this.getClient();
      for (; ;) {
        const {
          value,
          '@odata.nextLink': nextLink,
          '@odata.deltaLink': deltaLink,
          // eslint-disable-next-line no-await-in-loop
        } = await client.get(next);
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
          throw new Error('Received response with neither next nor delta link.');
        }
      }
    } catch (e) {
      const error = StatusCodeError.fromError(e);
      this.log[(error.statusCode === 404) ? 'warn' : 'error'](error);
      throw error;
    }
  }
}

module.exports = Object.assign(OneDrive, {
  MAX_SUBSCRIPTION_EXPIRATION_TIME,
  driveItemToURL,
  driveItemFromURL,
});
