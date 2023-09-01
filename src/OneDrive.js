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
import { Workbook } from './excel/Workbook.js';
import { StatusCodeError } from './StatusCodeError.js';
import { editDistance, sanitizeName, splitByExtension } from './utils.js';
import { SharePointSite } from './SharePointSite.js';
import { RateLimit } from './RateLimit.js';

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
 * Helper class that facilitates accessing one drive.
 *
 * @class
 * @field {ConfidentialClientApplication|PublicClientApplication} app
 */
export class OneDrive {
  /**
   * Returns an onedrive uri for the given drive item. the uri has the format:
   * `onedrive:/drives/<driveId>/items/<itemId>`
   *
   * @param {DriveItem} driveItem
   * @returns {URL} An url representing the drive item
   */
  static driveItemToURL(driveItem) {
    return new URL(`onedrive:/drives/${driveItem.parentReference.driveId}/items/${driveItem.id}`);
  }

  /**
   * Returns a partial drive item from the given url. The urls needs to have the format:
   * `onedrive:/drives/<driveId>/items/<itemId>`. if the url does not start with the correct
   * protocol, {@code null} is returned.
   *
   * @param {URL|string} url The url of the drive item.
   * @return {DriveItem} A (partial) drive item.
   */
  static driveItemFromURL(url) {
    if (!(url instanceof URL)) {
      // eslint-disable-next-line no-param-reassign
      url = new URL(String(url));
    }
    if (url.protocol !== 'onedrive:') {
      return null;
    }
    const [drives, driveId, items, itemId] = url.pathname.split('/').filter((s) => !!s);
    if (drives !== 'drives') {
      throw new Error(`URI not supported (missing 'drives' segment): ${url}`);
    }
    if (items !== 'items') {
      throw new Error(`URI not supported (missing 'items' segment): ${url}`);
    }
    return {
      id: itemId,
      parentReference: {
        driveId,
      },
    };
  }

  /**
   * @param {OneDriveOptions} opts Options
   */
  constructor(opts) {
    this.fetchContext = keepAliveNoCache({ userAgent: 'adobe-fetch' });

    if (!opts.auth) {
      throw new Error('Missing auth.');
    }
    this.auth = opts.auth;
    this._log = opts.auth.log;

    if (!opts.noShareLinkCache && !process.env.HELIX_ONEDRIVE_NO_SHARE_LINK_CACHE) {
      /** @type {Map<string, string>} */
      this.shareLinkCache = opts.shareLinkCache || globalShareLinkCache;
    }
  }

  /**
   */
  async dispose() {
    // TODO: clear other state?
    await this.auth.dispose();
    const { reset } = this.fetchContext;
    return reset();
  }

  /**
   */
  get log() {
    return this._log;
  }

  /**
   */
  async doFetch(relUrl, rawResponseBody = false, options = {}) {
    const opts = { ...options };
    const { accessToken } = await this.auth.authenticate();
    if (!opts.headers) {
      opts.headers = {};
    }
    opts.headers.authorization = `Bearer ${accessToken}`;

    const { log, auth: { logFields, tenant } } = this;
    const url = `https://graph.microsoft.com/v1.0${relUrl}`;
    const method = opts.method || 'GET';

    try {
      const { fetch } = this.fetchContext;
      const resp = await fetch(url, opts);
      log.info(`OneDrive API [tenant:${tenant}] ${logFields}: ${method} ${relUrl} ${resp.status}`);

      const rateLimit = RateLimit.fromHeaders(resp.headers);
      if (rateLimit) {
        log.warn({ sharepointRateLimit: { tenant, ...rateLimit.toJSON() } });
      }

      if (!resp.ok) {
        const text = await resp.text();
        let err;
        try {
          // try to parse json
          err = StatusCodeError.fromErrorResponse(JSON.parse(text), resp.status, rateLimit);
        } catch {
          err = new StatusCodeError(text, resp.status);
        }
        throw err;
      }
      // check content type before trying to parse a response body as JSON
      const contentType = resp.headers.get('content-type');
      const json = contentType && contentType.startsWith('application/json');

      // await result in order to be able to catch any error
      return await (rawResponseBody || !json ? resp.buffer() : resp.json());
    } catch (e) {
      let err = e;
      if (!(e instanceof StatusCodeError)) {
        err = StatusCodeError.fromError(e);
      }
      log.info(`OneDrive API [tenant:${tenant}] ${logFields}: ${method} ${relUrl} ${e.statusCode}`);
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
    await this.auth.initTenantFromUrl(sharingUrl);
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
    await this.auth.initTenantFromUrl(sharingUrl);
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
    const sanitizedName = sanitizeName(baseName);

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
      const sanitizedItemName = sanitizeName(itemName);
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
  async getParentDriveItem(driveItem) {
    const parentURI = `/drives/${driveItem.parentReference.driveId}/items/${driveItem.parentReference.id}`;
    return this.doFetch(parentURI, false);
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
      const authResult = await this.auth.authenticate();
      return new SharePointSite({
        owner,
        site,
        root,
        clientId: this.clientId,
        tenantId: authResult.tenantId,
        refreshToken: authResult.refreshToken,
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

OneDrive.MAX_SUBSCRIPTION_EXPIRATION_TIME = MAX_SUBSCRIPTION_EXPIRATION_TIME;
