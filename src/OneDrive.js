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
const { AuthenticationContext } = require('adal-node');
const rp = require('request-promise-native');
const url = require('url');

const AZ_AUTHORITY_HOST_URL = 'https://login.windows.net';
const AZ_RESOURCE = 'https://graph.microsoft.com'; // '00000002-0000-0000-c000-000000000000'; ??
const AZ_TENANT = 'common';
const AZ_AUTHORITY_URL = `${AZ_AUTHORITY_HOST_URL}/${AZ_TENANT}`;

class StatusCodeError extends Error {
  constructor(msg, statusCode) {
    super(msg);
    this.statusCode = statusCode;
  }
}

/**
 * the maximum subscription time in milliseconds
 * @see https://docs.microsoft.com/en-us/graph/api/resources/subscription?view=graph-rest-1.0#maximum-length-of-subscription-per-resource-type
 */
const MAX_SUBSCRIPTION_EXPIRATION_TIME = 4230 * 60 * 1000;


/**
 * Remember the access token for future action invocations.
 */
let tokenCache = {};

/**
 * map that caches share item data. key is a sharing url, the value a drive item.
 * @type {Map<string, *>}
 */
const shareItemCache = new Map();

/**
 * Helper class that facilitates accessing one drive.
 */
class OneDrive extends EventEmitter {
  constructor(opts) {
    super(opts);
    this.clientId = opts.clientId;
    this.clientSecret = opts.clientSecret;
    this.refreshToken = opts.refreshToken;
    this._log = opts.log || console;
    tokenCache.accessToken = opts.accessToken || '';
    tokenCache.expiresOn = opts.expiresOn || undefined;
  }

  // eslint-disable-next-line class-methods-use-this
  get authenticated() {
    return !!tokenCache.accessToken;
  }

  async getAccessToken() {
    const { log } = this;
    if (tokenCache.accessToken) {
      const expires = Date.parse(tokenCache.expiresOn);
      if (expires >= (Date.now())) {
        log.info('access token still valid.');
        return tokenCache.accessToken;
      }
      log.info('access token is expired. Requesting new one.');
    }

    return new Promise((resolve, reject) => {
      const context = new AuthenticationContext(AZ_AUTHORITY_URL);
      context.acquireTokenWithRefreshToken(
        this.refreshToken,
        this.clientId,
        this.clientSecret,
        AZ_RESOURCE,
        (err, response) => {
          if (err) {
            log.error('Error while refreshing access token', err);
            reject(err);
          } else {
            tokenCache = response;
            this.emit('tokens', response);
            resolve(tokenCache.accessToken);
          }
        },
      );
    });
  }

  createLoginUrl(redirectUri, state) {
    return `${AZ_AUTHORITY_URL}/oauth2/authorize?response_type=code&client_id=${this.clientId}&redirect_uri=${redirectUri}&state=${state}&resource=${AZ_RESOURCE}`;
  }

  async acquireToken(redirectUri, code) {
    const context = new AuthenticationContext(AZ_AUTHORITY_URL);
    return new Promise((resolve, reject) => {
      context.acquireTokenWithAuthorizationCode(
        code,
        redirectUri,
        AZ_RESOURCE,
        this.clientId,
        this.clientSecret,
        (err, response) => {
          if (err) {
            reject(err);
          } else {
            tokenCache = response;
            this.emit('tokens', response);
            resolve();
          }
        },
      );
    });
  }

  async getClient(raw = false) {
    const token = await this.getAccessToken();
    const opts = {
      baseUrl: 'https://graph.microsoft.com/v1.0',
      json: true,
      auth: {
        bearer: token,
      },
    };
    if (raw) {
      delete opts.json;
      opts.encoding = null;
    }
    return rp.defaults(opts);
  }

  get log() {
    return this._log;
  }

  async me() {
    try {
      return (await this.getClient())
        .get('/me');
    } catch (e) {
      this.log.error(e);
      throw new StatusCodeError(e.msg, 500);
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
      .from(sharingUrl, 'utf-8')
      .toString('base64')
      .replace(/=/, '')
      .replace(/\//, '_')
      .replace(/\+/, '-');
    return `u!${base64}`;
  }

  async resolveShareLink(sharingUrl) {
    const link = OneDrive.encodeSharingUrl(sharingUrl);
    this.log.info(`resolving sharelink ${sharingUrl} (${link})`);
    try {
      return (await this.getClient())
        .get(`/shares/${link}/driveItem`);
    } catch (e) {
      this.log.error(e);
      throw new StatusCodeError(e.msg, 500);
    }
  }

  async getDriveItemFromShareLink(sharingUrl) {
    let driveItem = shareItemCache.get(sharingUrl);
    if (!driveItem) {
      driveItem = await this.resolveShareLink(sharingUrl);
      shareItemCache.set(sharingUrl, driveItem);
    }
    return driveItem;
  }

  async listChildren(folderItem, relPath) {
    // eslint-disable-next-line no-param-reassign
    relPath = relPath.replace(/\/+$/, '');
    const rootPath = `/drives/${folderItem.parentReference.driveId}/items/${folderItem.id}`;
    const uri = !relPath ? `${rootPath}/children` : `${rootPath}:${relPath}:/children`;
    try {
      return (await this.getClient())
        .get(uri);
    } catch (e) {
      this.log.error(e);
      throw new StatusCodeError(e.msg, 500);
    }
  }

  async getDriveItem(folderItem, relPath, download = false) {
    // eslint-disable-next-line no-param-reassign
    relPath = relPath.replace(/\/+$/, '');
    const uri = `/drives/${folderItem.parentReference.driveId}/items/${folderItem.id}:${relPath}`;
    try {
      if (download) {
        return (await this.getClient(true))
          .get(`${uri}:/content`);
      }
      return (await this.getClient())
        .get(uri);
    } catch (e) {
      this.log.error(e);
      throw new StatusCodeError(e.msg, 500);
    }
  }

  async downloadDriveItem(driveItem) {
    const uri = `/drives/${driveItem.parentReference.driveId}/items/${driveItem.id}/content`;
    try {
      return (await this.getClient(true))
        .get(uri);
    } catch (e) {
      this.log.error(e);
      throw new StatusCodeError(e.msg, 500);
    }
  }

  async listSubscriptions() {
    try {
      return (await this.getClient())
        .get('/subscriptions');
    } catch (e) {
      this.log.error(e);
      throw new StatusCodeError(e.msg, 500);
    }
  }

  async refreshSubscription(id, expiresIn = MAX_SUBSCRIPTION_EXPIRATION_TIME) {
    this.log.debug(`refreshing expiration time of subscription ${id} by ${expiresIn} ms`);
    try {
      return (await this.getClient())({
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
      this.log.error(e);
      throw new StatusCodeError(e.msg, 500);
    }
  }

  /**
   * Fetches the changes from the respective resource using the provided delta token.
   * Use an empty token to fetch the initial state or `latest` to fetch the latest state.
   * @param {string} resource OneDrive resource path.
   * @param {string} [token] Delta token.
   * @returns {Promise<[]>} A return object with the values and a `@odata.deltaLink`.
   */
  async fetchChanges(resource, token) {
    let next = token ? `${resource}/delta?token=${token}` : `${resource}/delta`;
    let items = [];

    try {
      const client = await this.getClient();
      for (;;) {
        const {
          value,
          '@odata.nextLink': nextLink,
          '@odata.deltaLink': deltaLink,
          // eslint-disable-next-line no-await-in-loop
        } = await client.get(next);
        items = items.concat(value);
        if (nextLink) {
          // not the last page, we have a next link
          const nextToken = url.parse(nextLink, true).query.token;
          next = `${resource}/delta?token=${nextToken}`;
        } else if (deltaLink) {
          // last page, we have a next link
          return {
            value: items,
            '@odata.deltaLink': deltaLink,
          };
        } else {
          throw new Error('Received response with neither next nor delta link.');
        }
      }
    } catch (e) {
      this.log.error(e);
      throw new StatusCodeError(e.msg, 500);
    }
  }
}

OneDrive.MAX_SUBSCRIPTION_EXPIRATION_TIME = MAX_SUBSCRIPTION_EXPIRATION_TIME;

module.exports = OneDrive;
