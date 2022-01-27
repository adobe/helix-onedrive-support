/*
 * Copyright 2021 Adobe. All rights reserved.
 * This file is licensed to you under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License. You may obtain a copy
 * of the License at http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software distributed under
 * the License is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR REPRESENTATIONS
 * OF ANY KIND, either express or implied. See the License for the specific language
 * governing permissions and limitations under the License.
 */

const fetchAPI = require('@adobe/helix-fetch');
const StatusCodeError = require('./StatusCodeError.js');

const { fetch } = process.env.HELIX_FETCH_FORCE_HTTP1
  ? fetchAPI.context({
    alpnProtocols: [fetchAPI.ALPN_HTTP1_1],
    userAgent: 'helix-fetch', // static user agent for test recordings
  })
  /* istanbul ignore next */
  : fetchAPI;

/**
 * Helper class accessing folders and files using the SharePoint V1 API.
 */
class SharePointSite {
  constructor(opts) {
    this._owner = opts.owner;
    this._site = opts.site;
    this._clientId = opts.clientId;
    this._tenantId = opts.tenantId;
    this._refreshToken = opts.refreshToken;
    this._root = opts.root || '';
    this._log = opts.log || console;
  }

  async getAccessToken() {
    const { log } = this;
    if (!this._accessToken || Date.now() >= this._expires) {
      const url = `https://login.microsoftonline.com/${this._tenantId}/oauth2/v2.0/token`;
      let resp;

      try {
        resp = await fetch(url, {
          method: 'POST',
          body: new URLSearchParams({
            client_id: this._clientId,
            refresh_token: this._refreshToken,
            grant_type: 'refresh_token',
            scope: `https://${this._owner}.sharepoint.com/Sites.ReadWrite.All`,
          }),
        });
      } catch (e) {
        log.error(`Error while getting a SharePoint API token: ${e}`);
        throw e;
      }
      if (!resp.ok) {
        const text = await resp.text();
        log.error(`Error while getting a SharePoint API token: ${text}}`);
        throw new StatusCodeError(text, resp.status);
      }
      const json = await resp.json();
      this._accessToken = json.access_token;
      this._expires = Date.now() + json.expires_in * 1000;
    }
    return this._accessToken;
  }

  static splitDirAndBase(file) {
    const idx = file.lastIndexOf('/');
    if (idx < 0) {
      return ['', file];
    }
    return [file.substring(0, idx), file.substring(idx + 1)];
  }

  async getFile(file) {
    const [dir, base] = SharePointSite.splitDirAndBase(file);
    const folder = dir ? `${this._root}/${dir}` : this._root;

    try {
      const result = await this.doFetch(`/GetFolderByServerRelativeUrl('${folder}')/Files('${base}')?$expand=ModifiedBy`);
      return result;
    } catch (e) {
      if (e.statusCode === 401) {
        // an inexistant share returns 401, we prefer to just say it wasn't found
        throw new StatusCodeError(e.message, 404, e.details);
      }
      throw e;
    }
  }

  async getFolder(folder) {
    try {
      const relPath = folder ? `${this._root}/${folder}` : this._root;
      const result = await this.doFetch(`/GetFolderByServerRelativeUrl('${relPath}')`);
      return result;
    } catch (e) {
      if (e.statusCode === 401) {
        // an inexistant share returns 401, we prefer to just say it wasn't found
        throw new StatusCodeError(e.message, 404, e.details);
      }
      if (e.statusCode === 404) {
        return null;
      }
      throw e;
    }
  }

  async getFileContents(file) {
    const [dir, base] = SharePointSite.splitDirAndBase(file);
    const folder = dir ? `${this._root}/${dir}` : this._root;

    try {
      const result = await this.doFetch(`/GetFolderByServerRelativeUrl('${folder}')/Files('${base}')/$value`, true);
      return result;
    } catch (e) {
      if (e.statusCode === 401) {
        // an inexistant share returns 401, we prefer to just say it wasn't found
        throw new StatusCodeError(e.message, 404, e.details);
      }
      throw e;
    }
  }

  async getFilesAndFolders(folder) {
    try {
      const relPath = folder ? `${this._root}/${folder}` : this._root;
      const result = await this.doFetch(`/GetFolderByServerRelativeUrl('${relPath}')?$expand=Files/ModifiedBy,Folders`);
      return result;
    } catch (e) {
      if (e.statusCode === 401) {
        // an inexistant share returns 401, we prefer to just say it wasn't found
        throw new StatusCodeError(e.message, 404, e.details);
      }
      throw e;
    }
  }

  async doFetch(relUrl, rawResponseBody = false, options = {}) {
    const opts = { ...options };
    const accessToken = await this.getAccessToken();
    if (!opts.headers) {
      opts.headers = {};
    }
    opts.headers.authorization = `Bearer ${accessToken}`;
    if (!rawResponseBody) {
      opts.headers.accept = 'application/json;odata=verbose';
    }

    const url = `https://${this._owner}.sharepoint.com/sites/${this._site}/_api/web${relUrl}`;
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
      throw StatusCodeError.fromError(e);
    }
  }

  get log() {
    return this._log;
  }
}

module.exports = SharePointSite;
