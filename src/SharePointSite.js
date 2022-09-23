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
import { keepAliveNoCache } from '@adobe/fetch';
import { StatusCodeError } from './StatusCodeError.js';

const { fetch } = keepAliveNoCache({ userAgent: 'adobe-fetch' });

/**
 * Helper class accessing folders and files using the SharePoint V1 API.
 */
export class SharePointSite {
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
      const resp = await fetch(url, {
        method: 'POST',
        body: new URLSearchParams({
          client_id: this._clientId,
          refresh_token: this._refreshToken,
          grant_type: 'refresh_token',
          scope: `https://${this._owner}.sharepoint.com/Sites.ReadWrite.All`,
        }),
      });
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

  _splitDirAndBase(file) {
    const idx = file.lastIndexOf('/');
    const [dir, base] = (idx < 0)
      ? ['', file]
      : [file.substring(0, idx), file.substring(idx + 1)];
    return dir ? [`${this._root}/${dir}`, base] : [this._root, base];
  }

  _getRelativePath(folder) {
    return folder ? `${this._root}/${folder}` : this._root;
  }

  async getFile(file) {
    const [dir, base] = this._splitDirAndBase(file);
    return this.doFetch(`/GetFolderByServerRelativeUrl('${dir}')/Files('${base}')?$expand=ModifiedBy`);
  }

  async getFolder(folder) {
    const dir = this._getRelativePath(folder);
    return this.doFetch(`/GetFolderByServerRelativeUrl('${dir}')`);
  }

  async getFileContents(file) {
    const [dir, base] = this._splitDirAndBase(file);
    return this.doFetch(`/GetFolderByServerRelativeUrl('${dir}')/Files('${base}')/$value`, true);
  }

  async getFilesAndFolders(folder) {
    const dir = this._getRelativePath(folder);
    return this.doFetch(`/GetFolderByServerRelativeUrl('${dir}')?$expand=Files/ModifiedBy,Folders`);
  }

  async doFetch(relUrl, rawResponseBody = false) {
    const opts = { headers: {} };
    const accessToken = await this.getAccessToken();
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
      /* c8 ignore next 4 */
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
