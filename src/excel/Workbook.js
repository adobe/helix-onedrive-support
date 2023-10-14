/*
 * Copyright 2022 Adobe. All rights reserved.
 * This file is licensed to you under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License. You may obtain a copy
 * of the License at http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software distributed under
 * the License is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR REPRESENTATIONS
 * OF ANY KIND, either express or implied. See the License for the specific language
 * governing permissions and limitations under the License.
 */
import { NamedItemContainer } from './NamedItemContainer.js';
import { StatusCodeError } from '../StatusCodeError.js';
import { Table } from './Table.js';
import { Worksheet } from './Worksheet.js';

export class Workbook extends NamedItemContainer {
  constructor(oneDrive, uri, log) {
    super(oneDrive);

    this._oneDrive = oneDrive;
    this._uri = uri;
    this._log = log;
  }

  async createSession() {
    if (this._sessionId) {
      throw new StatusCodeError('This workbook is already associated with a session', 400);
    }
    const uri = `${this.uri}/createSession`;
    const result = await this._oneDrive.doFetch(uri, false, {
      method: 'POST',
    });
    this._sessionId = result.id;
    return this._sessionId;
  }

  async closeSession() {
    if (this._sessionId) {
      const uri = `${this.uri}/closeSession`;
      await this._oneDrive.doFetch(uri, false, {
        method: 'POST',
        headers: {
          'Workbook-Session-Id': this._sessionId,
        },
      });
      this._sessionId = null;
      return;
    }
    throw new StatusCodeError('No session associated with workbook', 400);
  }

  async refreshSession() {
    if (this._sessionId) {
      const uri = `${this.uri}/refreshSession`;
      await this._oneDrive.doFetch(uri, false, {
        method: 'POST',
        headers: {
          'Workbook-Session-Id': this._sessionId,
        },
      });
      return;
    }
    throw new StatusCodeError('No session associated with workbook', 400);
  }

  getSessionId() {
    return this._sessionId;
  }

  setSessionId(sessionId) {
    if (this._sessionId) {
      throw new StatusCodeError('This workbook is already associated with a session', 400);
    }
    this._sessionId = sessionId;
  }

  async doFetch(relUrl, rawResponseBody, options) {
    const opts = { ...options };
    if (!opts.headers) {
      opts.headers = {};
    }
    if (this._sessionId) {
      opts.headers['Workbook-Session-Id'] = this._sessionId;
    }
    return this._oneDrive.doFetch(relUrl, rawResponseBody, opts);
  }

  async getData() {
    const result = await this.doFetch(this._uri);
    return result.value;
  }

  async getWorksheetNames() {
    this.log.debug(`get worksheet names from ${this._uri}/worksheets`);
    const result = await this.doFetch(`${this._uri}/worksheets`);
    return result.value.map((v) => v.name);
  }

  worksheet(name) {
    return new Worksheet(this, `${this._uri}/worksheets`, name, this._log);
  }

  async createWorksheet(sheetName) {
    const uri = `${this.uri}/worksheets`;
    await this.doFetch(uri, false, {
      method: 'POST',
      body: { name: sheetName },
      headers: { 'content-type': 'application/json' },
    });
    return this.worksheet(sheetName);
  }

  async deleteWorksheet(sheetName) {
    const uri = `${this.uri}/worksheets/${sheetName}`;
    await this.doFetch(uri, false, {
      method: 'DELETE',
      headers: { 'content-type': 'application/json' },
    });
  }

  async getTableNames() {
    this.log.debug(`get table names from ${this._uri}/tables`);
    const result = await this.doFetch(`${this._uri}/tables`);
    return result.value.map((v) => v.name);
  }

  table(name) {
    return new Table(this, `${this._uri}/tables`, name, this._log);
  }

  async addTable(address, hasHeaders, name) {
    if (name) {
      const names = await this.getTableNames();
      if (names.includes(name)) {
        throw new StatusCodeError(`Table name already exists: ${name}`, 409);
      }
    }
    const result = await this.doFetch(`${this.uri}/tables/add`, false, {
      method: 'POST',
      body: {
        address,
        hasHeaders,
      },
    });
    const table = this.table(result.name);
    if (name && name !== table.name) {
      await table.rename(name);
    }
    return table;
  }

  get uri() {
    return this._uri;
  }

  get log() {
    return this._log;
  }
}
