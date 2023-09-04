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
import { superTrim } from '../utils.js';

export class Range {
  constructor(oneDrive, uri, log) {
    this._oneDrive = oneDrive;
    this._uri = uri;
    this._log = log;
  }

  get uri() {
    return this._uri;
  }

  get log() {
    return this._log;
  }

  async getData() {
    if (!this._data) {
      this.log.debug(`get range data from ${this.uri}`);
      this._data = await this._oneDrive.doFetch(this.uri);
    }
    return this._data;
  }

  async getAddress() {
    return (await this.getData()).address;
  }

  async getAddressLocal() {
    return (await this.getData()).addressLocal;
  }

  async getColumnNames() {
    return (await this.getData()).values[0];
  }

  async getRowsAsObjects({ trim = false } = {}) {
    const values = await this.getValues();

    const columnNames = values[0];
    const rows = values.map((row) => columnNames.reduce((obj, name, index) => {
      if (trim) {
        // eslint-disable-next-line no-param-reassign
        obj[superTrim(name)] = superTrim(row[index]);
      } else {
        // eslint-disable-next-line no-param-reassign
        obj[name] = row[index];
      }
      return obj;
    }, {}));
    // discard first row
    rows.shift();
    return rows;
  }

  async getValues() {
    if (!this._values) {
      if (this._data) {
        this._values = this._data.values;
      } else {
        // optimization: ask for the values, only, not the complete range object
        this.log.debug(`get range values from ${this.uri}`);
        this._values = (await this._oneDrive.doFetch(`${this.uri}?$select=values`)).values;
      }
    }
    return this._values;
  }

  async update(newValues) {
    const result = await this._oneDrive.doFetch(this.uri, false, {
      method: 'PATCH',
      body: newValues,
    });
    this._values = result.values;
  }

  async delete(shift = 'Up') {
    await this._oneDrive.doFetch(`${this.uri}/delete`, false, {
      method: 'POST',
      body: { shift },
    });
    this._values = null;
    this._data = null;
  }

  async insert(shift = 'Down') {
    const result = await this._oneDrive.doFetch(`${this.uri}/insert`, false, {
      method: 'POST',
      body: { shift },
    });
    this._values = result.values;
  }
}
