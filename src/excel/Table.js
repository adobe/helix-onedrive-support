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

export class Table {
  constructor(oneDrive, prefix, name, log) {
    this._oneDrive = oneDrive;
    this._prefix = prefix;
    this._name = name;
    this._log = log;
  }

  async rename(name) {
    // TODO: check name for allowed characters and length
    const result = await this._oneDrive.doFetch(this.uri, false, {
      method: 'PATCH',
      body: {
        name,
      },
    });
    this._name = name;
    return result;
  }

  async getHeaderNames() {
    const result = await this._oneDrive.doFetch(`${this.uri}/headerRowRange`);
    return result.values[0];
  }

  async getRows() {
    const result = await this._oneDrive.doFetch(`${this.uri}/rows`);
    return result.value.map((v) => v.values[0]);
  }

  async getRowsAsObjects({ trim = false } = {}) {
    const { log } = this;
    this.log.debug(`get columns from ${this.uri}/columns`);
    const result = await this._oneDrive.doFetch(`${this.uri}/columns`);
    const columnNames = result.value.map(({ name }) => name);
    log.debug(`got column names: ${columnNames}`);

    const rowValues = result.value[0].values
      .map((_, rownum) => columnNames.reduce((row, name, column) => {
        const [value] = result.value[column].values[rownum];
        if (trim) {
          // eslint-disable-next-line no-param-reassign
          row[superTrim(name)] = superTrim(value);
        } else {
          // eslint-disable-next-line no-param-reassign
          row[name] = value;
        }
        return row;
      }, {}));

    // discard the first row
    rowValues.shift();
    return rowValues;
  }

  async getRow(index) {
    const result = await this._oneDrive.doFetch(`${this.uri}/rows/itemAt(index=${index})`);
    return result.values[0];
  }

  async addRow(values) {
    const result = await this.addRows([values]);
    return result;
  }

  async addRows(values) {
    const result = await this._oneDrive.doFetch(`${this.uri}/rows/add`, false, {
      method: 'POST',
      body: {
        index: null,
        values,
      },
    });
    return result.index;
  }

  async replaceRow(index, values) {
    return this._oneDrive.doFetch(`${this.uri}/rows/itemAt(index=${index})`, false, {
      method: 'PATCH',
      body: {
        values: [values],
      },
    });
  }

  async deleteRow(index) {
    return this._oneDrive.doFetch(`${this.uri}/rows/itemAt(index=${index})`, true, {
      method: 'DELETE',
    });
  }

  async getRowCount() {
    const result = await this._oneDrive.doFetch(`${this.uri}/dataBodyRange?$select=rowCount`);
    return result.rowCount;
  }

  async getColumn(name) {
    const result = await this._oneDrive.doFetch(`${this.uri}/columns('${name}')`);
    return result.values;
  }

  async addColumn(name, index) {
    const body = {
      name,
    };
    if (index !== undefined) {
      body.index = index;
    }
    return this._oneDrive.doFetch(`${this.uri}/columns`, false, {
      method: 'POST',
      body,
    });
  }

  async deleteColumn(name) {
    return this._oneDrive.doFetch(`${this.uri}/columns/${name}`, true, {
      method: 'DELETE',
    });
  }

  get name() {
    return this._name;
  }

  get uri() {
    return `${this._prefix}/${this._name}`;
  }

  get log() {
    return this._log;
  }
}
