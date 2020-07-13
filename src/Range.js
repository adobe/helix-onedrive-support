/*
 * Copyright 2020 Adobe. All rights reserved.
 * This file is licensed to you under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License. You may obtain a copy
 * of the License at http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software distributed under
 * the License is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR REPRESENTATIONS
 * OF ANY KIND, either express or implied. See the License for the specific language
 * governing permissions and limitations under the License.
 */
const StatusCodeError = require('./StatusCodeError.js');

/**
 * Returns the actual error, recursively descending through all error properties.
 *
 * @param {Error} e error caught
 */
function getActualError(e) {
  let error = e;
  while ('error' in error) {
    error = error.error;
  }
  return error;
}

class Range {
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
    try {
      if (!this._data) {
        const client = await this._oneDrive.getClient();
        this.log.debug(`get range data from ${this.uri}`);
        this._data = await client.get(this.uri);
      }
      return this._data;
    } catch (e) {
      this.log.error(getActualError(e));
      throw new StatusCodeError(e.message, e.statusCode || 500);
    }
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

  async getRowsAsObjects() {
    const { values } = await this.getData();
    const columnNames = values[0];
    const rows = values.map((row) => columnNames.reduce((obj, name, index) => {
      // eslint-disable-next-line no-param-reassign
      obj[name] = row[index];
      return obj;
    }, {}));
    // discard first row
    rows.shift();
    return rows;
  }
}

module.exports = Range;
