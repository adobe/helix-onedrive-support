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

'use strict';

const NamedItemContainer = require('./NamedItemContainer.js');
const Table = require('./Table.js');
const Worksheet = require('./Worksheet.js');

class Workbook extends NamedItemContainer {
  constructor(oneDrive, uri, log) {
    super(oneDrive);

    this._oneDrive = oneDrive;
    this._uri = uri;
    this._log = log;
  }

  async getData() {
    const result = await this._oneDrive.doFetch(this._uri);
    return result.value;
  }

  async getWorksheetNames() {
    this.log.debug(`get worksheet names from ${this._uri}/worksheets`);
    const result = await this._oneDrive.doFetch(`${this._uri}/worksheets`);
    return result.value.map((v) => v.name);
  }

  worksheet(name) {
    return new Worksheet(this._oneDrive, `${this._uri}/worksheets`, name, this._log);
  }

  async getTableNames() {
    this.log.debug(`get table names from ${this._uri}/tables`);
    const result = await this._oneDrive.doFetch(`${this._uri}/tables`);
    return result.value.map((v) => v.name);
  }

  table(name) {
    return new Table(this._oneDrive, `${this._uri}/tables`, name, this._log);
  }

  get uri() {
    return this._uri;
  }

  get log() {
    return this._log;
  }
}

module.exports = Workbook;
