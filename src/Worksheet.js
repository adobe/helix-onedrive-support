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

'use strict';

const NamedItemContainer = require('./NamedItemContainer.js');
const StatusCodeError = require('./StatusCodeError.js');
const Table = require('./Table.js');
const Range = require('./Range.js');

class Worksheet extends NamedItemContainer {
  constructor(oneDrive, prefix, name, log) {
    super(oneDrive);

    this._oneDrive = oneDrive;
    this._uri = `${prefix}/${name}`;
    this._name = name;
    this._log = log;
  }

  get uri() {
    return this._uri;
  }

  get log() {
    return this._log;
  }

  async getTableNames() {
    try {
      const client = await this._oneDrive.getClient();
      this.log.debug(`get table names from ${this._uri}/tables`);
      const result = await client.get(`${this._uri}/tables`);
      return result.value.map((v) => v.name);
    } catch (e) {
      this.log.error(e);
      throw new StatusCodeError(e.msg, 500);
    }
  }

  table(name) {
    return new Table(this._oneDrive, `${this._uri}/tables`, name, this._log);
  }

  usedRange() {
    return new Range(this._oneDrive, `${this._uri}/usedRange`, this._log);
  }
}

module.exports = Worksheet;
