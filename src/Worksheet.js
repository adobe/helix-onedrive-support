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

const StatusCodeError = require('./StatusCodeError.js');

class Worksheet {
  constructor(oneDrive, prefix, name, log) {
    this._oneDrive = oneDrive;
    this._prefix = prefix;
    this._name = name;
    this._log = log;
  }

  async getNamedItems() {
    try {
      const client = await this._oneDrive.getClient();
      const result = await client.get(`${this.uri}/names`);
      return result.value.map((v) => v.name);
    } catch (e) {
      this.log.error(e);
      throw new StatusCodeError(e.msg, 500);
    }
  }

  async getNamedItem(name) {
    try {
      const client = await this._oneDrive.getClient(false);
      return await client.get(`${this.uri}/names/${name}`);
    } catch (e) {
      if (e.statusCode === 404) {
        return null;
      }
      this.log.error(e);
      throw new StatusCodeError(e.msg, 500);
    }
  }

  async addNamedItem(name, reference, comment) {
    try {
      const client = await this._oneDrive.getClient();
      await client({
        uri: `${this.uri}/names/add`,
        method: 'POST',
        body: {
          name,
          reference,
          comment,
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

  get uri() {
    return `${this._prefix}/${this._name}`;
  }

  get log() {
    return this._log;
  }
}

module.exports = Worksheet;
