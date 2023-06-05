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
import { StatusCodeError } from '../StatusCodeError.js';

export class NamedItemContainer {
  constructor(oneDrive) {
    this._oneDrive = oneDrive;
  }

  async getNamedItems() {
    const result = await this._oneDrive.doFetch(`${this.uri}/names`);
    return result.value.map((v) => ({
      name: v.name,
      value: v.value,
      comment: v.comment,
    }));
  }

  async getNamedItem(name) {
    try {
      // await result in order to be able to catch errors
      return await this._oneDrive.doFetch(`${this.uri}/names/${name}`);
    } catch (e) {
      if (e.statusCode === 404) {
        return null;
      }
      throw e;
    }
  }

  async addNamedItem(name, reference, comment) {
    try {
      // await result in order to be able to catch errors
      return await this._oneDrive.doFetch(`${this.uri}/names/add`, false, {
        method: 'POST',
        body: {
          name,
          reference,
          comment,
        },
      });
    } catch (e) {
      if ((e.details && e.details.code === 'ItemAlreadyExists') && e.statusCode !== 409) {
        throw new StatusCodeError(e.message, 409);
      }
      throw e;
    }
  }

  async deleteNamedItem(name) {
    try {
      // await result in order to be able to catch errors
      return await this._oneDrive.doFetch(`${this.uri}/names/${name}`, true, {
        method: 'DELETE',
      });
    } catch (e) {
      if ((e.details && e.details.code === 'ItemNotFound') && e.statusCode !== 404) {
        throw new StatusCodeError(e.message, 404);
      }
      throw e;
    }
  }
}
