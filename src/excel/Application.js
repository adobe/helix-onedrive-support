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
export class Application {
  /**
   * Create a new instance of this class.
   * @param {import('../GraphAPI.js').GraphAPI} graphAPI graph API
   * @param {string} uri uri for this application
   * @param {any} log logger
   */
  constructor(graphAPI, uri, log) {
    this._graphAPI = graphAPI;
    this._uri = uri;
    this._log = log;
  }

  async calculate(calculationType) {
    return this._graphAPI.doFetch(`${this.uri}/calculate`, false, {
      method: 'POST',
      body: { calculationType },
      headers: { 'content-type': 'application/json' },
    });
  }

  get uri() {
    return this._uri;
  }
}
