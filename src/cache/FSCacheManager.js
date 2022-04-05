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
const fs = require('fs/promises');
const path = require('path');
const { FSCachePlugin } = require('./FSCachePlugin.js');

/**
 * aliases
 * @typedef {import("@azure/msal-node").ICachePlugin} ICachePlugin
 * @typedef {import("@azure/msal-node").TokenCacheContext} TokenCacheContext
 */

class FSCacheManager {
  constructor(opts) {
    this.dirPath = opts.dirPath;
    this.log = opts.log || console;
  }

  getCacheFilePath(key) {
    return path.resolve(this.dirPath, `auth-${key}.json`);
  }

  async listCacheKeys() {
    try {
      const files = await fs.readdir(this.dirPath);
      return files
        .filter((name) => (name.startsWith('auth-') && name.endsWith('.json')))
        .map((name) => name.replace(/auth-([a-z0-9]+).json/i, '$1'));
    } catch (e) {
      if (e.code === 'ENOENT') {
        return [];
      }
      throw e;
    }
  }

  /**
   * @param key
   * @returns {FSCachePlugin}
   */
  getCache(key) {
    return new FSCachePlugin({
      log: this.log,
      filePath: this.getCacheFilePath(key),
    });
  }

  async createCache(key) {

  }

  async deleteCache(key) {
    try {
      await fs.rm(this.getCacheFilePath(key));
    } catch (e) {
      this.log.warn(`error deleting cache: ${e.message}`);
      // ignore
    }
  }
}

module.exports = {
  FSCacheManager,
};
