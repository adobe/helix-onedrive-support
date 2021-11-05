/*
 * Copyright 2021 Adobe. All rights reserved.
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

/**
 * Cache plugin for MSAL
 * @class FSCachePlugin
 * @implements ICachePlugin
 */
module.exports = class FSCachePlugin {
  constructor(filePath) {
    this.filePath = filePath;
    this.log = console;
  }

  withLogger(logger) {
    this.log = logger;
    return this;
  }

  async beforeCacheAccess(cacheContext) {
    const { log, filePath } = this;
    try {
      cacheContext.tokenCache.deserialize(await fs.readFile(filePath, 'utf-8'));
    } catch (e) {
      log.warn('unable to deserialize', e);
    }
  }

  async afterCacheAccess(cacheContext) {
    const { filePath } = this;
    if (cacheContext.cacheHasChanged) {
      // reparse and create a nice formatted JSON
      const tokens = JSON.parse(cacheContext.tokenCache.serialize());
      await fs.writeFile(filePath, JSON.stringify(tokens, null, 2), 'utf-8');
    }
  }
};
