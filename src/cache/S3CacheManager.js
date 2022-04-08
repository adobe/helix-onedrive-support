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
const { basename } = require('path');
const { S3Client, ListObjectsV2Command } = require('@aws-sdk/client-s3');
const { S3CachePlugin } = require('./S3CachePlugin.js');

/**
 * aliases
 * @typedef {import("@azure/msal-node").ICachePlugin} ICachePlugin
 * @typedef {import("@azure/msal-node").TokenCacheContext} TokenCacheContext
 */

class S3CacheManager {
  constructor(opts) {
    this.log = opts.log || console;
    this.bucket = opts.bucket;
    this.prefix = opts.prefix;
    this.secret = opts.secret;
    this.type = opts.type;
    this.s3 = new S3Client();
  }

  getAuthObjectKey(key) {
    return `${this.prefix}/auth-${this.type}-${key}.json`;
  }

  async listCacheKeys() {
    const {
      log, s3, bucket, prefix,
    } = this;
    log.info('s3: list token cache', prefix);
    try {
      const res = await s3.send(new ListObjectsV2Command({
        Bucket: bucket,
        Prefix: `${prefix}/`,
      }));
      return (res.Contents || [])
        .map((entry) => basename(entry.Key))
        .filter((name) => (name.startsWith('auth-') && name.endsWith('.json')))
        .map((name) => name.replace(/auth-([a-z0-9]+)-([a-z0-9]+).json/i, '$2'));
    } catch (e) {
      log.info('s3: unable to list token caches', e);
      return [];
    }
  }

  /**
   * @param key
   * @returns {S3CachePlugin}
   */
  async getCache(key) {
    return new S3CachePlugin({
      log: this.log,
      key: this.getAuthObjectKey(key),
      secret: this.secret,
      bucket: this.bucket,
    });
  }
}

module.exports = {
  S3CacheManager,
};
