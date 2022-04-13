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
const { OneDrive } = require('./OneDrive.js');
const { OneDriveAuth } = require('./OneDriveAuth.js');
const { FSCachePlugin } = require('./cache/FSCachePlugin.js');
const { FSCacheManager } = require('./cache/FSCacheManager.js');
const { MemCachePlugin } = require('./cache/MemCachePlugin.js');
const { S3CachePlugin } = require('./cache/S3CachePlugin.js');
const { S3CacheManager } = require('./cache/S3CacheManager.js');
const { OneDriveMock } = require('./OneDriveMock.js');
const {
  splitByExtension,
  editDistance,
  sanitizeName,
  sanitizePath,
} = require('./fuzzy-helper.js');

module.exports = {
  OneDrive,
  OneDriveAuth,
  OneDriveMock,
  FSCachePlugin,
  FSCacheManager,
  MemCachePlugin,
  S3CachePlugin,
  S3CacheManager,
  utils: {
    splitByExtension,
    editDistance,
    sanitizeName,
    sanitizePath,
  },
};
