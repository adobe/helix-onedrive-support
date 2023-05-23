/*
 * Copyright 2023 Adobe. All rights reserved.
 * This file is licensed to you under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License. You may obtain a copy
 * of the License at http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software distributed under
 * the License is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR REPRESENTATIONS
 * OF ANY KIND, either express or implied. See the License for the specific language
 * governing permissions and limitations under the License.
 */

/* eslint-disable no-console,no-await-in-loop */

import { S3CacheManager } from '@adobe/helix-shared-tokencache';
import { ListObjectsV2Command, S3Client } from '@aws-sdk/client-s3';

async function list(s3, bucket, prefix, deep = false) {
  let ContinuationToken;
  const objects = [];
  do {
    // eslint-disable-next-line no-await-in-loop
    const result = await s3.send(new ListObjectsV2Command({
      Bucket: bucket,
      ContinuationToken,
      Prefix: prefix,
      Delimiter: deep ? '' : '/',
    }));
    ContinuationToken = result.IsTruncated ? result.NextContinuationToken : '';
    (result.Contents || []).forEach((content) => {
      objects.push({
        key: content.Key,
        lastModified: content.LastModified,
        contentLength: content.Size,
      });
    });
    (result.CommonPrefixes || []).forEach((content) => {
      objects.push({
        key: content.Prefix.substring(0, content.Prefix.length - 1),
        folder: true,
      });
    });
    // console.log(result);
  } while (ContinuationToken);
  return objects;
}

function createCacheContext(type) {
  return {
    tokenCache: {
      deserialize(json) {
        const data = JSON.parse(json);
        // console.log(data);
        console.log(`\n\n${type} Account: ${Object.values(data.Account)[0].username}`);
        if (data.AccessToken) {
          const accessToken = Object.values(data.AccessToken)[0];
          console.log(`Access token expires on: ${new Date(Number(accessToken.expires_on) * 1000).toISOString()}`);
        } else {
          console.log('no access token');
        }
      },

    },
  };
}

async function run() {
  const s3 = new S3Client();
  console.log('loading projects....');
  const ids = (await list(s3, 'helix-content-bus', ''))
    .map(({ key }) => key)
    .filter((key) => key.length === 59);
  console.log('loading projects.... done', ids.length);
  const onedriveContext = createCacheContext('OneDrive');
  const googleContext = createCacheContext('Google');
  for (const id of ids) {
    const onedriveCache = new S3CacheManager({
      log: console,
      prefix: `${id}/.helix-auth`,
      secret: id,
      bucket: 'helix-content-bus',
      type: 'onedrive',
    });
    const googleCache = new S3CacheManager({
      log: console,
      prefix: `${id}/.helix-auth`,
      secret: id,
      bucket: 'helix-content-bus',
      type: 'google',
    });
    if (await onedriveCache.hasCache('content')) {
      const p = await onedriveCache.getCache('content');
      await p.beforeCacheAccess(onedriveContext);
    }
    if (await googleCache.hasCache('content')) {
      const p = await googleCache.getCache('content');
      await p.beforeCacheAccess(googleContext);
    }
  }
}

run().catch(console.error);
