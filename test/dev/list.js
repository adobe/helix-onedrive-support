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
import { GetObjectCommand, ListObjectsV2Command, S3Client } from '@aws-sdk/client-s3';
import { Console } from 'node:console';

const out = process.stdout.write.bind(process.stdout);

// eslint-disable-next-line no-global-assign
console = new Console(process.stderr, process.stderr);

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

async function loadInfo(s3, id) {
  try {
    const ret = await s3.send(new GetObjectCommand({
      Bucket: 'helix-content-bus',
      Key: `${id}/.hlx.json`,
    }));
    return JSON.parse(await ret.Body.transformToString('utf-8'));
  } catch (e) {
    console.error(`unable to load .hlx.json for ${id}`, e);
    return null;
  }
}

function createCacheContext(type, project) {
  return {
    tokenCache: {
      deserialize(json) {
        try {
          const data = JSON.parse(json);
          // console.log(data);
          project.type = type;
          console.log('Mountpoint:', project.mountpoint);
          console.log('Repository:', project.repository);
          if (data.Account) {
            project.account = Object.values(data.Account)[0]?.username;
            console.log(`   Account: ${project.account}`);
          }
          if (data.id_token) {
            const payload = JSON.parse(Buffer.from(data.id_token.split('.')[1], 'base64'));
            project.user = payload.email;
            console.log(`      User: ${project.user}`);
          }
        } catch (e) {
          console.error('error deserializing cache', e);
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
  const projects = {};
  let counter = 0;
  for (const id of ids) {
    // eslint-disable-next-line no-plusplus
    if (++counter % 100 === 0) {
      console.log('processing projects', counter, '/', ids.length);
    }
    const getProject = async () => {
      let p = projects[id];
      if (!p) {
        p = {};
        const info = await loadInfo(s3, id);
        if (info) {
          p.repository = info['original-repository'];
          p.mountpoint = info.mountpoint;
        }
        projects[id] = p;
      }
      return p;
    };

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
      const project = await getProject();
      const onedriveContext = createCacheContext('OneDrive', project);
      const p = await onedriveCache.getCache('content');
      await p.beforeCacheAccess(onedriveContext);
    }
    if (await googleCache.hasCache('content')) {
      const project = await getProject();
      const googleContext = createCacheContext('Google', project);
      const p = await googleCache.getCache('content');
      await p.beforeCacheAccess(googleContext);
    }
  }

  out(JSON.stringify(projects, null, 2));
}

run().catch(console.error);
