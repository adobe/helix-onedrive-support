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
import { GetObjectCommand, S3Client } from '@aws-sdk/client-s3';
import { Response } from '@adobe/fetch';
import { promisify } from 'util';
import zlib from 'zlib';

const gunzip = promisify(zlib.gunzip);

export async function getContentBusId(org, site) {
  let contentBusId;
  try {
    const s3 = new S3Client();
    const res = await s3.send(new GetObjectCommand({
      Bucket: 'helix-code-bus',
      Key: `${org}/${site}/main/helix-config.json`,
    }));
    contentBusId = res.Metadata['x-contentbus-id'].substring(2);
  } catch (e) {
    console.error(`unable to load helix-config.json:${e.message}`);
  }
  if (!contentBusId) {
    // load from helix5 config
    try {
      const s3 = new S3Client();
      const res = await s3.send(new GetObjectCommand({
        Bucket: 'helix-config-bus',
        Key: `orgs/${org}/sites/${site}.json`,
      }));
      let buf = await new Response(res.Body, {}).buffer();
      if (res.ContentEncoding === 'gzip') {
        buf = await gunzip(buf);
      }
      contentBusId = JSON.parse(buf).content.contentBusId;
    } catch (e) {
      console.error(`unable to load helix5 config:${e.message}`);
    }
  }
  return contentBusId;
}
