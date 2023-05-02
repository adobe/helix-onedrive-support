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

/* eslint-disable no-console */
import { decrypt } from '@adobe/helix-shared-tokencache';
import fs from 'fs/promises';

async function run() {
  if (process.argv.length < 3) {
    console.error('usage: node src/decrypt.js <file>');
    return -1;
  }
  const data = await fs.readFile(process.argv[2]);
  const decrypted = decrypt('***', data);
  process.stdout.write(JSON.stringify(JSON.parse(decrypted.toString('utf-8')), null, 2));
  process.stdout.write('\n');
  return 0;
}

run().then(process.exit);
