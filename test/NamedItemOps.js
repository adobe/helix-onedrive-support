/*
 * Copyright 2020 Adobe. All rights reserved.
 * This file is licensed to you under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License. You may obtain a copy
 * of the License at http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software distributed under
 * the License is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR REPRESENTATIONS
 * OF ANY KIND, either express or implied. See the License for the specific language
 * governing permissions and limitations under the License.
 */

'use strict';

const StatusCodeError = require('../src/StatusCodeError.js');

const namedItemOps = (namedItems) => ({ command, method, body }) => {
  if (!command) {
    return { value: namedItems };
  }
  if (command === 'add') {
    const namedItem = namedItems.find((i) => i.name === body.name);
    if (namedItem) {
      throw new StatusCodeError(`Named item already exists: ${namedItem.name}`, 400);
    }
    const len = namedItems.push({
      name: body.name,
      value: body.reference,
      comment: body.comment,
    });
    return namedItems[len - 1];
  }
  const index = namedItems.findIndex((i) => i.name === command);
  if (index === -1) {
    throw new StatusCodeError(`Named item not found: ${command}`, 404);
  }
  const item = namedItems[index];
  if (method === 'DELETE') {
    namedItems.splice(index, 1);
  }
  return item;
};

module.exports = namedItemOps;
