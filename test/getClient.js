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

const querystring = require('querystring');

function getClient(ops) {
  const f = ({
    method, uri, body,
  }) => {
    // eslint-disable-next-line prefer-const
    let [, , component, command] = uri.split('/');
    let query = [];
    if (component) {
      const index = component.indexOf('?');
      if (index !== -1) {
        query = querystring.parse(component.substring(index + 1));
        component = component.substring(0, index);
      }
    }
    return ops({
      method, component, query, command, body,
    });
  };
  f.get = (uri) => f({ method: 'GET', uri });
  return f;
}

module.exports = getClient;
