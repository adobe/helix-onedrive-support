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
const Workbook = require('../src/Workbook.js');
const StatusCodeError = require('../src/StatusCodeError.js');
const namedItemOps = require('./NamedItemOps.js');

class MockOneDrive {
  constructor() {
    this.workbooks = [];
  }

  registerWorkbook(driveId, itemId, data) {
    this.workbooks.push({
      uri: `/drives/${driveId}/items/${itemId}/workbook`,
      data,
    });
    return this;
  }

  getWorkbook(driveItem) {
    const uri = driveItem
      ? `/drives/${driveItem.parentReference.driveId}/items/${driveItem.id}/workbook`
      : this.workbooks[0].uri;
    return new Workbook(this, uri, console);
  }

  getClient() {
    const f = ({ method, uri, body }) => {
      const wb = this.workbooks.find((w) => (uri.startsWith(w.uri)));
      if (!wb) {
        throw new StatusCodeError('not found', 404);
      }
      const { data } = wb;

      // eslint-disable-next-line prefer-const
      let [, component, command] = uri.substring(wb.uri.length).split('/');
      let name = null;
      if (component) {
        [, component, , name] = component.match(/([^?(]+)(\('([^)]+)'\))?(\?(.+))?/);
      }
      console.log(`component: '${component}' command: '${command}, name: '${name}`);

      switch (component) {
        case 'worksheets':
          if (command) {
            const sheet = data.sheets.find((s) => (s.name === command));
            console.log(sheet);
            if (!sheet) {
              throw StatusCodeError(command, 404);
            }
            return sheet;
          } else {
            return { value: data.sheets.map((sheet) => ({ name: sheet.name })) };
          }
        case 'tables':
          return { value: data.tables.map((table) => ({ name: table.name })) };
        case 'names':
          return namedItemOps(data.namedItems)({
            command,
            method,
            body,
          });
        default:
          return { values: data.name };
      }
    };
    f.get = (uri) => f({ method: 'GET', uri });
    return f;
  }

}

module.exports = MockOneDrive;
