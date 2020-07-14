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

class MockOneDrive {
  constructor() {
    this.workbooks = [];
  }

  registerWorkbook(driveId, itemId, data) {
    this.workbooks.push({
      uri: `/drives/${driveId}/items/${itemId}/workbook`,
      // poor mans deep clone
      data: JSON.parse(JSON.stringify(data)),
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

      // eslint-disable-next-line no-unused-vars
      const [path, query] = uri.substring(wb.uri.length).split('?');
      const segs = path.split('/');
      segs.shift();

      let sheet = data;
      if (segs[0] === 'worksheets') {
        segs.shift();
        if (segs[0]) {
          const sheetName = segs.shift();
          sheet = data.sheets.find((s) => (s.name === sheetName));
          if (!sheet) {
            throw new StatusCodeError(sheetName, 404);
          }
          if (!segs[0]) {
            // if no more segments, return the sheet data
            return sheet;
          }
        } else {
          return { value: data.sheets.map((st) => ({ name: st.name })) };
        }
      }

      switch (segs.shift()) {
        case 'usedRange':
          return sheet.usedRange;
        case 'tables':
          if (segs[0]) {
            const tableName = segs.shift();
            const table = sheet.tables.find((t) => t.name === tableName);
            if (!table) {
              throw new StatusCodeError(tableName, 404);
            }
            let command;
            let name;
            if (segs[0]) {
              [, command, , name] = segs.shift().match(/([^?(]+)(\('([^)]+)'\))?(\?(.+))?/);
            }
            switch (command) {
              case 'dataBodyRange':
                return { rowCount: table.rows.length };
              case 'headerRowRange':
                return { values: [table.headerNames] };
              case 'rows': {
                if (!segs[0]) {
                  return { value: table.rows.map((row) => ({ values: [row] })) };
                }
                const subCommand = segs.shift();
                if (subCommand === 'add') {
                  table.rows.push(...body.values);
                  return { index: table.rows.length - 1 };
                }
                const index = parseInt(subCommand.replace(/itemAt\(index=([0-9]+)\)/, '$1'), 10);
                if (index < 0 || index >= table.rows.length) {
                  throw new StatusCodeError(`Index out of range: ${index}`, 400);
                }
                if (body) {
                  [table.rows[index]] = body.values;
                }
                return { values: [table.rows[index]] };
              }
              case 'columns': {
                if (!name) {
                  const cols = table.headerNames.map((n) => ({
                    name: n,
                    values: [[n]],
                  }));
                  table.rows.forEach((row) => {
                    row.forEach((value, idx) => {
                      cols[idx].values.push([value]);
                    });
                  });
                  return {
                    value: cols,
                  };
                }
                const columnName = name;
                const index = table.headerNames.findIndex((n) => n === columnName);
                if (index === -1) {
                  throw new StatusCodeError(`Column name not found: ${columnName}`, 400);
                }
                return {
                  values: [
                    [table.headerNames[index]],
                    ...table.rows.map((row) => [row[index]]),
                  ],
                };
              }

              default:
                if (body) {
                  table.name = body.name;
                }
                return { values: table.name };
            }
          } else {
            return { value: sheet.tables.map((table) => ({ name: table.name })) };
          }
        case 'names':
          return namedItemOps(sheet.namedItems)({
            command: segs[0],
            method,
            body,
          });
        default:
          return { values: sheet.name };
      }
    };
    f.get = (uri) => f({ method: 'GET', uri });
    return f;
  }
}

module.exports = MockOneDrive;
