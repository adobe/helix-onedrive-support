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
const { OneDrive } = require('./OneDrive.js');
const { OneDriveAuth } = require('./OneDriveAuth.js');
const Workbook = require('./Workbook.js');
const StatusCodeError = require('./StatusCodeError.js');

/**
 * Handle the `namedItems` operation on a workbook / worksheet
 * @param {object} sheet The mock data
 * @param {string[]} segs Array of path segments
 * @param {string} method Request method
 * @param {object} body Request body
 * @returns {object} The response value
 */
function handleNamedItems(sheet, segs, method, body) {
  const { namedItems } = sheet;
  const command = segs.shift();
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
}

/**
 * Handle the `table` operation on a workbook / worksheet
 * @param {object} sheet The mock data
 * @param {string[]} segs Array of path segments
 * @param {string} method Request method
 * @param {object} body Request body
 * @returns {object} The response value
 */
function handleTable(sheet, segs, method, body) {
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
        if (method === 'DELETE') {
          table.rows.splice(index, 1);
          return null;
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
}

/**
 * Mock OneDrive client that supports some of the operations the `OneDrive` class does.
 */
class OneDriveMock extends OneDrive {
  constructor({ auth } = {}) {
    if (!auth) {
      // eslint-disable-next-line no-param-reassign
      auth = new OneDriveAuth({
        clientId: 'mock-id',
        tenant: 'test-tenant',
      });
      // eslint-disable-next-line no-param-reassign
      auth.accessToken = {
        accessToken: 'dummy-token',
        tenantId: 'test-tenant',
      };
    }
    super({
      auth,
    });
    this.workbooks = [];
    this.sharelinks = {};
    this.driveItems = {};
  }

  /**
   * Register a mock workbook.
   *
   * @param {string} driveId The drive id
   * @param {string} itemId the item id
   * @param {object} data Mock workbook data
   * @returns {OneDriveMock} this
   */
  registerWorkbook(driveId, itemId, data) {
    this.workbooks.push({
      uri: `/drives/${driveId}/items/${itemId}/workbook`,
      // poor mans deep clone
      data: JSON.parse(JSON.stringify(data)),
    });
    return this;
  }

  /**
   * Registers a mock drive item
   * @param {string} driveId The drive id
   * @param {string} itemId the item id
   * @param {object} data Mock item data
   * @returns {OneDriveMock} this
   */
  registerDriveItem(driveId, itemId, data) {
    this.driveItems[`/drives/${driveId}/items/${itemId}`] = data;
    return this;
  }

  /**
   * Registers a mock drive item child list
   * @param {string} driveId The drive id
   * @param {string} itemId the item id
   * @param {object} data Mock item data
   * @returns {OneDriveMock} this
   */
  registerDriveItemChildren(driveId, itemId, data) {
    this.driveItems[`/drives/${driveId}/items/${itemId}/children`] = data;
    return this;
  }

  /**
   * Register a mock sharelink.
   *
   * @param {string} uri The sharelink uri
   * @param {string} driveId The drive id
   * @param {string} itemId The the item id
   * @returns {OneDriveMock} this;
   */
  registerShareLink(uri, driveId, itemId) {
    this.sharelinks[uri] = {
      parentReference: {
        driveId,
      },
      id: itemId,
    };
    return this;
  }

  /**
   * @see OneDrive#getDriveItemFromShareLink
   */
  async getDriveItemFromShareLink(uri) {
    let driveItem = OneDriveMock.driveItemFromURL(uri);
    if (!driveItem) {
      driveItem = this.sharelinks[uri];
    }
    if (!driveItem) {
      throw new StatusCodeError(uri, 404);
    }
    return driveItem;
  }

  /**
   * @see OneDrive#getWorkbook
   */
  getWorkbook(driveItem) {
    const uri = driveItem
      ? `/drives/${driveItem.parentReference.driveId}/items/${driveItem.id}/workbook`
      : this.workbooks[0].uri;
    return new Workbook(this, uri, console);
  }

  /**
   * @see OneDrive#doFetch
   */
  doFetch(uri, _, { method = 'GET', body } = {}) {
    const url = new URL(`https://dummy.org${uri}`);
    if (url.pathname in this.driveItems) {
      const result = this.driveItems[url.pathname];
      if (!Array.isArray(result.value)) {
        return result;
      }
      const data = result.value;
      const max = Number.parseInt(url.searchParams.get('$top') || 200, 10);
      // note that we abuse the skiptoken a `skip` param here and totally ignore the real `$skip`
      const skiptoken = Number.parseInt(url.searchParams.get('$skiptoken') || 0, 10);
      const len = data.length - skiptoken;
      if (len > max) {
        url.searchParams.set('$skiptoken', skiptoken + max);
        return {
          value: data.slice(skiptoken, skiptoken + max),
          '@odata.nextLink': url.toString(),
        };
      } else if (skiptoken) {
        return {
          value: data.slice(skiptoken, skiptoken + len),
        };
      } else {
        return result;
      }
    }
    const wb = this.workbooks.find((w) => (uri.startsWith(w.uri)));
    if (!wb) {
      throw new StatusCodeError('not found', 404);
    }
    const { data } = wb;

    // eslint-disable-next-line no-unused-vars
    const [path, query] = uri.substring(wb.uri.length).split('?');
    const segs = path.split('/');
    segs.shift();

    // default the sheet to the entire data
    let sheet = data;

    // handle the '/workbook/worksheets/:name' portion
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
          return { value: sheet };
        }
      } else {
        return { value: data.sheets.map((st) => ({ name: st.name })) };
      }
    }

    // handle the operations on the workbook / worksheet
    switch (segs.shift()) {
      case 'usedRange':
        return sheet.usedRange;
      case 'tables':
        return handleTable(sheet, segs, method, body);
      case 'names':
        return handleNamedItems(sheet, segs, method, body);
      default:
        // default return the data
        return { value: data };
    }
  }
}

module.exports = {
  OneDriveMock,
};
