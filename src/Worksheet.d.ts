/*
 * Copyright 2019 Adobe. All rights reserved.
 * This file is licensed to you under the Apache License, Version 2.0 (the 'License');
 * you may not use this file except in compliance with the License. You may obtain a copy
 * of the License at http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software distributed under
 * the License is distributed on an 'AS IS' BASIS, WITHOUT WARRANTIES OR REPRESENTATIONS
 * OF ANY KIND, either express or implied. See the License for the specific language
 * governing permissions and limitations under the License.
 */
import { NamedItem } from './NamedItem';
import { GraphResult } from './OneDrive';
import { Table } from './Table';
import { Range } from './Range';

/**
 * Excel work sheet
 */
export declare interface Worksheet {
  /**
   * Return the named items in a work sheet
   * @returns array of named items when resolved
   */
  getNamedItems(): Promise<NamedItem[]>;

  /**
   * Return a named item
   * @param {string} name name
   * @returns named item
   */
  getNamedItem(name: string): Promise<NamedItem>;

  /**
   * Add a named item
   * @param name name
   * @param reference reference
   * @param comment comment
   */
  addNamedItem(name: string, reference: string, comment: string): Promise<GraphResult>;

  /**
   * Delete a named item.
   * @param name name
   */
  deleteNamedItem(name: string): Promise<void>;

  /**
   * Return the table names contained in a work book.
   * @returns array of table names when resolved
   */
  getTableNames(): Promise<string[]>;

  /**
   * Return a new `Table` instance given its name
   * @param name table name
   */
  table(name: string): Table;

  /**
   * Returns the use name.
   */
  getUsedName(): Promise<any>;

  /**
   * Returns a new range object that reflects the `usedRange` of a work sheet.
   */
  usedRange(): Range;

  /**
   * Returns a new range object that spans the address given
   * @param address address, e.g. A1:C2
   */
   range(address): Range;
}
