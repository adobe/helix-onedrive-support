/*
 * Copyright 2022 Adobe. All rights reserved.
 * This file is licensed to you under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License. You may obtain a copy
 * of the License at http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software distributed under
 * the License is distributed on an "AS IS" BASIS, WITHOUT WARRANTIES OR REPRESENTATIONS
 * OF ANY KIND, either express or implied. See the License for the specific language
 * governing permissions and limitations under the License.
 */

import { GraphResult, Logger } from '../OneDrive'
import { FormatOptions } from "./Table";

/**
 * Excel Range
 */
export declare interface Range {

  uri:string;

  log:Logger;

  /**
   * Returns the graph result of the range
   */
  getData(): Promise<GraphResult>;

  /**
   * Returns the range address
   */
  getAddress(): Promise<string>;

  /**
   * Returns the range local address
   */
  getAddressLocal(): Promise<string>;

  /**
   * Returns the column names of the range
   */
  getColumnNames(): Promise<string[]>;

  /**
   * Returns the rows as a list of objects. the rows have the columns names as property names
   * and the row values as value.
   */
  getRowsAsObjects(opts?:FormatOptions): Promise<Array<object>>;

  /**
   * Returns the values of the range.
   */
  getValues(): Promise<Array<object>>;

  /**
   * Updates the range with new values.
   * @param values new values, that may contain a combination of properties values,
   *               formulas and numberFormat. Those properties should have the same
   *               array dimension as the range addressed
   */
  update(values: object): Promise<void>;

  /**
   * Deletes the cells associated with the range.
   * @param shiftValue Specifies which way to shift the cells. 
   *                   The possible values are: Up, Left.
   */
  delete(shiftValue: string): Promise<void>;

  /**
   * Inserts a cell or a range of cells into the worksheet in place of this range, 
   * and shifts the other cells to make space.
   * @param shiftValue Specifies which way to shift the cells. 
   *                   The possible values are: Down, Right.
   */
  insert(shiftValue: string): Promise<void>;
}
