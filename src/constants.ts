/**
 * Copyright 2023 Google LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *       http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
export const GOAL_EVERYDAY = 3000;
export const GOAL_UNTIL_NOV = 100_000;
export const GOAL_UNTIL_NEXT_YEAR = 300_000;
export const GOAL_CELL = '$G$2';
export const GOAL_CELLS_RANGE = 'G1:I3';

export const NOVEL_FOLDER_ID = '1QkI_Gy8AQu78p_bKA0P0IqmfdqukER_W';

export const DIARY_SHEET = {
  NAME: "日付ごと",
  COL_DATE: 1,// 日付列
  COL_DIFFERENCE: 2,// 差分列
  COL_COUNT: 3,// 文字数列
};

export const HOURLY_SHEET = {
  NAME: "時間ごと",
  COL_DATE: 1,// 日付列
  COL_TIME: 2,// 時刻列
  COL_DIFFERENCE: 3,// 差分列
  COL_COUNT: 4,// 文字数列
};

export const DATE_FORMAT = "yyyy/MM/dd";
export const TIME_FORMAT = "HH:mm:ss";
