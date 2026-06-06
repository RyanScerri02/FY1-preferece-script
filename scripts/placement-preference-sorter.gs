/**
 * Google Sheets placement preference sorter.
 *
 * Non-coders can customise this from the sheet itself:
 * - Placement Sorter > Create / Refresh Setup Sheets
 * - Edit the Sorter Settings tab
 * - Reorder the Preferences tab from most preferred to least preferred
 * - Placement Sorter > Sort Programmes by Preferences
 */

const DEFAULT_SETTINGS = {
  settingsSheetName: 'Sorter Settings',
  preferencesSheetName: 'Preferences',
  targetSheetName: '', // blank = first normal sheet that is not Settings/Preferences
  firstDataRow: 3,
  rankColumn: 'E',
  codeColumn: 'F',
  firstPlacementColumn: 'G',
  placementColumnCount: 4,
  rankHeaderCell: 'E1',
  preferenceWeightPower: 2,
  preserveFormatting: 'Yes',
};

const SETTINGS_ROWS = [
  ['Target sheet name', 'targetSheetName', DEFAULT_SETTINGS.targetSheetName, 'Leave blank to use the first normal sheet. Put the exact tab name here if needed.'],
  ['First data row', 'firstDataRow', DEFAULT_SETTINGS.firstDataRow, 'The first row containing a programme, for example 6.'],
  ['Rank column', 'rankColumn', DEFAULT_SETTINGS.rankColumn, 'Column where ranks should be written, for example E.'],
  ['Programme code column', 'codeColumn', DEFAULT_SETTINGS.codeColumn, 'Column containing codes such as FY1-26-001, for example F.'],
  ['First placement column', 'firstPlacementColumn', DEFAULT_SETTINGS.firstPlacementColumn, 'First quarter placement column, for example G.'],
  ['Preference weighting strength', 'preferenceWeightPower', DEFAULT_SETTINGS.preferenceWeightPower, '1 = gentle, 2 = balanced, 3 = strong, 4 = very strong. default is 2. Higher values give more advantage to getting a top preference.'],
  ['Preserve colours and text styling', 'preserveFormatting', DEFAULT_SETTINGS.preserveFormatting, 'Recommended to keep as Yes. Only set to No if script is taking too long to run.'],
];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Placement Sorter')
    .addItem('Create / Refresh Setup Sheets', 'createSetupSheets')
    .addItem('Sort Programmes by Preferences', 'sortProgrammesByPreferences')
    .addToUi();
}

function createSetupSheets() {
  createSettingsSheet_();
  createPreferencesSheet();
  SpreadsheetApp.getUi().alert('Setup sheets are ready. Edit Sorter Settings if needed, reorder Preferences from most preferred to least preferred, then run Sort Programmes by Preferences.');
}

function createSettingsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(DEFAULT_SETTINGS.settingsSheetName);

  if (!sheet) {
    sheet = ss.insertSheet(DEFAULT_SETTINGS.settingsSheetName);
  }

  const existing = readSettingsSheetValues_(sheet);
  const rows = SETTINGS_ROWS.map(([label, key, defaultValue, help]) => [
    label,
    existing[key] !== undefined && existing[key] !== '' ? existing[key] : defaultValue,
    help,
  ]);

  sheet.clear();
  sheet.getRange('A1').setValue('Setting');
  sheet.getRange('B1').setValue('Value');
  sheet.getRange('C1').setValue('Help');
  sheet.getRange('A1:C1').setFontWeight('bold').setBackground('#d9ead3');
  sheet.getRange(2, 1, rows.length, 3).setValues(rows);
  sheet.getRange('A:C').setWrap(true);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 3);
}

function createPreferencesSheet() {
  const settings = getSettings_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(settings.preferencesSheetName);

  if (!sheet) {
    sheet = ss.insertSheet(settings.preferencesSheetName);
  }

  const existingPreferences = readPreferences_(settings);
  const placements = getUniquePlacements_(settings);
  const orderedPlacements = mergeExistingAndNewPreferences_(existingPreferences, placements);

  sheet.clear();
  sheet.getRange('A1').setValue('Rank');
  sheet.getRange('B1').setValue('Placement');
  sheet.getRange('C1').setValue('Instructions');
  sheet.getRange('A1:C1').setFontWeight('bold').setBackground('#d9ead3');
  sheet.getRange('C2').setValue('Reorder column B so your favourite placement is at the top. Leave no blank rows in the preference list.');
  sheet.getRange('C2').setWrap(true);

  if (orderedPlacements.length) {
    sheet.getRange(2, 1, orderedPlacements.length, 1).setValues(orderedPlacements.map((_, index) => [index + 1]));
    sheet.getRange(2, 2, orderedPlacements.length, 1).setValues(orderedPlacements.map(value => [value]));
  }

  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 3);
}

function sortProgrammesByPreferences() {
  const settings = getSettings_();
  const sheet = getTargetSheet_(settings);
  const lastRow = findLastDataRow_(sheet, settings);

  if (lastRow < settings.firstDataRow) {
    throw new Error('No programme rows were found. Check Sorter Settings > First data row.');
  }

  const preferences = readPreferences_(settings);
  if (preferences.length === 0) {
    createPreferencesSheet();
    throw new Error('No preferences found yet. Fill the Preferences sheet first, then sort again.');
  }

  const preferenceMap = buildPreferenceMap_(preferences);
  const rankColumn = columnToNumber_(settings.rankColumn);
  const dataColumnCount = settings.placementColumnCount + 2; // rank + code + placements
  const range = sheet.getRange(settings.firstDataRow, rankColumn, lastRow - settings.firstDataRow + 1, dataColumnCount);
  const values = range.getValues();
  const formatting = settings.preserveFormatting ? captureFormatting_(range) : null;

  const rows = values
    .filter(row => row[1] !== '')
    .map((row, index) => {
      const placements = row.slice(2, 2 + settings.placementColumnCount);
      const placementRanks = placements.map(name => getPlacementRank_(name, preferenceMap, preferences.length));
      const sortedPlacementRanks = [...placementRanks].sort((a, b) => a - b);

      return {
        row,
        style: formatting ? getFormattingRow_(formatting, index) : null,
        score: placementRanks.reduce((sum, rank) => sum + getWeightedPreferenceScore_(rank, preferences.length, settings), 0),
        tieBreakers: sortedPlacementRanks,
        code: String(row[1]),
      };
    });

  rows.sort(compareProgrammeRows_);

  rows.forEach((item, index) => {
    item.row[0] = index + 1;
  });

  const outputRange = sheet.getRange(settings.firstDataRow, rankColumn, rows.length, dataColumnCount);
  outputRange.setValues(rows.map(item => item.row));

  if (formatting) {
    applyFormatting_(outputRange, rows.map(item => item.style));
  }

  sheet.getRange(settings.rankHeaderCell).setValue('RANK YOUR\nPREFERENCE');
  sheet.getRange(settings.firstDataRow, rankColumn, rows.length, 1).setHorizontalAlignment('center');
}

function getSettings_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DEFAULT_SETTINGS.settingsSheetName);
  const rawSettings = sheet ? readSettingsSheetValues_(sheet) : {};

  return {
    settingsSheetName: DEFAULT_SETTINGS.settingsSheetName,
    preferencesSheetName: DEFAULT_SETTINGS.preferencesSheetName,
    targetSheetName: String(rawSettings.targetSheetName || DEFAULT_SETTINGS.targetSheetName).trim(),
    firstDataRow: toPositiveInteger_(rawSettings.firstDataRow, DEFAULT_SETTINGS.firstDataRow),
    rankColumn: normalizeColumnLetter_(rawSettings.rankColumn || DEFAULT_SETTINGS.rankColumn),
    codeColumn: normalizeColumnLetter_(rawSettings.codeColumn || DEFAULT_SETTINGS.codeColumn),
    firstPlacementColumn: normalizeColumnLetter_(rawSettings.firstPlacementColumn || DEFAULT_SETTINGS.firstPlacementColumn),
    placementColumnCount: toPositiveInteger_(rawSettings.placementColumnCount, DEFAULT_SETTINGS.placementColumnCount),
    rankHeaderCell: String(rawSettings.rankHeaderCell || DEFAULT_SETTINGS.rankHeaderCell).trim(),
    preferenceWeightPower: toNumberInRange_(rawSettings.preferenceWeightPower, DEFAULT_SETTINGS.preferenceWeightPower, 1, 6),
    preserveFormatting: normalizeYesNo_(rawSettings.preserveFormatting || DEFAULT_SETTINGS.preserveFormatting),
  };
}

function readSettingsSheetValues_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {};

  const labelToKey = SETTINGS_ROWS.reduce((map, [label, key]) => {
    map[normalizeSettingLabel_(label)] = key;
    return map;
  }, {});

  return sheet
    .getRange(2, 1, lastRow - 1, 2)
    .getValues()
    .reduce((settings, [label, value]) => {
      const key = labelToKey[normalizeSettingLabel_(label)];
      if (key) settings[key] = value;
      return settings;
    }, {});
}

function readPreferences_(settings) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const prefSheet = ss.getSheetByName(settings.preferencesSheetName);
  if (!prefSheet) return [];

  const lastRow = prefSheet.getLastRow();
  if (lastRow < 2) return [];

  return prefSheet
    .getRange(2, 2, lastRow - 1, 1)
    .getValues()
    .flat()
    .map(value => String(value).trim())
    .filter(Boolean);
}

function getTargetSheet_(settings) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (settings.targetSheetName) {
    const namedSheet = ss.getSheetByName(settings.targetSheetName);
    if (!namedSheet) throw new Error(`Target sheet not found: ${settings.targetSheetName}`);
    return namedSheet;
  }

  const activeSheet = ss.getActiveSheet();
  if (isNormalDataSheet_(activeSheet, settings)) return activeSheet;

  const sheet = ss.getSheets().find(candidate => isNormalDataSheet_(candidate, settings));
  if (!sheet) throw new Error('No data sheet found. Add the target tab name in Sorter Settings.');
  return sheet;
}

function isNormalDataSheet_(sheet, settings) {
  const name = sheet.getName();
  return name !== settings.settingsSheetName && name !== settings.preferencesSheetName;
}

function buildPreferenceMap_(preferences) {
  return preferences.reduce((map, name, index) => {
    map[normalizePlacementName_(name)] = index + 1;
    return map;
  }, {});
}

function getPlacementRank_(name, preferenceMap, preferenceCount) {
  const normalized = normalizePlacementName_(name);
  return preferenceMap[normalized] || preferenceCount + 100;
}

function getWeightedPreferenceScore_(rank, preferenceCount, settings) {
  if (rank > preferenceCount) return 0;

  const inverseRank = preferenceCount - rank + 1;
  return Math.pow(inverseRank, settings.preferenceWeightPower);
}

function normalizePlacementName_(name) {
  return String(name || '')
    .toLowerCase()
    .replace(/&/g, 'and')
    .replace(/\s+/g, ' ')
    .trim();
}

function compareProgrammeRows_(a, b) {
  if (a.score !== b.score) return b.score - a.score;

  for (let i = 0; i < Math.max(a.tieBreakers.length, b.tieBreakers.length); i++) {
    const left = a.tieBreakers[i] || 9999;
    const right = b.tieBreakers[i] || 9999;
    if (left !== right) return left - right;
  }

  return a.code.localeCompare(b.code, undefined, { numeric: true, sensitivity: 'base' });
}

function findLastDataRow_(sheet, settings) {
  const lastRow = sheet.getLastRow();
  if (lastRow < settings.firstDataRow) return lastRow;

  const codeColumn = columnToNumber_(settings.codeColumn);
  const codeValues = sheet
    .getRange(settings.firstDataRow, codeColumn, lastRow - settings.firstDataRow + 1, 1)
    .getValues();

  for (let index = codeValues.length - 1; index >= 0; index--) {
    if (codeValues[index][0] !== '') {
      return settings.firstDataRow + index;
    }
  }

  return settings.firstDataRow - 1;
}

function getUniquePlacements_(settings) {
  const sheet = getTargetSheet_(settings);
  const lastRow = findLastDataRow_(sheet, settings);
  if (lastRow < settings.firstDataRow) return [];

  const firstPlacementColumn = columnToNumber_(settings.firstPlacementColumn);
  const values = sheet
    .getRange(settings.firstDataRow, firstPlacementColumn, lastRow - settings.firstDataRow + 1, settings.placementColumnCount)
    .getValues()
    .flat()
    .map(value => String(value).trim())
    .filter(Boolean);

  return [...new Set(values)].sort((a, b) => a.localeCompare(b));
}

function mergeExistingAndNewPreferences_(existingPreferences, detectedPlacements) {
  const seen = {};
  const merged = [];

  existingPreferences.concat(detectedPlacements).forEach(placement => {
    const key = normalizePlacementName_(placement);
    if (!seen[key]) {
      seen[key] = true;
      merged.push(placement);
    }
  });

  return merged;
}

function captureFormatting_(range) {
  return {
    backgrounds: range.getBackgrounds(),
    fontColors: range.getFontColors(),
    fontWeights: range.getFontWeights(),
    fontStyles: range.getFontStyles(),
    horizontalAlignments: range.getHorizontalAlignments(),
  };
}

function getFormattingRow_(formatting, index) {
  return {
    backgrounds: formatting.backgrounds[index],
    fontColors: formatting.fontColors[index],
    fontWeights: formatting.fontWeights[index],
    fontStyles: formatting.fontStyles[index],
    horizontalAlignments: formatting.horizontalAlignments[index],
  };
}

function applyFormatting_(range, formattingRows) {
  range.setBackgrounds(formattingRows.map(row => row.backgrounds));
  range.setFontColors(formattingRows.map(row => row.fontColors));
  range.setFontWeights(formattingRows.map(row => row.fontWeights));
  range.setFontStyles(formattingRows.map(row => row.fontStyles));
  range.setHorizontalAlignments(formattingRows.map(row => row.horizontalAlignments));
}

function columnToNumber_(columnLetter) {
  const column = normalizeColumnLetter_(columnLetter);
  let number = 0;

  for (let i = 0; i < column.length; i++) {
    number = number * 26 + column.charCodeAt(i) - 64;
  }

  return number;
}

function normalizeColumnLetter_(value) {
  const column = String(value || '').trim().toUpperCase().replace(/[^A-Z]/g, '');
  if (!column) throw new Error('A column setting is blank. Check Sorter Settings.');
  return column;
}

function normalizeSettingLabel_(value) {
  return String(value || '').toLowerCase().replace(/\s+/g, ' ').trim();
}

function normalizeYesNo_(value) {
  return String(value || '').trim().toLowerCase() !== 'no';
}

function toPositiveInteger_(value, fallback) {
  const number = parseInt(value, 10);
  return Number.isFinite(number) && number > 0 ? number : fallback;
}

function toNumberInRange_(value, fallback, min, max) {
  const number = Number(value);
  if (!Number.isFinite(number)) return fallback;
  return Math.min(max, Math.max(min, number));
}
