/*************************************************
 * Academy attendance schedule automation - portfolio-safe version
 *
 * Sensitive information such as student names, academy names, and
 * external attendance spreadsheet IDs has been removed or replaced.
 *************************************************/

const TIMEZONE = 'Asia/Seoul';

const SHEET = {
  COURSE: '\uB2E8\uACFC\uC218\uAC15\uC0DD',
  MANUAL: '\uD0C0\uD559\uC6D0\uC0DD \uAD00\uB9AC',
  CONFLICT: '\uCDA9\uB3CC\uD655\uC778'
};

const ROOM_SHEETS = [
  '301\uD638', '401\uD638', '402\uD638', '403\uD638',
  '501\uD638', '502\uD638', '503\uD638',
  '601\uD638', '602\uD638', '603\uD638',
  '2\uAD00301\uD638', '2\uAD00302\uD638',
  '2\uAD00401\uD638', '2\uAD00402\uD638'
];

const ROOM_SHEET_SET = new Set(ROOM_SHEETS);

const DAY_KEY_TO_OUTPUT_INDEX = {
  A: 0, B: 1, C: 2, D: 3, E: 4, F: 5, G: 6
};

const DAY_KEY_TO_SHEET_COL = {
  A: 8, B: 9, C: 10, D: 11, E: 12, F: 13, G: 14
};

const COURSE_COL = {
  STUDENT_ID: 0,
  NAME: 1,
  GRADE: 2,
  CODE: 3,
  TEACHER: 4,
  SUBJECT: 5,
  START_DATE: 6,
  END_DATE: 7,
  CANCEL_DATE: 8
};

const COURSE_SCHEDULE_DATA_WIDTH = 6;
const COURSE_ATTENDANCE_DATA_WIDTH = 9;

const ATTENDANCE_SYNC_CONFIG = {
  attendanceSpreadsheetId: 'YOUR_ATTENDANCE_SPREADSHEET_ID',
  attendanceSheets: [
    '3\uCE35', '4\uCE35', '5\uCE35',
    '6\uCE35', '2\uAD003\uCE35', '2\uAD004\uCE35'
  ],
  headerRow: 1,
  dataStartRow: 5,
  studentIdCol: 2,
  roomCol: 13,
  periodCol: 16,
  dateStartCol: 17
};

const PROTECTED_VALUES = [
  '1',
  '\uB2A6\uC7A0',
  '\uAD50\uD1B5',
  '\uC9C8\uBCD1',
  '\uBCD1\uC6D0',
  '\uC678\uCD9C',
  '\uC870\uD1F4',
  '\uD559\uAD50',
  '\uD0C0\uD559\uC6D0',
  '\uBD80\uC7AC',
  '\uC9C0\uAC01',
  '\uBCD1\uACB0',
  '\uACB0\uC11D'
];

const PROTECTED_PREFIXES = PROTECTED_VALUES.slice(1);
const PROTECTED_VALUE_SET = new Set(PROTECTED_VALUES);

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('Schedule Automation')
    .addItem('Run all rooms', 'runScheduleAutomation')
    .addItem('Run current room', 'runSingleRoom')
    .addToUi();

  ui.createMenu('Attendance Sync')
    .addItem('Sync attendance sheet', 'syncAttendanceNext7Days')
    .addToUi();
}

function runScheduleAutomation() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const source = readScheduleSourceData_(ss);

  if (!source) {
    ui.alert('Required sheets are missing: course sheet and manual schedule sheet.');
    return;
  }

  const conflictSheet = getOrCreateConflictSheet_(ss);
  resetConflictSheet_(conflictSheet);

  const conflictRows = [];
  let processedRooms = 0;
  const missingRooms = [];

  ROOM_SHEETS.forEach(room => {
    const sh = ss.getSheetByName(room);
    if (!sh) {
      missingRooms.push(room);
      return;
    }

    processOneRoom(sh, source.dataValues, source.manualByRoom[room] || [], conflictRows);
    processedRooms++;
  });

  writeConflictRows_(conflictSheet, conflictRows);
  SpreadsheetApp.flush();

  let message =
    'Schedule automation complete.\n' +
    `Processed rooms: ${processedRooms}\n` +
    `Conflicts: ${conflictRows.length}`;

  if (missingRooms.length) {
    message += `\n\nMissing room sheets: ${missingRooms.join(', ')}`;
  }

  ui.alert(message);
}

function runSingleRoom() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const sh = ss.getActiveSheet();
  const room = sh.getName();

  if (!ROOM_SHEET_SET.has(room)) {
    ui.alert(`'${room}' is not a target room sheet.`);
    return;
  }

  const source = readScheduleSourceData_(ss);
  if (!source) {
    ui.alert('Required sheets are missing: course sheet and manual schedule sheet.');
    return;
  }

  const conflictSheet = getOrCreateConflictSheet_(ss);
  resetConflictSheet_(conflictSheet);

  const conflictRows = [];
  processOneRoom(sh, source.dataValues, source.manualByRoom[room] || [], conflictRows);
  writeConflictRows_(conflictSheet, conflictRows);
  SpreadsheetApp.flush();

  ui.alert(`'${room}' schedule automation complete.\nConflicts: ${conflictRows.length}`);
}

function readScheduleSourceData_(ss) {
  const dataSheet = ss.getSheetByName(SHEET.COURSE);
  const manualSheet = ss.getSheetByName(SHEET.MANUAL);

  if (!dataSheet || !manualSheet) return null;

  return {
    dataValues: getDataRows_(dataSheet, 2, 1, COURSE_SCHEDULE_DATA_WIDTH),
    manualByRoom: buildManualByRoom_(getDataRows_(manualSheet, 2, 1, 4))
  };
}

function readAttendanceSourceData_(ss) {
  const dataSheet = ss.getSheetByName(SHEET.COURSE);
  const manualSheet = ss.getSheetByName(SHEET.MANUAL);

  if (!dataSheet || !manualSheet) return null;

  return {
    dataValues: getDataRows_(dataSheet, 2, 1, COURSE_ATTENDANCE_DATA_WIDTH),
    manualByRoom: buildManualByRoom_(getDataRows_(manualSheet, 2, 1, 4))
  };
}

function getDataRows_(sheet, startRow, startCol, width) {
  const lastRow = sheet.getLastRow();
  if (lastRow < startRow) return [];

  return sheet.getRange(startRow, startCol, lastRow - startRow + 1, width).getValues();
}

function buildManualByRoom_(manualValues) {
  const manualByRoom = Object.create(null);

  manualValues.forEach(row => {
    const room = normalizeRoomName(row[0]);
    const seat = normalizeSeat(row[1]);
    const code = cellText(row[2]);
    const reason = cellText(row[3]);

    if (!room || !seat || !code || !reason) return;

    if (!manualByRoom[room]) manualByRoom[room] = [];
    manualByRoom[room].push({ seat, code, reason });
  });

  return manualByRoom;
}

function getOrCreateConflictSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET.CONFLICT);
  if (!sheet) sheet = ss.insertSheet(SHEET.CONFLICT);
  return sheet;
}

function resetConflictSheet_(sheet) {
  sheet.clear();
  sheet.getRange(1, 1, 1, 6).setValues([
    ['Room', 'Seat', 'Day', 'Period', 'Existing', 'Incoming']
  ]);
}

function writeConflictRows_(sheet, rows) {
  if (!rows.length) return;
  sheet.getRange(2, 1, rows.length, 6).setValues(rows);
}

function processOneRoom(sh, dataValues, manualList, conflictRows) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;

  const n = lastRow - 1;
  const roomName = sh.getName();
  const isAnnex = isAnnexRoom(roomName);

  const ids = sh.getRange(2, 1, n, 1).getValues().map(row => normalizeStudentId(row[0]));
  const seats = sh.getRange(2, 6, n, 1).getValues().map(row => normalizeSeat(row[0]));
  const periods = sh.getRange(2, 7, n, 1).getValues().map(row => normalizePeriod(row[0]));

  const out = Array.from({ length: n }, () => Array(7).fill(''));
  const idIdxMap = buildIndexMap_(ids);
  const seatIdxMap = buildIndexMap_(seats);

  dataValues.forEach(row => {
    const id = normalizeStudentId(row[COURSE_COL.STUDENT_ID]);
    if (!id || !idIdxMap[id]) return;

    const parsed = parseCode(row[COURSE_COL.CODE]);
    if (!parsed.letters || !parsed.letters.length || !parsed.block) return;

    const teacher = cellText(row[COURSE_COL.TEACHER]);
    const subject = abbreviateSubject(cellText(row[COURSE_COL.SUBJECT]));
    const content = [subject, teacher].filter(Boolean).join(' ');
    if (!content) return;

    parsed.letters.forEach(dayKey => {
      const colOff = DAY_KEY_TO_OUTPUT_INDEX[dayKey];
      if (colOff == null) return;

      idIdxMap[id].forEach(i => {
        if (!inBlock(periods[i], parsed.block, isAnnex)) return;

        if (out[i][colOff] && out[i][colOff] !== content) {
          conflictRows.push([roomName, seats[i], dayKey, periods[i], out[i][colOff], content]);
        }

        out[i][colOff] = content;
      });
    });
  });

  manualList.forEach(item => {
    if (!item.seat || !seatIdxMap[item.seat]) return;

    parseManualCodeMulti(item.code).forEach(parsed => {
      const validNums = normalizeManualNumsForRoom(parsed.nums, isAnnex);
      if (!parsed.letters || !validNums.length) return;

      parsed.letters.split('').forEach(dayKey => {
        const colOff = DAY_KEY_TO_OUTPUT_INDEX[dayKey];
        if (colOff == null) return;

        seatIdxMap[item.seat].forEach(i => {
          if (!validNums.some(num => nearlyEqual(periods[i], num))) return;

          if (out[i][colOff] && out[i][colOff] !== item.reason) {
            conflictRows.push([roomName, item.seat, dayKey, periods[i], out[i][colOff], item.reason]);
          }

          out[i][colOff] = item.reason;
        });
      });
    });
  });

  const writeRange = sh.getRange(2, 8, n, 7);
  writeRange.setValues(out);

  if (isAnnex) {
    applyAnnexBackgrounds_(writeRange, periods);
  }
}

function buildIndexMap_(values) {
  const map = Object.create(null);

  values.forEach((value, index) => {
    if (!value) return;
    if (!map[value]) map[value] = [];
    map[value].push(index);
  });

  return map;
}

function applyAnnexBackgrounds_(range, periods) {
  const backgrounds = range.getBackgrounds();

  for (let i = 0; i < periods.length; i++) {
    if (periods[i] >= 9 && periods[i] <= 11) {
      for (let j = 0; j < 7; j++) {
        backgrounds[i][j] = '#e6e6e6';
      }
    }
  }

  range.setBackgrounds(backgrounds);
}

function syncAttendanceNext7Days() {
  const sourceSs = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const startedAt = new Date();

  let targetSs;

  try {
    const attendanceSpreadsheetId = getAttendanceSpreadsheetId_();

    if (!attendanceSpreadsheetId) {
      ui.alert('Attendance spreadsheet ID is not configured in this portfolio version.');
      return;
    }

    targetSs = SpreadsheetApp.openById(attendanceSpreadsheetId);
  } catch (e) {
    ui.alert('Cannot open attendance spreadsheet. Check ID and permissions.');
    return;
  }

  const scheduleMap = buildScheduleMapByRoomAndIdPeriod(sourceSs);

  let touchedSheets = 0;
  let changedCells = 0;
  let matchedRows = 0;
  let checkedRows = 0;
  let targetColsCount = 0;

  ATTENDANCE_SYNC_CONFIG.attendanceSheets.forEach(sheetName => {
    const sh = targetSs.getSheetByName(sheetName);
    if (!sh) return;

    const result = syncOneAttendanceSheetByRoomPeriodFast(sh, scheduleMap);

    if (result.touched) touchedSheets++;
    changedCells += result.changedCells;
    matchedRows += result.matchedRows;
    checkedRows += result.checkedRows;
    targetColsCount += result.targetCols;
  });

  SpreadsheetApp.flush();

  const elapsedSec = ((new Date()) - startedAt) / 1000;
  const targetDateKeys = getNext7DateKeysFromToday();

  ui.alert(
    'Attendance sync complete\n' +
    `Target dates: ${targetDateKeys.join(', ')}\n` +
    `Touched sheets: ${touchedSheets}\n` +
    `Target date columns: ${targetColsCount}\n` +
    `Changed cells: ${changedCells}\n` +
    `Matched rows: ${matchedRows} / ${checkedRows}\n` +
    `Elapsed: ${elapsedSec.toFixed(1)}s`
  );
}

function buildScheduleMapByRoomAndIdPeriod(ss) {
  const result = Object.create(null);
  const source = readAttendanceSourceData_(ss);

  if (!source) return result;

  const courseRowsByStudentId = Object.create(null);

  source.dataValues.forEach(row => {
    const id = normalizeStudentId(row[COURSE_COL.STUDENT_ID]);
    if (!id) return;

    const parsed = parseCode(row[COURSE_COL.CODE]);
    if (!parsed.letters || !parsed.letters.length || !parsed.block) return;

    const teacher = cellText(row[COURSE_COL.TEACHER]);
    const subject = abbreviateSubject(cellText(row[COURSE_COL.SUBJECT]));
    const content = [subject, teacher].filter(Boolean).join(' ');
    if (!content) return;

    const item = {
      id,
      parsed,
      content,
      startDateKey: parseDateKeyFlexible_(row[COURSE_COL.START_DATE]),
      endDateKey: parseDateKeyFlexible_(row[COURSE_COL.END_DATE]),
      cancelDateKeys: parseDateKeyListFlexible_(row[COURSE_COL.CANCEL_DATE])
    };

    if (!courseRowsByStudentId[id]) courseRowsByStudentId[id] = [];
    courseRowsByStudentId[id].push(item);
  });

  ROOM_SHEETS.forEach(roomName => {
    const sh = ss.getSheetByName(roomName);
    if (!sh) return;

    const lastRow = sh.getLastRow();
    if (lastRow < 2) return;

    const isAnnex = isAnnexRoom(roomName);
    const values = sh.getRange(2, 1, lastRow - 1, 7).getValues();
    const roomMap = Object.create(null);
    result[roomName] = roomMap;

    const roomRowsByStudentId = Object.create(null);
    const roomRowsBySeat = Object.create(null);

    values.forEach(row => {
      const studentId = normalizeStudentId(row[0]);
      if (!studentId) return;

      const roomRow = {
        studentId,
        seat: normalizeSeat(row[5]),
        period: normalizePeriod(row[6])
      };

      if (!roomRowsByStudentId[studentId]) roomRowsByStudentId[studentId] = [];
      roomRowsByStudentId[studentId].push(roomRow);

      if (roomRow.seat) {
        if (!roomRowsBySeat[roomRow.seat]) roomRowsBySeat[roomRow.seat] = [];
        roomRowsBySeat[roomRow.seat].push(roomRow);
      }
    });

    function ensureWeekly_(studentId, period) {
      if (!studentId || period === '' || period == null) return null;
      if (!roomMap[studentId]) roomMap[studentId] = Object.create(null);
      if (!roomMap[studentId][period]) {
        roomMap[studentId][period] = Array.from({ length: 7 }, () => ({
          value: '',
          startDateKey: null,
          endDateKey: null,
          cancelDateKeys: []
        }));
      }
      return roomMap[studentId][period];
    }

    function setScheduleCell_(studentId, period, weekdayIndex, value, startDateKey, endDateKey, cancelDateKeys) {
      const weekly = ensureWeekly_(studentId, period);
      if (!weekly) return;

      weekly[weekdayIndex] = {
        value: cellText(value),
        startDateKey: startDateKey || null,
        endDateKey: endDateKey || null,
        cancelDateKeys: cancelDateKeys || []
      };
    }

    Object.keys(roomRowsByStudentId).forEach(studentId => {
      const roomRows = roomRowsByStudentId[studentId];
      const courseRows = courseRowsByStudentId[studentId];
      if (!roomRows || !courseRows) return;

      courseRows.forEach(course => {
        course.parsed.letters.forEach(dayKey => {
          const weekdayIndex = DAY_KEY_TO_OUTPUT_INDEX[dayKey];
          if (weekdayIndex == null) return;

          roomRows.forEach(roomRow => {
            if (!inBlock(roomRow.period, course.parsed.block, isAnnex)) return;

            setScheduleCell_(
              roomRow.studentId,
              roomRow.period,
              weekdayIndex,
              course.content,
              course.startDateKey,
              course.endDateKey,
              course.cancelDateKeys
            );
          });
        });
      });
    });

    const manualList = source.manualByRoom[roomName] || [];

    manualList.forEach(item => {
      if (!item.seat) return;

      const matchedSeatRows = roomRowsBySeat[item.seat];
      if (!matchedSeatRows || !matchedSeatRows.length) return;

      parseManualCodeMulti(item.code).forEach(parsed => {
        const validNums = normalizeManualNumsForRoom(parsed.nums, isAnnex);
        if (!parsed.letters || !validNums.length) return;

        parsed.letters.split('').forEach(dayKey => {
          const weekdayIndex = DAY_KEY_TO_OUTPUT_INDEX[dayKey];
          if (weekdayIndex == null) return;

          matchedSeatRows.forEach(roomRow => {
            if (!validNums.some(num => nearlyEqual(roomRow.period, num))) return;

            setScheduleCell_(
              roomRow.studentId,
              roomRow.period,
              weekdayIndex,
              item.reason,
              null,
              null,
              []
            );
          });
        });
      });
    });
  });

  return result;
}

function syncOneAttendanceSheetByRoomPeriodFast(sh, scheduleMap) {
  const cfg = ATTENDANCE_SYNC_CONFIG;
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();

  if (lastRow < cfg.dataStartRow || lastCol < cfg.dateStartCol) {
    return { touched: false, changedCells: 0, matchedRows: 0, checkedRows: 0, targetCols: 0 };
  }

  const readStartCol = cfg.studentIdCol;
  const readWidth = lastCol - readStartCol + 1;
  const allValues = sh.getRange(1, readStartCol, lastRow, readWidth).getValues();

  const headers = allValues[cfg.headerRow - 1];
  const data = allValues.slice(cfg.dataStartRow - 1);

  const idxStudent = cfg.studentIdCol - readStartCol;
  const idxRoom = cfg.roomCol - readStartCol;
  const idxPeriod = cfg.periodCol - readStartCol;
  const idxDate0 = cfg.dateStartCol - readStartCol;

  const targetDateKeys = getNext7DateKeysFromToday();
  const targetDateSet = new Set(targetDateKeys);
  const targetCols = [];

  for (let c = idxDate0; c < headers.length; c++) {
    const key = extractDateKeyFromHeader(headers[c], targetDateSet);
    if (!key || !targetDateSet.has(key)) continue;

    targetCols.push({
      idx: c,
      weekdayIndex: getWeekdayIndexFromDateKey(key),
      dateKey: key
    });
  }

  if (!targetCols.length) {
    return { touched: false, changedCells: 0, matchedRows: 0, checkedRows: data.length, targetCols: 0 };
  }

  targetCols.sort((a, b) => a.idx - b.idx);

  let touched = false;
  let changedCells = 0;
  let matchedRows = 0;
  const changedTargetIdxSet = new Set();

  for (let r = 0; r < data.length; r++) {
    const row = data[r];
    const studentId = normalizeStudentId(row[idxStudent]);
    const room = normalizeRoomName(row[idxRoom]);
    const period = normalizePeriod(row[idxPeriod]);

    if (!studentId || !room) continue;
    if (period === '' || period == null) continue;

    const weekly =
      scheduleMap[room] &&
      scheduleMap[room][studentId] &&
      scheduleMap[room][studentId][period];

    if (!weekly) continue;
    matchedRows++;

    for (let i = 0; i < targetCols.length; i++) {
      const tc = targetCols[i];
      const scheduleCell = weekly[tc.weekdayIndex];
      const currentValue = cellText(row[tc.idx]);
      const newValue = scheduleCell ? cellText(scheduleCell.value) : '';
      const startDateKey = scheduleCell ? scheduleCell.startDateKey : null;
      const endDateKey = scheduleCell ? scheduleCell.endDateKey : null;
      const cancelDateKeys = scheduleCell ? (scheduleCell.cancelDateKeys || []) : [];

      if (isProtectedAttendanceValue(currentValue)) continue;

      if (startDateKey && tc.dateKey < startDateKey) {
        if (isAutoScheduleValue(currentValue)) {
          row[tc.idx] = '';
          changedCells++;
          touched = true;
          changedTargetIdxSet.add(tc.idx);
        }
        continue;
      }

      if (endDateKey && tc.dateKey > endDateKey) {
        if (isAutoScheduleValue(currentValue)) {
          row[tc.idx] = '';
          changedCells++;
          touched = true;
          changedTargetIdxSet.add(tc.idx);
        }
        continue;
      }

      if (cancelDateKeys.indexOf(tc.dateKey) !== -1) {
        if (isAutoScheduleValue(currentValue)) {
          row[tc.idx] = '';
          changedCells++;
          touched = true;
          changedTargetIdxSet.add(tc.idx);
        }
        continue;
      }

      if (!newValue || newValue === '-') {
        if (isAutoScheduleValue(currentValue)) {
          row[tc.idx] = '';
          changedCells++;
          touched = true;
          changedTargetIdxSet.add(tc.idx);
        }
        continue;
      }

      if (!currentValue || currentValue === '-') {
        row[tc.idx] = newValue;
        changedCells++;
        touched = true;
        changedTargetIdxSet.add(tc.idx);
        continue;
      }

      if (isAutoScheduleValue(currentValue) && currentValue !== newValue) {
        row[tc.idx] = newValue;
        changedCells++;
        touched = true;
        changedTargetIdxSet.add(tc.idx);
      }
    }
  }

  if (touched) {
    targetCols.forEach(tc => {
      if (!changedTargetIdxSet.has(tc.idx)) return;

      const out = data.map(row => [row[tc.idx]]);
      sh.getRange(cfg.dataStartRow, readStartCol + tc.idx, out.length, 1).setValues(out);
    });
  }

  return {
    touched,
    changedCells,
    matchedRows,
    checkedRows: data.length,
    targetCols: targetCols.length
  };
}

function parseCode(code) {
  const m = cellText(code).match(/([A-G]{1,})([1-3])/i);
  return m ? { letters: m[1].toUpperCase().split(''), block: m[2] } : {};
}

function inBlock(period, block, isAnnex) {
  const p = Number(period);
  if (p == null || isNaN(p)) return false;

  if (!isAnnex) {
    if (block === '1') return p >= 2 && p <= 3;
    if (block === '2') return p >= 4 && p <= 6;
    return p >= 7 && p <= 8;
  }

  if (block === '1') return (p >= 2 && p <= 3) || nearlyEqual(p, 3.5);
  if (block === '2') return p >= 4 && p <= 6;
  return (p >= 7 && p <= 8) || nearlyEqual(p, 8.5);
}

function parseManualCode(code) {
  let s = normalizeKoreanDaysToLetters(expandKoreanDayRanges(code)).toUpperCase();

  let letters = '';
  for (const ch of s) {
    if (DAY_KEY_TO_OUTPUT_INDEX[ch] != null) letters += ch;
  }

  const numericRaw = s
    .replace(/[A-G]/g, '')
    .replace(/[~\uFF5E\uFF0D\u2013\u2014-]/g, '~')
    .replace(/[^0-9.~]/g, '');

  let nums = [];

  if (numericRaw.includes('~')) {
    const parts = numericRaw.split('~').filter(Boolean);
    if (parts.length === 2) {
      const start = Number(parts[0]);
      const end = Number(parts[1]);

      if (!isNaN(start) && !isNaN(end) && start <= end) {
        for (let p = start; p <= end + 1e-9; p += 0.5) {
          nums.push(Number(p.toFixed(1)));
        }
      }
    }
  }

  if (nums.length === 0) {
    const tokens = numericRaw.match(/8\.5|3\.5|11|10|[1-9]/g) || [];
    nums = tokens.map(t => Number(t));
  }

  return { letters, nums: uniqueNumbers_(nums) };
}

function parseManualCodeMulti(code) {
  return cellText(code)
    .split(/\s*,\s*/)
    .filter(Boolean)
    .map(parseManualCode);
}

function normalizeManualNumsForRoom(nums, isAnnex) {
  let arr = (nums || []).map(Number).filter(n => !isNaN(n));

  if (!isAnnex) {
    arr = arr.filter(n => nearlyEqual(n, Math.round(n)));
  } else {
    arr = arr.filter(n =>
      nearlyEqual(n, Math.round(n)) ||
      nearlyEqual(n, 3.5) ||
      nearlyEqual(n, 8.5)
    );
  }

  return uniqueNumbers_(arr);
}

function uniqueNumbers_(nums) {
  const result = [];

  nums.forEach(n => {
    if (!result.some(x => nearlyEqual(x, n))) result.push(n);
  });

  return result;
}

function expandKoreanDayRanges(value) {
  const days = ['\uC6D4', '\uD654', '\uC218', '\uBAA9', '\uAE08', '\uD1A0', '\uC77C'];

  return cellText(value).replace(
    /([\uC6D4\uD654\uC218\uBAA9\uAE08\uD1A0\uC77C])\s*[~\uFF5E\uFF0D\u2013\u2014-]\s*([\uC6D4\uD654\uC218\uBAA9\uAE08\uD1A0\uC77C])/g,
    (match, start, end) => {
      const startIndex = days.indexOf(start);
      const endIndex = days.indexOf(end);

      if (startIndex === -1 || endIndex === -1 || startIndex > endIndex) return match;

      return days.slice(startIndex, endIndex + 1).join('');
    }
  );
}

function normalizeKoreanDaysToLetters(value) {
  const map = {
    '\uC6D4': 'A',
    '\uD654': 'B',
    '\uC218': 'C',
    '\uBAA9': 'D',
    '\uAE08': 'E',
    '\uD1A0': 'F',
    '\uC77C': 'G'
  };

  return cellText(value).replace(/[\uC6D4\uD654\uC218\uBAA9\uAE08\uD1A0\uC77C]/g, ch => map[ch] || ch);
}

function normalizeStudentId(value) {
  let s = cellText(value);
  if (!s) return '';

  s = s.replace(/,/g, '');
  s = s.replace(/\.0+$/, '');
  s = s.replace(/\s+/g, '');

  return s;
}

function normalizeSeat(value) {
  let s = cellText(value);
  if (!s) return '';

  s = s.replace(/\.0+$/, '');
  s = s.replace(/\s+/g, '');

  return s;
}

function normalizeRoomName(value) {
  let s = cellText(value);
  if (!s) return '';

  s = s.replace(/\s+/g, '');
  s = s.replace(/\(.*?\)/g, '');
  s = s.replace(/\uC2E4$/, '\uD638');

  if (/^2\uAD00\d{3}\uD638$/.test(s)) return s;
  if (/^\d{3}\uD638$/.test(s)) return s;

  const m1 = s.match(/^(2\uAD00\d{3})$/);
  if (m1) return m1[1] + '\uD638';

  const m2 = s.match(/^(\d{3})$/);
  if (m2) return m2[1] + '\uD638';

  const m3 = s.match(/^(2\uAD00\d{3})/);
  if (m3) return m3[1] + '\uD638';

  const m4 = s.match(/^(\d{3})/);
  if (m4) return m4[1] + '\uD638';

  return s;
}

function normalizePeriod(value) {
  if (value === null || value === undefined || value === '') return null;
  if (!isNaN(Number(value))) return Number(value);

  const m = cellText(value).match(/\d+(?:\.\d+)?/);
  return m ? Number(m[0]) : null;
}

function isAnnexRoom(name) {
  return /^2\uAD00/.test(cellText(name));
}

function nearlyEqual(a, b) {
  return Math.abs(Number(a) - Number(b)) < 1e-9;
}

function cellText(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}

function isProtectedAttendanceValue(value) {
  const text = cellText(value);
  if (!text) return false;

  if (PROTECTED_VALUE_SET.has(text)) return true;
  if (PROTECTED_PREFIXES.some(prefix => text.startsWith(prefix))) return true;

  return false;
}

function isAutoScheduleValue(value) {
  const text = cellText(value);
  if (!text) return false;
  if (text === '-') return true;
  if (isProtectedAttendanceValue(text)) return false;

  return true;
}

function getNext7DateKeysFromToday() {
  const todayKey = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd');
  const result = [];

  for (let i = 0; i <= 6; i++) {
    result.push(addDaysToDateKey_(todayKey, i));
  }

  return result;
}

function addDaysToDateKey_(dateKey, daysToAdd) {
  const parts = String(dateKey).split('-').map(Number);
  const dt = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2] + daysToAdd));

  return normalizeDateKey(
    dt.getUTCFullYear(),
    dt.getUTCMonth() + 1,
    dt.getUTCDate()
  );
}

function formatDateKey(dateObj) {
  return Utilities.formatDate(dateObj, TIMEZONE, 'yyyy-MM-dd');
}

function parseDateKeyFlexible_(value) {
  const keys = parseDateKeyListFlexible_(value);
  return keys.length ? keys[0] : null;
}

function parseDateKeyListFlexible_(value) {
  if (value === null || value === undefined || value === '') return [];

  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return [formatDateKey(value)];
  }

  let text = cellText(value);
  if (!text || text === '-') return [];

  text = text
    .replace(/[\uFF0C\u3001]/g, ',')
    .replace(/\s+/g, ' ')
    .replace(/(\d{1,2})\s*\uC6D4\s*(\d{1,2})\s*\uC77C/g, '$1/$2');

  const result = [];
  const currentYear = Number(Utilities.formatDate(new Date(), TIMEZONE, 'yyyy'));
  const rangeRegex = /(\d{4}[.\-/]\s*\d{1,2}[.\-/]\s*\d{1,2}|\d{1,2}[/.]\s*\d{1,2})\s*[~\uFF5E\uFF0D\u2013\u2014-]\s*(\d{4}[.\-/]\s*\d{1,2}[.\-/]\s*\d{1,2}|\d{1,2}[/.]\s*\d{1,2})/g;

  text = text.replace(rangeRegex, (match, startText, endText) => {
    const startKey = parseOneDateKeyText_(startText, currentYear);
    const endKey = parseOneDateKeyText_(endText, currentYear);

    if (startKey && endKey && startKey <= endKey) {
      let key = startKey;
      let guard = 0;

      while (key <= endKey && guard < 370) {
        result.push(key);
        key = addDaysToDateKey_(key, 1);
        guard++;
      }
    }

    return ' ';
  });

  let m;

  const fullRegex = /(\d{4})[.\-/]\s*(\d{1,2})[.\-/]\s*(\d{1,2})/g;
  while ((m = fullRegex.exec(text)) !== null) {
    result.push(normalizeDateKey(Number(m[1]), Number(m[2]), Number(m[3])));
  }

  const compactRegex = /\b(\d{4})(\d{2})(\d{2})\b/g;
  while ((m = compactRegex.exec(text)) !== null) {
    result.push(normalizeDateKey(Number(m[1]), Number(m[2]), Number(m[3])));
  }

  const mdRegex = /(?:^|[^0-9])(\d{1,2})[/.]\s*(\d{1,2})(?:[^0-9]|$)/g;
  while ((m = mdRegex.exec(text)) !== null) {
    result.push(normalizeDateKey(currentYear, Number(m[1]), Number(m[2])));
  }

  return uniqueStrings_(result);
}

function parseOneDateKeyText_(text, defaultYear) {
  const s = cellText(text);
  if (!s) return null;

  let m = s.match(/^(\d{4})[.\-/]\s*(\d{1,2})[.\-/]\s*(\d{1,2})$/);
  if (m) return normalizeDateKey(Number(m[1]), Number(m[2]), Number(m[3]));

  m = s.match(/^(\d{4})(\d{2})(\d{2})$/);
  if (m) return normalizeDateKey(Number(m[1]), Number(m[2]), Number(m[3]));

  m = s.match(/^(\d{1,2})[/.]\s*(\d{1,2})$/);
  if (m) return normalizeDateKey(defaultYear, Number(m[1]), Number(m[2]));

  return null;
}

function uniqueStrings_(values) {
  const seen = Object.create(null);
  const result = [];

  values.forEach(value => {
    const text = cellText(value);
    if (!text || seen[text]) return;

    seen[text] = true;
    result.push(text);
  });

  return result;
}

function extractDateKeyFromHeader(headerValue, targetDateSet) {
  if (Object.prototype.toString.call(headerValue) === '[object Date]' && !isNaN(headerValue)) {
    return formatDateKey(headerValue);
  }

  const text = cellText(headerValue);
  if (!text) return null;

  let m = text.match(/(\d{4})[.\-/]\s*(\d{1,2})[.\-/]\s*(\d{1,2})/);
  if (m) return normalizeDateKey(Number(m[1]), Number(m[2]), Number(m[3]));

  m = text.match(/(\d{1,2})[/.]\s*(\d{1,2})/);
  if (m) {
    const month = Number(m[1]);
    const day = Number(m[2]);
    const currentYear = Number(Utilities.formatDate(new Date(), TIMEZONE, 'yyyy'));
    const candidateYears = [currentYear - 1, currentYear, currentYear + 1];

    for (let i = 0; i < candidateYears.length; i++) {
      const key = normalizeDateKey(candidateYears[i], month, day);
      if (targetDateSet && targetDateSet.has(key)) return key;
    }

    return normalizeDateKey(currentYear, month, day);
  }

  return null;
}

function normalizeDateKey(year, month, day) {
  return [
    String(year).padStart(4, '0'),
    String(month).padStart(2, '0'),
    String(day).padStart(2, '0')
  ].join('-');
}

function getWeekdayIndexFromDateKey(dateKey) {
  const parts = String(dateKey).split('-').map(Number);
  const dt = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2]));
  const day = dt.getUTCDay();

  return day === 0 ? 6 : day - 1;
}

function getSpreadsheetId(input) {
  const text = cellText(input);
  const m = text.match(/\/d\/([a-zA-Z0-9-_]+)/);

  return m ? m[1] : text;
}

function getAttendanceSpreadsheetId_() {
  const id = getSpreadsheetId(ATTENDANCE_SYNC_CONFIG.attendanceSpreadsheetId);

  if (!id || id === 'YOUR_ATTENDANCE_SPREADSHEET_ID') {
    return '';
  }

  return id;
}

function getDayMap() {
  return DAY_KEY_TO_SHEET_COL;
}

function abbreviateSubject(value) {
  const s = cellText(value);
  const table = {
    '\uC0AC\uD68C\uBB38\uD654': '\uC0AC\uBB38',
    '\uD55C\uAD6D\uC9C0\uB9AC': '\uD55C\uC9C0',
    '\uC138\uACC4\uC9C0\uB9AC': '\uC138\uC9C0',
    '\uC0DD\uD65C\uACFC\uC724\uB9AC': '\uC0DD\uC724',
    '\uC724\uB9AC\uC640\uC0AC\uC0C1': '\uC724\uC0AC',
    '\uC815\uCE58\uC640\uBC95': '\uC815\uBC95',
    '\uB3D9\uC544\uC2DC\uC544\uC0AC': '\uB3D9\uC0AC',
    '\uC138\uACC4\uC0AC': '\uC138\uC0AC',
    '\uBB3C\uB9AC\uD559\u2160': '\uBB3C\uB9AC',
    '\uBB3C\uB9AC\uD559I': '\uBB3C\uB9AC',
    '\uBB3C\uB9AC\uD5591': '\uBB3C\uB9AC',
    '\uD654\uD559\u2160': '\uD654\uD559',
    '\uD654\uD559I': '\uD654\uD559',
    '\uD654\uD5591': '\uD654\uD559',
    '\uC0DD\uBA85\uACFC\uD559\u2160': '\uC0DD\uBA85',
    '\uC0DD\uBA85\uACFC\uD559I': '\uC0DD\uBA85',
    '\uC0DD\uBA85\uACFC\uD5591': '\uC0DD\uBA85',
    '\uC9C0\uAD6C\uACFC\uD559\u2160': '\uC9C0\uAD6C',
    '\uC9C0\uAD6C\uACFC\uD559I': '\uC9C0\uAD6C',
    '\uC9C0\uAD6C\uACFC\uD5591': '\uC9C0\uAD6C',
    '\uD1B5\uD569\uC0AC\uD68C': '\uD1B5\uC0AC',
    '\uD1B5\uD569\uACFC\uD559': '\uD1B5\uACFC',
    '\uC218\uB9AC\uB17C\uC220': '\uB17C\uC220'
  };

  return table[s] || s;
}
