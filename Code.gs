var YEAR = 2020;
var FIRST_PTO_ROW = 6;

// ----- SHEETS TOOLBAR SETUP -----
function onOpen() {
    var menuItems = [{
        name: 'Calculate!',
        functionName: 'main'
    }, ];
    var spreadsheet = SpreadsheetApp.getActive();
    spreadsheet.addMenu('PTO Scheduler', menuItems);
}

// ----- CLASSES -----

var TimeEntry = function(sheet, rowNum) {
    this.rowNum = rowNum;
    this.date = new Date(sheet.getRange('A' + this.rowNum).getValue());

    this.getMinutesUsed = function(code) {
        var column = "";
        switch (code) {
            case "V":
                column = "C";
                break;
            case "H":
                column = "D";
                break;
            case "P":
                column = "E";
                break;
            default:
                throw new Error("Invalid TimeType code \"" + code + "\"");
        }

        var cell = sheet.getRange(column + this.rowNum);
        var time = cell.getDisplayValue();
        var mins = 0;
        if (time == "X" || time == "x") {
            mins = 7 * 60;
        } else if (!cell.isBlank()) {
            mins = getDurationAsMinutes(time);
        } 
        return mins;
    }
};

var TimeType = function(code, accrualPerPayMins, maxCarryMins, carryStart, carryEnd) {
    this.code = code;
    this.accrualPerPayMins = accrualPerPayMins;
    this.maxCarryMins = maxCarryMins;
    this.carryStart = carryStart;
    this.carryEnd = carryEnd;
}

// ----- HELPERS -----

function logObject(obj) {
    Logger.log(JSON.stringify(obj, null, 2));
}

Date.prototype.getWeek = function() {
    var firstDayOfYear = new Date("1/1/" + YEAR);
    return Math.ceil((((this - firstDayOfYear) / 86400000) + firstDayOfYear.getDay() + 1) / 7);
}

function getDurationAsMinutes(d) {
    var error = new Error("Please check the format of value \"" + d + "\". Must be in the format HH:MM");
    try {
        var parts = d.split(":");
        var hrs = Number(parts[0]);
        var mins = Number(parts[1]);
        if (isNaN(hrs) || isNaN(mins)) {
            throw error;
        }
        return (hrs * 60) + mins;
    } catch (e) {
        throw error;
    }
}

function addDays(date, days) {
    var timeToAdd = days * (24 * 3600 * 1000);
    var newDate = new Date(date.getTime() + timeToAdd);
    return newDate;
}

function writeDurationToCell(sheet, cell, durationMins) {
    var hours = Math.floor(durationMins / 60);
    var mins = durationMins % 60;
    var outputHours = Utilities.formatString("%02d", hours);
    var outputMins = Utilities.formatString("%02d", mins);
    var output = outputHours + ":" + outputMins;

    var hrFormat = "[h]";
    if (hours >= 10 && hours <= 100) {
        hrFormat = "[hh]";
    } else if (hours >= 100) {
        hrFormat = "[hhh]";
    }

    var destinationCell = sheet.getRange(cell);
    destinationCell.setValue(output).setNumberFormat(hrFormat + ':mm');
}

// ----- PTO FUNCTIONS -----

/* Gets all PTO entries from sheet */
function getSortedEntries(sheet) {
    // find the index of the last PTO record
    var lastRecord = 0;
    for (var i = FIRST_PTO_ROW; i < sheet.getMaxRows(); i++) {
        var cellValue = sheet.getRange('A' + i).getValue();
        if (Object.prototype.toString.call(cellValue) !== "[object Date]") {
            lastRecord = i - 1;
            break;
        }
    }

    if (lastRecord == FIRST_PTO_ROW-1) {
        throw new Error("Please enter at least one PTO date"); 
    }

    // populate
    var arr = new Array(lastRecord - FIRST_PTO_ROW);
    for (var row = FIRST_PTO_ROW; row <= lastRecord; row++) {
        var entry = new TimeEntry(sheet, row)
        arr[row - FIRST_PTO_ROW] = entry;
    }

    // sort
    arr.sort(function(x, y) {
        var xDate = x.date;
        var yDate = y.date;
        return xDate == yDate ? 0 : xDate < yDate ? -1 : 1;
    });

    return arr;
}

function getStartBalanceMinutes(sheet, code) {
    var rowNum = null;
    switch (code) {
        case "V":
            rowNum = 1;
            break;
        case "H":
            rowNum = 2;
            break;
        case "P":
            rowNum = 3;
            break;
        default:
            throw new Error("Invalid TimeType code \"" + code + "\"");
    }
    var duration = sheet.getRange("D" + rowNum).getDisplayValue();
    var mins = getDurationAsMinutes(duration);
    return mins;
}

function getAccruedTimeMinutes(timeType, date, startDate) {
    var weekNum = date.getWeek();

    // offset week according to start date
    if (startDate != null && date > startDate) {
        weekNum -= startDate.getWeek();
    }

    // set weekNum to the closest even number, rounding down
    if (weekNum % 2 != 0) {
        weekNum--;
    }

    var accrued = timeType.accrualPerPayMins * (weekNum / 2);
    return accrued;
}

/* Returns YTD minutes used for specified TimeType
   (optional param "startDate" sets lower limit for date range) */
function getUsedTimeMinutes(timeType, sortedEntries, date, startDate) {
    if (date.getFullYear() != YEAR) {
        var dateStr = Utilities.formatDate(date, "GMT+1", "M/d/yyyy")
        throw new Error("Date is not in current year \"" + dateStr + "\"");
    }

    var timeUsedMins = 0;
    for (var i = 0; i < sortedEntries.length; i++) {
        var entry = sortedEntries[i];
        if (entry.date <= date) {
            // skip entry if before startDate
            if (startDate != null && entry.date < startDate) {
                continue;
            }
            timeUsedMins += entry.getMinutesUsed(timeType.code);
        } else if (entry.date > date) {
            // expecting a sorted list, so end the loop if entry is after the specified date
            break;
        }
    }
    return timeUsedMins;
}

function getRemainingTimeAsMinutes(timeType, entries, startBalanceMins, date) {
    var rem = 0;

    if (date < timeType.carryStart) {
        /* --before carryover period-- */
        var accrued = getAccruedTimeMinutes(timeType, date);
        var used = getUsedTimeMinutes(timeType, entries, date);
        rem = startBalanceMins + accrued - used;

    } else if (date >= timeType.carryStart && date <= timeType.carryEnd) {
        /* --during carryover period--*/

        // calculate time before carryover start
        var carryover = startBalanceMins;
        var firstDayOfYear = new Date("1/1/" + YEAR);
        if (timeType.carryStart.getTime() !== firstDayOfYear.getTime()) {
            var dayBeforeCarryStart = addDays(timeType.carryStart, -1);
            var accruedBeforeCarryStart = getAccruedTimeMinutes(timeType, dayBeforeCarryStart);
            var usedBeforeCarryStart = getUsedTimeMinutes(timeType, entries, dayBeforeCarryStart);
            carryover = Math.min(startBalanceMins + accruedBeforeCarryStart - usedBeforeCarryStart, timeType.maxCarryMins);
        }

        // calculate time after carryover start
        var accruedSinceCarryStart = getAccruedTimeMinutes(timeType, date, timeType.carryStart);
        var usedSinceCarryStart = getUsedTimeMinutes(timeType, entries, date, timeType.carryStart);

        // get remaining time
        rem = carryover + accruedSinceCarryStart - usedSinceCarryStart;

    } else if (date > timeType.carryEnd) {
        /* --after carryover period ends-- */
        var accruedSinceCarryStart = getAccruedTimeMinutes(timeType, date, timeType.carryStart);
        var usedSinceCarryStart = getUsedTimeMinutes(timeType, entries, date, timeType.carryStart);
        rem = accruedSinceCarryStart - usedSinceCarryStart;
    }

    return rem;
}

// ----- MAIN -----

function main() {
    // init spreadsheet 
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[0];
    // get array of TimeEntry objects 
    var entries = getSortedEntries(sheet);

    // init pto types
    var timeTypeVacation = new TimeType("V", 323, 4200, new Date("1/1/" + YEAR), new Date("6/30/" + YEAR));
    var timeTypeHealth = new TimeType("H", 161, 8400, new Date("7/1/" + YEAR), null);
    var timeTypePersonal = new TimeType("P", 48, 1260, new Date("7/1/" + YEAR), new Date("9/30/" + YEAR));

    // get starting balances
    var startBalanceVacation = getStartBalanceMinutes(sheet, timeTypeVacation.code);
    var startBalanceHealth = getStartBalanceMinutes(sheet, timeTypeHealth.code);
    var startBalancePersonal = getStartBalanceMinutes(sheet, timeTypePersonal.code);

    // loop through entries
    for (var i = 0; i < entries.length; i++) {
        // get remaining time for each category
        var entry = entries[i];
        var remVacation = getRemainingTimeAsMinutes(timeTypeVacation, entries, startBalanceVacation, entry.date);
        var remHealth = getRemainingTimeAsMinutes(timeTypeHealth, entries, startBalanceHealth, entry.date);
        var remPersonal = getRemainingTimeAsMinutes(timeTypePersonal, entries, startBalancePersonal, entry.date);

        // output remaining time values
        writeDurationToCell(sheet, ("F" + entry.rowNum), remVacation);
        writeDurationToCell(sheet, ("G" + entry.rowNum), remHealth);
        writeDurationToCell(sheet, ("H" + entry.rowNum), remPersonal);
    }

    // clean up extra cells in "remaining" columns
    var firstEmptyRow = entries.length + FIRST_PTO_ROW;
    if (sheet.getLastRow() >= firstEmptyRow) {
        var rangeA1Notation = "F" + firstEmptyRow + ":H" + sheet.getLastRow();
        sheet.getRange(rangeA1Notation).clearContent();
    }
}