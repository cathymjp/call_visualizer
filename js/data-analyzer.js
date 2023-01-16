/*
The following is an example of how a row in the Excel file will be represented in JSON format:
{
    "Date/Time": "2/28/17 23:38",
    "Direction": "IN",
    "Duration": "0",
    "Phone Number": "XXXX",
    "Extension": "0",
    "Call Line": "1",
    "Trunk Type": "Unknown",
    "Call Status": "RNA",
    "End Time": "2/28/17 23:38",
    "isTAProc": "0",
    "Call HoldDuration": "0",
    "Source Country Code": "1",
    "ACDExt": "0",
    "Transfer Ext": "0",
    "AgentName": "                                ",
    "Termtime": "                                ",
    "Party Duration": "0",
    "PartyIPPort": "0",
    "PartyType": "1",
    "PartyCodeType": "0",
    "WaitDuration": "0",
    "VMDuration": "0",
    "RecDuration": "0",
    "AADuration": "0",
    "HoldDuration": "0",
    "RingDuration": "0",
    "PartyWrapDuration": "0"
}
*/

/**
    Delete an entry's extra fields that are not relevant to the analysis.
**/
function deleteRedundantFields(jsonWorkbookEntry) {
    var entry = jsonWorkbookEntry;
    delete entry["Trunk Type"];
    delete entry.isTAProc;
    delete entry["Call HoldDuration"]
    delete entry["Source Country Code"]
    delete entry.ACDExt
    delete entry["Transfer Ext"]
    delete entry.AgentName
    delete entry.Termtime
    delete entry["Party Duration"]
    delete entry.PartyIPPort
    delete entry.PartyType
    delete entry.PartyCodeType
    delete entry.WaitDuration
    delete entry.VMDuration
    delete entry.RecDuration
    delete entry.AADuration
    delete entry.HoldDuration
    delete entry.RingDuration
    delete entry.PartyWrapDuration
}

/**
    Clean the JSON representation of the call data.
**/
function cleanJSONWorkbook(jsonWorkbookEntries) {
    var entry;
    for (var i = 0; i < jsonWorkbookEntries.length; i++) {
        entry = jsonWorkbookEntries[i];
        deleteRedundantFields(entry);
        cleanDates(entry);
        cleanDuration(entry);
    }
}

/**
    Format the date of each entry.
**/
function cleanDates(jsonWorkbookEntry) {
    jsonWorkbookEntry.Date = new Date(jsonWorkbookEntry["Date/Time"]);
    delete jsonWorkbookEntry["Date/Time"];
    jsonWorkbookEntry["End Time"] = new Date(jsonWorkbookEntry["End Time"]);
}

/**
    Format the duration of each entry.
**/
function cleanDuration(jsonWorkbookEntry) {
    var parsedDuration = parseFloat(jsonWorkbookEntry.Duration);
    if (parsedDuration != NaN) {
        jsonWorkbookEntry.Duration = parsedDuration;
    } else {
        delete jsonWorkbookEntry.Duration;
    }
}

/**
    Label each entry as missed or not missed by creating/modifying a boolean attribute "Missed", based on the
    missedDurationRule parameter. The parameter "inclusive" specifies whether to label an entry as missed when its
    duration == missedDurationRule. If inclusive is true, then the entry is labeled as missed.
**/
function labelMissed(jsonWorkbookArray, missedDurationRule, inclusive) {
    var entry;
    for (var i = 0; i < jsonWorkbookArray.length; i++) {
        entry = jsonWorkbookArray[i];
        if (entry.Direction == "IN" && entry.Duration > 1) {
            var duration = entry.Duration;
            if (duration != undefined) { // Duration exists
                if (duration < missedDurationRule) {
                    entry.Missed = true;
                } else if (duration > missedDurationRule) {
                    entry.Missed = false;
                } else { // duration == missedDurationRule
                    if (inclusive) {
                        entry.Missed = true;
                    } else {
                        entry.Missed = false;
                    }
                }
            }
        } else {
          jsonWorkbookArray.splice(i, 1);
        }
    }
}



/**
    Convert the given month number to its corresponding month string. Input ranges
    between 0 and 11, inclusive, where 0 maps to January and 11 maps to December.
**/
function convertMonth(monthNum) {
    switch (monthNum) {
        case 0: return "January"; break;
        case 1: return "February"; break;
        case 2: return "March"; break;
        case 3: return "April"; break;
        case 4: return "May"; break;
        case 5: return "June"; break;
        case 6: return "July"; break;
        case 7: return "August"; break;
        case 8: return "September"; break;
        case 9: return "October"; break;
        case 10: return "November"; break;
        case 11: return "December"; break;
        default: return null;
    }
}

/**
    Get total number of made/missed calls per year. Return an array of objects in the
    following format:

    [
        {
            "numCallsTotal": 5555,
            "numMissedCalls": 200,
            "numMadeCalls": 5355,
            "expanded": true
        },

        {
            "numCallsTotal": 5112,
            "numMissedCalls": 112,
            "numMadeCalls": 5000,
            "expanded": false
        },

        ...
    ]

    This array will have one entry per year of data.
**/
function getCallDataPerYear(jsonWorkbookEntries, years) {
    // the array of call data per year
    var callDataPerYear = [];

    // make one entry of call data per year
    for (var i = 0; i < years.length; i++) {
        var obj = new Object();
        obj.numCallsTotal = 0;
        obj.numMissedCalls = 0;
        obj.numMadeCalls = 0;
        obj.expanded = false;
        callDataPerYear.push(obj);
    }

    // aggregate the call data for each year
    for (var i = 0; i < years.length; i++) {
        for (var j = 0; j < jsonWorkbookEntries.length; j++) {
            var entry = jsonWorkbookEntries[j];
            var entryYear = entry.Date.getFullYear();
            var entryMonth = entry.Date.getMonth();
            var entryTime = entry.Date.getTime();
            // only add data corresponding to the proper year, and only if it fits the user options
            if (entryYear === years[i] && checkedMonths[entryMonth] && entryTime >= userOptions.startDate.getTime()
                && entryTime <= userOptions.endDate.getTime()) {
                callDataPerYear[i].numCallsTotal++;        // increment the total number of calls by one
            if (entry.Missed) {
                    callDataPerYear[i].numMissedCalls++;   // if the call was missed, increment total number of missed calls
                } else {
                    callDataPerYear[i].numMadeCalls++;     // else increment total number of made calls
                }
            }
        }
    }

    return callDataPerYear;
}

/**
    Get total number of made/missed calls per month. Return an array of objects,
    indexed by month, in the following format:

    [
        {
            "numCallsTotal": 555,
            "numMissedCalls": 100,
            "numMadeCalls": 455,
            "expanded": false
        },

        {
            "numCallsTotal": 0,
            "numMissedCalls": 0,
            "numMadeCalls": 0
            "expanded": true
        },

        ...
    ]

    This array will be of size 12, with one entry per month.
    Note: It might make sense to edit this function to specify which
    months should be analyzed based on the user's input.
**/
function getCallDataPerMonth(jsonWorkbookEntries, year) {
    // array of objects
    var callDataPerMonth = [];

    // initialize each month object
    for (var i = 0; i < 12; i++) {
        var obj = new Object();
        obj.numCallsTotal = 0;
        obj.numMissedCalls = 0;
        obj.numMadeCalls = 0;
        obj.expanded = false;
        callDataPerMonth.push(obj);
    }

    // aggregate the call data over each month of the specified year
    for (var i = 0; i < jsonWorkbookEntries.length; i++) {
        var entry = jsonWorkbookEntries[i];
        var entryMonth = entry.Date.getMonth();
        var entryYear = entry.Date.getFullYear();
        var entryTime = entry.Date.getTime();
        // only add data for the specified year, and only if it fits the user options
        if (entryYear == year && checkedMonths[entryMonth] && entryTime >= userOptions.startDate.getTime()
            && entryTime <= userOptions.endDate.getTime()) {
            callDataPerMonth[entryMonth].numCallsTotal++;        // increment the total number of calls by one
        if (entry.Missed) {
                callDataPerMonth[entryMonth].numMissedCalls++;   // if the call was missed, increment total number of missed calls
            } else {
                callDataPerMonth[entryMonth].numMadeCalls++;     // else increment total number of made calls
            }
        }
    }

    return callDataPerMonth;
}

/**
    Get total number of made/missed calls per day for a given month of a given year. Return array
    of objects, indexed by day [0, 30], in the following format:

    [
        {
            "numCallsTotal": 100,
            "numMissedCalls": 10,
            "numMadeCalls": 90,
            "expanded": false
        },

        {
            "numCallsTotal": 0,
            "numMissedCalls": 0,
            "numMadeCalls": 0,
            "expanded": true
        },

        ...
    ]

    This array will be of size 31, with one entry per day of the month (assumes 31 days each month).
**/
function getCallDataPerDay(jsonWorkbookEntries, month, year) {
    // array of objects
    var callDataPerDay = [];

    // create initial objects (assumes 31 days for each month)
    for (var i = 0; i < 31; i++) {
        var obj = new Object();
        obj.numCallsTotal = 0;
        obj.numMissedCalls = 0;
        obj.numMadeCalls = 0;
        obj.expanded = false;
        callDataPerDay.push(obj);
    }

    // aggregate call data
    for (var i = 0; i < jsonWorkbookEntries.length; i++) {
        var entry = jsonWorkbookEntries[i];
        var entryMonth = entry.Date.getMonth();
        var entryYear = entry.Date.getFullYear();
        var entryTime = entry.Date.getTime();
        // only add data for the given month of the given year, and only if it fits the user options
        if (entryMonth === month && entryYear === year && checkedMonths[entryMonth] && entryTime >= userOptions.startDate.getTime()
            && entryTime <= userOptions.endDate.getTime()) {
            var entryDate = entry.Date.getDate();
        var callObj = callDataPerDay[entryDate - 1];
            callObj.numCallsTotal++;        // increment total number of calls on this day by one
            if (entry.Missed) {
                callObj.numMissedCalls++;   // if call was missed, increment total number of missed calls
            } else {
                callObj.numMadeCalls++;     // else increment total number of made calls
            }
        }
    }

    return callDataPerDay;
}

/**
    Get total number of made/missed calls per day for a given day in a given month
    of a given year. Return array of objects, indexed by hour [0, 23], in the following format:

    [
        {
            "numCallsTotal": 7,
            "numMissedCalls": 1,
            "numMadeCalls": 6
        },

        {
            "numCallsTotal": 0,
            "numMissedCalls": 0,
            "numMadeCalls": 0
        },

        ...
    ]

    This array will be of size 24, with one entry per hour of the day.
**/
function getCallDataPerHour(jsonWorkbookEntries, day, month, year) {
    // array of objects
    var callDataPerHour = [];

    // create initial objects
    for (var j = 0; j < 24; j++) {
        var obj = new Object();
        obj.numCallsTotal = 0;
        obj.numMissedCalls = 0;
        obj.numMadeCalls = 0;
        callDataPerHour.push(obj);
    }

    // aggregate call data
    for (var j = 0; j < jsonWorkbookEntries.length; j++) {
        var entry = jsonWorkbookEntries[j];
        var entryDate = entry.Date.getDate();
        var entryMonth = entry.Date.getMonth();
        var entryYear = entry.Date.getFullYear();
        var entryTime = entry.Date.getTime();
        // only add data for the given day of the given month of the given year, and only
        // if it fits the user options
        if ((entryDate - 1) === day && entryMonth === month && entryYear === year && checkedMonths[entryMonth]
            && entryTime >= userOptions.startDate.getTime() && entryTime <= userOptions.endDate.getTime()) {
            var entryHour = entry.Date.getHours();
        var callObj = callDataPerHour[entryHour];
            callObj.numCallsTotal++;        // increment total number of calls on this day by one
            if (entry.Missed) {
                callObj.numMissedCalls++;   // if call was missed, increment total number of missed calls
            } else {
                callObj.numMadeCalls++;     // else increment total number of made calls
            }
        }
    }

    return callDataPerHour;
}

/**
    Collect the months that the data spans.
**/
function getActiveMonths(jsonWorkbookEntries) {
    var monthArray = [];
    var entry;

    for (var i = 0; i < 12; i++) { // Initialize all months to inactive
        monthArray[i] = false;
    }

    for (var i = 0; i < jsonWorkbookEntries.length; i++) {
        entry = jsonWorkbookEntries[i];
        monthArray[entry.Date.getMonth()] = true
    }

    return monthArray;
}

/**
    Collect the years that the data spans.
**/
function getActiveYears(jsonWorkbookEntries) {
    var uniqueDict = {};
    var uniqueYears = [];
    for (var i = 0; i < jsonWorkbookEntries.length; i++) {
        var year = jsonWorkbookEntries[i].Date.getFullYear();
        if (typeof(uniqueDict[year]) == "undefined") {
            uniqueYears.push(year);
            uniqueDict[year] = 0;
        }
    }
    return uniqueYears;
}

/**
    Sort the data by date in ascending order.
**/
function sortByDate(jsonWorkbookEntries) {
    jsonWorkbookEntries.sort(function(a, b) {
        return a.Date - b.Date;
    });
    return jsonWorkbookEntries;
}
