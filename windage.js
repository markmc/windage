/*

The basic idea of spreadsheet is as follows:

 - A "Boats" sheet containing the sail number, name, owner and handicaps
   of each boat.

   We have the following handicaps for each boat: IRC, whitesail IRC and
   2011 end of season ECHO.

   As the season continues we will adjust the ECHO handicap of each boat.
   A new ECHo handicap column will be added each time we make an
   adjustment.

 - A "Races" sheet with a row for each race containing the sheet name for
   that race, the OOD sail number, start time, date and the echo handicap
   applicable to that race.

   The sheet also contains automatically calculated columns for the OOD
   boat name and and day of the week of the race. This is just to help
   double check that the OOD sail number and the date values are correct.

 - A series sheet for each race series. We will have series 1/2/3, a
   saturday series, an overall series and a whitesail series. For each
   of these, we will have an IRC and ECHO series.

 - A sheet for each race.

   The only results that need to be entered into each race sheet is the
   sail number of each starter and their finishing time. Enter DNF if the
   boat did not finish. There is no need to enter boats who did not start
   the race.

The series and race sheets have a MYC menu with a "Calculate" menu item.
When you enter a race result, you need to hit the calculate button for
the race itself and each series which the race is part of.
*/

function calculateResults() {
    var sheet = SpreadsheetApp.getActiveSheet();

    // Race sheets are named SXRY where X = series, Y = race
    // WS race sheets are named WSY where Y = race
    if (sheet.getName().search(/S[0-9]+R[0-9]+/) != -1 ||
        sheet.getName().search(/WS[0-9]+/) != -1) {
        return calculateRaceResults(sheet);
    }
    // IRC series sheets are named SXIRC where X = series
    // WS IRC series sheet is WSIRC
    else if (sheet.getName().search(/S[0-9]IRC/) != -1 ||
             sheet.getName() == "WSIRC") {
        return calculateSeriesResults(sheet, IRC_PLACE_RANGE);
    }
    // ECHO series sheets are named SXECHO where X = series
    // WS ECHO series sheet is WSECHO
    else if (sheet.getName().search(/S[0-9]ECHO/) != -1 ||
             sheet.getName() == "WSECHO") {
        return calculateSeriesResults(sheet, ECHO_PLACE_RANGE);
    }
}

// The name of the Boats sheet
var BOATS_SHEET = 'Boats';

// The max number of columns in the Boats
// sheet which contain data; will increase
// as new echo handicaps are added
var N_BOAT_COLUMNS = 10;

// The max number of boats in the Boats sheet
var N_BOAT_ROWS = 30;

// DNF is scored as number of starters plus one
var DNF_PENALTY = 1;
// DNC is scored as number of starters plus three
var DNC_PENALTY = 3;

// Range in each race sheet for ECHO placing
var ECHO_PLACE_RANGE = "O3:O23";
// Range in each race sheet for IRC placing
var IRC_PLACE_RANGE = "J3:J23";

// Bizarre stuff. See:
// http://stackoverflow.com/questions/10363169/reading-and-writing-time-values-from-to-a-spreadsheet-using-gas
var MINUTES_DELTA = 25;
var SECONDS_DELTA = 21;

// Construct an object representing a boat from a
// row in the Boats sheet.
// The headers of the columns in the sheet are used
// as the names of the properties on the object.
function boat_(headers, row) {
    for (var h = 0; h < headers.length; h++) {
        if (headers[h]) {
            this[headers[h].toLowerCase()] = row[h];
        }
    }
}

// Return an array of boat objects, one for each row
// in the Boats sheet
function getBoats_() {
    var boatsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BOATS_SHEET);
    var boatsValues = boatsSheet.getRange(1, 1, N_BOAT_ROWS+1, N_BOAT_COLUMNS).getValues();
    var boats = [];
    for (var i = 1; i < boatsValues.length; i++) {
        boats.push(new boat_(boatsValues[0], boatsValues[i]));
    }
    return boats;
}

// Lookup a boat object from an array of boat objects
// using the boats sail number.
function lookupBoat_(boats, sail) {
    if (!sail) {
        return null;
    }
    for (var i = 0; i < boats.length; i++) {
        if (boats[i].sail == sail) {
            return boats[i];
        }
    }
    return null;
}

//
// Calculate the contents of a series sheet
//
// The parameters are the sheet object and the range to
// use to extract the placings from each of the race
// sheets depending on whether its IRC or ECHO
//
// The stages are as follows:
//
// 1 - Look up the boat names and owners from the
//     Boats sheet using the sail number
//
// 2 - Look up the race details for each race from
//     the Races sheet. We need the OOD (to award
//     average points) and which echo handicap to
//     use
//
// 3 - Get the sheet object for each race using the
//     name of the race in each of the columns in
//     the series sheet
//
// 4 - For each race, get the placing of each boat
//     in the race results
//
// 5 - For each boat, calculate its points for each
//     race and also its total points for that series
//     taking into account DNF, DNC and average
//     points for OOD.
//
// 6 - Write out all those results to the series table
//
// 7 - Generate hyperlinks to each of the individual
//     races
//
function calculateSeriesResults(sheet, placeRange) {
    var boats = getBoats_();

    var sailNos = sheet.getRange("A2:A22").getValues();
    var races = sheet.getRange("D1:N1").getValues();

    // Stage 1 - boat names and owners
    var boatNamesAndOwners = [];

    for (var i = 0; i < sailNos.length; i++) {
        var boat = lookupBoat_(boats, sailNos[i][0]);

        if (boat) {
            boatNamesAndOwners.push([boat.name, boat.owner]);
        } else {
            boatNamesAndOwners.push(["", ""]);
        }
    }

    sheet.getRange("B2:C22").setValues(boatNamesAndOwners);

    var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();

    // Stage 2 - race details
    var raceData = spreadSheet.getSheetByName("Races").getRange("A2:D40").getValues();

    // Stage 3 - sheet objects
    var sheets = [];
    var oods = [];
    for (var i = 0; i < races[0].length; i++) {
        if (races[0][i] && races[0][i] != "Total") {
            sheets.push(spreadSheet.getSheetByName(races[0][i]));

            for (var j = 0; j <= raceData.length; j++) {
                if (raceData[j][0] && raceData[j][0] == races[0][i]) {
                    oods.push(raceData[j][1]);
                    break;
                }
            }
        }
    }

    // Stage 4 - get the placings
    var places = [];
    for (var i = 0; i < sheets.length; i++) {
        var sailValues = sheets[i].getRange("A3:A23").getValues();
        var placeValues = sheets[i].getRange(placeRange).getValues();
        var racePlaces = [];
        for (var j = 0; j < sailValues.length; j++) {
            if (sailValues[j][0]) {
                racePlaces.push([sailValues[j][0], placeValues[j][0]]);
            }
        }
        places.push(racePlaces);

    }

    // Initialize the contents of the series table
    var results = [];
    var fontlines = []
    for (var i = 0; i < sailNos.length; i++) {
        results.push([]);
        fontlines.push([]);
        for (var j = 0; j < races[0].length; j++) {
            results[i].push("");
            fontlines[i].push(null);
        }
    }

    // Stage 5 - calculate the points
    for (var i = 0; i < sailNos.length; i++) {
        var boat = lookupBoat_(boats, sailNos[i][0]);
        if (!boat) {
            continue;
        }
        var points = [];
        var numAvgs = 0;
        for (var j = 0; j < sheets.length; j++) {
            var racePlaces = places[j];
            if (racePlaces.length) {
                var ood = oods[j];
                if (ood == boat.sail) {
                    results[i][j] = "AVG";
                } else {
                    for (var k = 0; k < racePlaces.length; k++) {
                        if (racePlaces[k][0] == boat.sail) {
                            results[i][j] = racePlaces[k][1];
                            break;
                        }
                    }
                }
                if (!results[i][j]) {
                    results[i][j] = "DNC";
                }
                if (results[i][j] == "DNF" ||
                    results[i][j] == "DSQ" ||
                    results[i][j] == "BFD") {
                    dnf_points = racePlaces.length + DNF_PENALTY
                    points.push([j, dnf_points]);
                    results[i][j] = results[i][j] + "(" + dnf_points + ")";
                } else if (results[i][j] == "DNC") {
                    dnc_points = racePlaces.length + DNC_PENALTY;
                    points.push([j, dnc_points]);
                    results[i][j] = "DNC(" + dnc_points + ")";
                } else if (results[i][j] == "AVG") {
                    points.push([j, 0]);
                    numAvgs++;
                } else {
                    points.push([j, results[i][j]]);
                }
            }
        }

        num_discards = numDiscards_(points.length)

        points.sort(compare_points_);
        discarded_points = points.slice(points.length - num_discards, points.length)
        points = points.slice(0, points.length - num_discards);

        var total = 0;
        for (var k = 0; k < points.length; k++) {
            total += points[k][1];
        }

        average_points = 0
        if (points.length) {
            average_points = total / (points.length - numAvgs);
        }
        total += average_points * numAvgs;

        if (average_points) {
            for (var j = 0; j < sheets.length; j++) {
                if (results[i][j] == "AVG") {
                    results[i][j] = "AVG(" + average_points.toFixed(2) + ")";
                }
                for (var k = 0; k < discarded_points.length; k++) {
                    if (discarded_points[k][0] == j) {
                        fontlines[i][j] = "line-through";
                    }
                }
            }
        }

        results[i][j] = total.toFixed(2);
    }

    // Stage 6 - write out the series table
    sheet.getRange(2, 4, sailNos.length, races[0].length).setValues(results);
    sheet.getRange(2, 4, sailNos.length, races[0].length).setFontLines(fontlines);

    // Stage 7 - generate hyperlinks
    var spreadSheetUrl = spreadSheet.getUrl().replace("/ccc?", "/pub?");

    var headers = [];
    for (var i = 0; i < races[0].length; i++) {
        if (races[0][i] != "Total") {
            var sheetId = spreadSheet.getSheetByName(races[0][i]).getSheetId();
            var url = spreadSheetUrl + "&gid=" + sheetId + "&single=true";
            headers.push("=hyperlink(\"" + url + "\", \"" + races[0][i] + "\")");
        } else {
            headers.push(races[0][i]);
        }
    }
    sheet.getRange(1, 4, 1, races[0].length).setValues([headers]);
}

//
// Calculate the number of discards to be applied
// for a given number of races.
//
// A single discard after 3 races, 2 discards after
// 5 races and an extra discard for every 3 races
// after that
//
function numDiscards_(numRaces) {
    var discards = 0;
    if ((numRaces -= 3) >= 0) {
        discards++;
    }
    if ((numRaces -= 2) >= 0) {
        discards++;
    }
    while ((numRaces -= 3) >= 0) {
        discards++;
    }
    return discards;
}

//
// Calculate the contents of a race sheet
//
// The stages are as follows:
//
// 1 - Find the start time and name of the set of ECHO handicaps
//     to be used for this race
//
// 2 - Get the sail numbers and finishing time of each entrant
//
// 3 - Look up the name and owner of each entrant
//
// 4 - Calculate the elapsed time of each entrant and apply their
//     IRC and ECHO handicaps to give an addjusted time for each
//     handicap system
//
// 5 - Calculate the placing and "to win" time of each entrant
//
// 6 - Insert all the calculated values into the sheet
//
function calculateRaceResults(sheet) {
    Logger.log("Calculating race results for " + sheet.getName());

    var boats = getBoats_();

    var spreadSheet = SpreadsheetApp.getActiveSpreadsheet();

    // 1 - Get the start time and ECHO handicap set for this race
    var raceData = spreadSheet.getSheetByName("Races").getRange("A2:E40").getValues();

    var startTime;
    var echoName;
    for (var i = 0; i <= raceData.length; i++) {
        if (raceData[i][0] && raceData[i][0] == sheet.getName()) {
            startTime = raceData[i][2];
            echoName = raceData[i][4].toLowerCase();;
            break;
        }
    }

    // 2 - Get the sail numbers and finishing time of entrants
    var sailNos = sheet.getRange("A3:A23").getValues();
    var finishTimes = sheet.getRange("D3:D23").getValues();

    var boatNamesAndOwners = [];
    var elapsedTimes = [];
    var ircTimes = [];
    var echoTimes = [];

    for (var i = 0; i < sailNos.length; i++) {
        var boat = lookupBoat_(boats, sailNos[i][0]);

        // 3 - Look up the boat names and owners
        if (boat) {
            boatNamesAndOwners.push([boat.name, boat.owner]);
        } else {
            boatNamesAndOwners.push(["", ""]);
        }

        // 4 - Calculate the elapsed time of each entrant and
        //     apply their IRC and ECHO handicaps to give an
        //     addjusted time for each handicap system
        var eTime = elapsedTime_(startTime, finishTimes[i][0]);
        if (eTime) {
            elapsedTimes.push([eTime]);
            ircTimes.push(correctedTime_(eTime, boat.irc));
            echoTimes.push(correctedTime_(eTime, boat[echoName]));
        } else {
            elapsedTimes.push([""]);
            ircTimes.push("");
            echoTimes.push("");
        }
    }

    // 5 - Calculate the placing and "to win" time of each entrant
    var ircResults = [];
    var echoResults = [];
    for (var i = 0; i < sailNos.length; i++) {
        var boat = lookupBoat_(boats, sailNos[i][0]);
        if (!boat) {
            ircResults.push(["", "", "", ""]);
            echoResults.push(["", "", "", ""]);
            continue;
        }
        ircResults.push([boat.irc,
                         ircTimes[i],
                         toWin_(elapsedTimes[i][0], boat.irc, ircTimes),
                         rank_(ircTimes[i], ircTimes)]);
        echoResults.push([boat[echoName],
                          echoTimes[i],
                          toWin_(elapsedTimes[i][0], boat[echoName], echoTimes),
                          rank_(echoTimes[i], echoTimes)]);
    }

    // 6 - Insert all the calculated values into the sheet
    sheet.getRange("B3:C23").setValues(boatNamesAndOwners);
    sheet.getRange("E3:E23").setValues(elapsedTimes);
    sheet.getRange("G3:J23").setValues(ircResults);
    sheet.getRange("L3:O23").setValues(echoResults);
}

// Calculate the elapsed time between a given
// start and finish time
//
// We convert both times to seconds, calculate
// the differences between the two and convert
// the resulting number of seconds back into
// a time object and return it
//
// A special case is where the finish time has
// been recorded as "DNF" and we just return
// "DNF
//
function elapsedTime_(start, finish) {
    if (finish == "DNF" || finish == "DSQ" || finish == "BFD" || finish == "AVG") {
        return finish;
    }
    var startInSeconds = dateToSeconds_(start);
    var finishInSeconds = dateToSeconds_(finish);
    return finishInSeconds > 0 ? newTimeInSeconds_(finishInSeconds - startInSeconds) : "";
}

// Adjust a given elapsed time based on a supplied
// handicap
//
// We convert the elapsed time into seconds, multiply
// it by the handicap and convert the resulting number
// of seconds into a time object
function correctedTime_(elapsed, handicap) {
    if (elapsed == "DNF" || elapsed == "DSQ" || elapsed == "BFD" || elapsed == "AVG") {
        return elapsed;
    }
    var ret = dateToSeconds_(elapsed) * handicap;
    return ret > 0 ? newTimeInSeconds_(ret) : "";
}

// Find the winning time in seconds
// Divide that by the handicap to give the elapsed time
//   this boat would need to win
// Subtract that time from its actual elapsed time to
//   give the number of seconds faster it would need to win
function toWin_(elapsed, handicap, correctedTimes) {
    if (elapsed == "DNF" || elapsed == "DSQ" || elapsed == "BFD" || elapsed == "AVG") {
        return elapsed;
    }
    var winningTime = Number.MAX_VALUE;
    for (var i = 0; i < correctedTimes.length; i++) {
        var correctedTime = dateToSeconds_(correctedTimes[i]);
        if (correctedTime > 0) {
            winningTime = Math.min(winningTime, correctedTime);
        }
    }
    var toWinSeconds = dateToSeconds_(elapsed) - (winningTime/handicap);
    return toWinSeconds > 0 ? newTimeInSeconds_(toWinSeconds) : "";
}

//
// Ppoints sorting function
// Each entry in the points array is a two item array
// where the second item is the actual points
//
function compare_points_(a, b) {
    return a[1] - b[1];
}

//
// Dumb integer sorting function
//
function compare_integers_(a, b) {
    return a - b;
}

// Rank a given corrected time relative to a supplied
// list of corrected times
//
// The supplied time and list of times are first converted
// to seconds. The list is sorted and then we find where
// within the list of times the input time is placed
//
// e.g. given time=1:10 and times=[1:09,1:08,1:10,1:13]
// we would return 3
//
// The special cases of an empty time value or a "DNF"
// is handled by just returning the input
function rank_(time, times) {
    if (!time || time == "DNF" || time == "DSQ" || time == "BFD" || time == "AVG") {
        return time;
    }
    time = dateToSeconds_(time);
    var timesInSeconds = [];
    for (var i = 0; i < times.length; i++) {
        var inSeconds = dateToSeconds_(times[i]);
        if (inSeconds > 0) {
            timesInSeconds.push(inSeconds);
        }
    }
    timesInSeconds.sort(compare_integers_);
    for (var i = 0; i < timesInSeconds.length; i++) {
        if (time <= timesInSeconds[i]) {
            return i+1;
        }
    }
    return "";
}

// Convert a given date object into the number of seconds
// since the start of the day
function dateToSeconds_(date) {
    try {
        var hours = date.getHours();
        var minutes = date.getMinutes() - MINUTES_DELTA;
        var seconds = date.getSeconds() - SECONDS_DELTA;
        return (hours*3600) + (minutes*60) + seconds;
    } catch (e) {
        //Logger.log("Failed to parse date: " + e);
        return -1;
    }
}

// Construct a new date object to represent a given
// time value
//
// The returned time is relative to a date in 1899
// because this appears to cause Google Docs to display
// the date as a simple time value
function newTime_(hours, minutes, seconds) {
    minutes += MINUTES_DELTA;
    seconds += SECONDS_DELTA;
    // 1899-12-30 is the epoch for time values, it seems
    // http://stackoverflow.com/questions/4051239/how-to-merge-date-and-time-as-a-datetime-in-google-apps-spreadsheet-script
    return new Date(1899, 11, 30, hours, minutes, seconds, 0);
}

// Construct a new date object to represent a given
// time value
function newTimeInSeconds_(seconds) {
    var hours = seconds / 3600; seconds = seconds % 3600;
    var minutes = seconds / 60; seconds = seconds % 60;
    return newTime_(hours, minutes, seconds);
}
