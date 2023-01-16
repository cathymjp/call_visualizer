var file;                           // the file to be processed
var workbook;                       // the workbook
var defaultDropColor = "#bc0033"    // default color of drop-area
var highlightedDropColor = "gray"   // color of drop-area when a file is being dragged over it
var jsonWorkbookEntries;            // the excel data in JSON form
var checkedMonths;                  // a boolean array indicating the user's selected months
var years;                          // an array of the years covered by the data
var callDataPerYear;                // an array of the call data organized into aggregate yearly entries
var callDataPerMonth;               // an array of the call data organized into aggregate monthly entries
var callDataPerDay;                 // an array of the call data organized into aggregate daily entries
var callDataPerHour;                // an array of the call data organized into aggregate hourly entries
var month;                          // the currently-selected month
var userOptions;                    // an object describing the user's input choices (along with default options)

// when the page document has finished loading, perform the following function
$(document).ready(function() {
	// no visualizations are initially expanded
	var expanded = false;

	// toggle the visualization upon clicking the submit button
	$("#submit-btn").click(function() {
		if (jsonWorkbookEntries && !expanded) {   // clicking on the "Generate Report" button only does anything if there is not yet a visualization
			var userOptions = processUserOptions();
			if (typeof(userOptions) != "string") { // if type IS a string, that means it's an error message. DON'T generate graphs.
				// remove any previous error messages
				d3.select("#initial-info")
					.select("h1")
					.remove();

				labelMissed(jsonWorkbookEntries, userOptions["missedCallRule"], false);   // determine which calls have been missed and label them as such
				// tally up the total number of missed calls
				var missedCounter = 0;
				var i = 0;
				for (var int = 0; i < jsonWorkbookEntries.length; i++) {
					if (jsonWorkbookEntries[i].Missed) {
						missedCounter++;
					}
				}

				// get the call data per year
				years = getActiveYears(jsonWorkbookEntries);
				callDataPerYear = getCallDataPerYear(jsonWorkbookEntries, years);

				// visualize the data at the year level
				visualizeYearLevel(callDataPerYear, jsonWorkbookEntries);

				// scroll to the bottom of the page (for a nice visual effect)
				$("body").delay(100).animate({ scrollTop: $(document).height()-$(window).height() }, 750);

				// the year-level visualization has been generated
				expanded = true;
			} else {   // there is some user error
				d3.select("#initial-info")
				.append("h1")
				.text("Error: " + userOptions);
			}
	} else if (!expanded) {
		d3.select("#initial-info")
		.append("h1")
		.text("No file selected!");
	}
});

// alert user of possible loss of functionality due to outdated browser
if (!window.File || !window.FileReader || !window.FileList) {
	alert("The File APIs are not fully supported in this browser.");
}
});

// handle drop events
var dropListener = {
	handleEvent: function(event) {
		if (event.type === 'dragenter') { this.onDragEnter(event); }
		if (event.type === 'dragleave') { this.onDragExit(event); }
		if (event.type === 'dragover') { this.onDragOver(event); }
		if (event.type === 'drop') { this.onDragDrop(event); }
	},

	onDragEnter: function(event) {
		event.stopPropagation();
		event.preventDefault();
		event.currentTarget.style.backgroundColor = highlightedDropColor;
	},

	onDragExit: function(event) {
		event.stopPropagation();
		event.preventDefault();
		event.currentTarget.style.backgroundColor = defaultDropColor;
	},

	onDragOver: function(event) {
		event.stopPropagation();
		event.preventDefault();
		event.dataTransfer.dropEffect = "copy"   // copy the dragged file, do not move it
	},

	onDragDrop: function(event) {
		event.stopPropagation();
		event.preventDefault();
		event.currentTarget.style.backgroundColor = defaultDropColor;

		// remove any previous error messages
		d3.select("#initial-info")
			.select("h1")
			.remove();

		// indicate to the user that the file is loading
		d3.select("#list")
			.append("h3")
			.text("Loading file...");

		file = event.dataTransfer.files[0];   // File object

		// read the file
		var reader = new FileReader();
		var name = file.name;
		reader.onload = function(event) {
			$("#list")[0].innerHTML = "<strong>Loading file...</strong>";

			workbook = XLSX.read(event.target.result, {type: 'binary'});
			var json_workbook = to_json(workbook);
			var sheetName = workbook.SheetNames[0];
			jsonWorkbookEntries = json_workbook[sheetName]; // Init main array of JSON call objects
			cleanJSONWorkbook(jsonWorkbookEntries); // clean the data
			var activeMonths = getActiveMonths(jsonWorkbookEntries);
			for (var i = 0; i < activeMonths.length; i++) {
				var monthName = convertMonth(i);
				var monthCheckbox = $("#" + monthName)
				if (!activeMonths[i]) { // month is not active
					monthCheckbox.attr("disabled", true)
				} else { // month is active
					monthCheckbox.prop("checked", true);
					//monthCheckbox.attr("") make bold
				}
			}
			checkedMonths = activeMonths;
			jsonWorkbookEntries = sortByDate(jsonWorkbookEntries);
			var minDate = dateToDashString(jsonWorkbookEntries[0].Date);
			var maxDate = dateToDashString(jsonWorkbookEntries[jsonWorkbookEntries.length - 1].Date);

			$("#start-calendar").attr("min", minDate);
			$("#start-calendar").attr("max", maxDate);
			$("#end-calendar").attr("min", minDate);
			$("#end-calendar").attr("max", maxDate);
			name_li = "<li><strong>Name: " + file.name + "</strong></li>";
			type_li = "<li><strong>Type: " + file.type + "</strong</li>";
			size_li = "<li><strong>Size: " + file.size + " bytes</strong></li>";
			$("#list")[0].innerHTML = "<ul>" + name_li + type_li + size_li + "</ul>";
		};
		reader.readAsBinaryString(file);
	}
};

// setup drag and drop listeners
var dropArea = $("#drop-area")[0];
dropArea.addEventListener('dragenter', dropListener, false);
dropArea.addEventListener('dragleave', dropListener, false);
dropArea.addEventListener('dragover', dropListener, false);
dropArea.addEventListener('drop', dropListener, false);

// convert given workbook to JSON
function to_json(workbook) {
	var result = {};
	workbook.SheetNames.forEach(function(sheetName) {
		var roa = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
		if(roa.length > 0){
			result[sheetName] = roa;
		}
	});
	return result;
}

/**
	Validates user input from month checkboxes, start and end date calendars, and the missed call rule input box.
	Returns an object with the following structure if ALL input is valid:
	{
		"checkedMonths" : [array of booleans (indices 0 to 11) indicating if a month was checked by the user].
		"missedCallRule" : integer (40 if not selected by user).
		"startDate" : Date object indicating the date (inclusive) from which the data should be analyzed.
		"endDate" : Date object indicating the date (inclusive) up to which data should be analyzed.
	}
	If input is NOT valid, then a error message (string) is returned.
**/
function processUserOptions() {
	userOptions = {};
	var errorMessage;
	// Process month checkboxes
	for (var i = 0; i < checkedMonths.length; i++) {
		var monthName = convertMonth(i);
		var isChecked = $("#" + monthName).is(":checked");
		checkedMonths[i] = isChecked;
	}
	userOptions["checkedMonths"] = checkedMonths;
	// At this point, checkedMonths contains valid info about which checkbox is checked
	var missedCallRule = $("#missed-call-rule").val();
	var missedCallRuleMin = $("#missed-call-rule").attr("min"); // 1 sec
	//console.debug(missedCallRuleMin)
	var missedCallRuleMax = $("#missed-call-rule").attr("max"); // 120 sec
	var missedCallRuleInt = parseInt(missedCallRule);
	// checking missedCallRule is valid
	// "" is when the form is not filled out
	if (missedCallRule != "" && (missedCallRuleInt < missedCallRuleMin || missedCallRuleInt > missedCallRuleMax)) {
		errorMessage = (missedCallRuleInt < missedCallRuleMin) ? "Missed call rule is less than 1 second." : "Missed call rule is more than 120 seconds"
		return errorMessage;
	} else {
		var defaultMissedCallRule = 40;
		userOptions["missedCallRule"] = (missedCallRule == "") ? defaultMissedCallRule : missedCallRuleInt;
	}
	// Checking calendar for validity of selected start and end dates
	var selectedStartDate = new Date($("#start-calendar").val());
	var startDateStamp = Date.parse(selectedStartDate);
	var selectedEndDate = new Date($("#end-calendar").val());
	var endDateStamp = Date.parse(selectedEndDate)
	var startMin = new Date($("#start-calendar").attr("min"));
	startMin.setUTCMilliseconds(startMin.getUTCMilliseconds() + 14400000);
	var startMax = new Date($("#start-calendar").attr("max"));
	var endMin = new Date($("#end-calendar").attr("min"));
	var endMax = new Date($("#end-calendar").attr("max"));
	endMax.setUTCMilliseconds(endMax.getUTCMilliseconds() + 100800000);
	if (isNaN(selectedStartDate) ^ isNaN(selectedEndDate)) {
		errorMessage = "Both or neither start and end dates must be selected.";
		return errorMessage;
	} else if (isNaN(selectedStartDate) && isNaN(selectedEndDate)) { // both dates are invalid or not selected
		userOptions["startDate"] = startMin;
		userOptions["endDate"] = endMax;
	} else {
		selectedStartDate.setUTCMilliseconds(selectedStartDate.getUTCMilliseconds() + 14400000);
		selectedEndDate.setUTCMilliseconds(selectedEndDate.getUTCMilliseconds() + 100800000);
		if (selectedStartDate.getTime() < startMin.getTime() || selectedStartDate.getTime() > startMax.getTime()) {
			errorMessage = "The start date selected is not in the range of data provided.";
			return errorMessage;
		} else if (selectedEndDate.getTime() < endMin.getTime() || selectedEndDate.getTime()  > endMax.getTime()) {
			errorMessage = "The end date selected is not in the range of data provided."
			return errorMessage;
		} else if (selectedStartDate.getTime() > selectedEndDate.getTime()) {
			errorMessage = "The selected start date is later than the selected end date."
			return errorMessage;
		} else {
			userOptions["startDate"] = selectedStartDate;
			userOptions["endDate"] = selectedEndDate
		}
	}

	return userOptions;
}


/**
	Convenience function that converts a date to dash-separated string to be included in html.
**/
function dateToDashString(date) {
	var day = date.getDate();
	var month = date.getMonth() + 1;
	var year = date.getFullYear();
	if (day < 10) {
		day = "0" + day
	}
	if (month < 10) {
		month = "0" + month;
	}
	var dateString = year + "-" + month + "-" + day;
	return dateString;
}
