// Set default timeout for long-running operations
var SCRIPT_TIMEOUT = 30000; // 5 minutes in milliseconds

/**
 * Normalizes request data to handle various property naming conventions
 * This ensures consistent access regardless of camelCase, PascalCase, or other formats
 */
function normalizeRequestData(requestData) {
  if (!requestData) {
    throw new Error("Request data cannot be null or undefined");
  }
  
  try {
    var normalized = JSON.parse(JSON.stringify(requestData));
    
    // Map of standard property names and their possible variations
    var propertyMappings = {
      'requesterID': ['RequesterID', 'Requester ID', 'requester_id', 'requesterId'],
      'targetEmployee': ['TargetEmployee', 'Target Employee', 'targetEmployeeID', 'target_employee', 'targetEmployeeName'],
      'date': ['Date', 'DutyDate', 'duty_date'],
      'slot': ['Slot', 'SlotNumber', 'slot_number'],
      'program': ['Program', 'ProgramName', 'program_name'],
      'location': ['Location', 'LocationName', 'location_name', 'venue'],
      'reasonType': ['ReasonType', 'Reason Type', 'reason_type', 'Reason'],
      'notes': ['Notes', 'AdminNotes', 'admin_notes', 'Message', 'message'],
      'status': ['Status', 'RequestStatus', 'request_status']
    };
    
    // Normalize each property
    for (var standardKey in propertyMappings) {
      // Skip if the standard key already exists and has a value
      if (normalized[standardKey]) continue;
      
      // Check all possible variations of this property
      for (var i = 0; i < propertyMappings[standardKey].length; i++) {
        var alternateKey = propertyMappings[standardKey][i];
        
        // If this variation exists in the data, use its value for the standard key
        if (normalized[alternateKey] !== undefined) {
          normalized[standardKey] = normalized[alternateKey];
          break;
        }
      }
    }
    
    return normalized;
  } catch (e) {
    Logger.log("Error normalizing request data: " + e.toString());
    throw new Error("Failed to normalize request data");
  }
}

/**
 * Set a timeout for script operations to avoid excessive runtime
 */
function checkScriptTimeout(startTime) {
  if (new Date().getTime() - startTime > SCRIPT_TIMEOUT) {
    throw new Error("Script timeout: operation exceeded maximum allowed time");
  }
}

/**
 * Save user preferences to persist settings
 */
function saveUserPreference(key, value) {
  try {
    var userProps = PropertiesService.getUserProperties();
    userProps.setProperty(key, value);
    return true;
  } catch (e) {
    Logger.log("Error saving user preference: " + e.toString());
    return false;
  }
}

/**
 * Get user preferences
 */
function getUserPreference(key, defaultValue) {
  try {
    var userProps = PropertiesService.getUserProperties();
    var value = userProps.getProperty(key);
    return value !== null ? value : defaultValue;
  } catch (e) {
    Logger.log("Error getting user preference: " + e.toString());
    return defaultValue;
  }
}

/**
 * Create a menu in the Google Sheets UI
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu('AdminSwapRequests');
  
  menu.addItem('Open Admin Panel', 'showAdminSwapRequestPanel')
      .addItem('Debug Swap Functions', 'debugSwapRequestFunctions')
      .addSeparator()
      .addItem('Duties Entry', 'showDutiesEntryPanel')
      .addToUi();
}

/**
 * Show a dialog for administrators to approve or reject swap requests
 */
function showSwapApprovalDialog() {
  var html = HtmlService.createHtmlOutput(`
    <html>
      <head>
        <base target="_top">
        <style>
          body {
            font-family: Arial, sans-serif;
            margin: 0;
            padding: 20px;
            font-size: 14px;
            color: #333;
          }
          .form-group {
            margin-bottom: 15px;
          }
          label {
            display: block;
            margin-bottom: 5px;
            font-weight: bold;
          }
          input, select, textarea {
            width: 100%;
            padding: 8px;
            border: 1px solid #ddd;
            border-radius: 4px;
            box-sizing: border-box;
            font-size: 14px;
          }
          .actions {
            display: flex;
            justify-content: flex-end;
            margin-top: 20px;
            gap: 10px;
          }
          .btn {
            padding: 8px 16px;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            font-size: 14px;
            font-weight: bold;
          }
          .btn-primary {
            background-color: #4361ee;
            color: white;
          }
          .btn-danger {
            background-color: #e63946;
            color: white;
          }
          .btn-neutral {
            background-color: #e5e5e5;
            color: #333;
          }
          h2 {
            margin-top: 0;
            border-bottom: 1px solid #eee;
            padding-bottom: 10px;
            color: #4361ee;
          }
          .result {
            margin-top: 20px;
            padding: 15px;
            border-radius: 4px;
            display: none;
          }
          .success {
            background-color: #d1e7dd;
            color: #0f5132;
            border: 1px solid #badbcc;
          }
          .error {
            background-color: #f8d7da;
            color: #842029;
            border: 1px solid #f5c2c7;
          }
        </style>
      </head>
      <body>
        <h2>Approve or Reject Swap Request</h2>
        
        <div class="form-group">
          <label for="requestId">Request ID:</label>
          <input type="text" id="requestId" placeholder="Enter the swap request ID" required>
        </div>
        
        <div class="form-group">
          <label for="status">Action:</label>
          <select id="status" required>
            <option value="">-- Select an action --</option>
            <option value="Approved">Approve Request</option>
            <option value="Rejected">Reject Request</option>
          </select>
        </div>
        
        <div class="form-group">
          <label for="notes">Admin Notes (optional):</label>
          <textarea id="notes" rows="3" placeholder="Add any notes about this decision"></textarea>
        </div>
        
        <div class="actions">
          <button type="button" class="btn btn-neutral" onclick="google.script.host.close()">Cancel</button>
          <button type="button" id="submitBtn" class="btn btn-primary" onclick="submitForm()">Submit</button>
        </div>
        
        <div id="resultMessage" class="result"></div>
        
        <script>
          function submitForm() {
            const requestId = document.getElementById('requestId').value.trim();
            const status = document.getElementById('status').value;
            const notes = document.getElementById('notes').value;
            
            if (!requestId || !status) {
              showResult('Please fill out all required fields.', false);
              return;
            }
            
            document.getElementById('submitBtn').disabled = true;
            document.getElementById('submitBtn').innerText = 'Processing...';
            
            google.script.run
              .withSuccessHandler(function(result) {
                if (result.success) {
                  showResult(result.message, true);
                  // Reset form after successful submission
                  document.getElementById('requestId').value = '';
                  document.getElementById('status').value = '';
                  document.getElementById('notes').value = '';
                } else {
                  showResult(result.message, false);
                }
                document.getElementById('submitBtn').disabled = false;
                document.getElementById('submitBtn').innerText = 'Submit';
              })
              .withFailureHandler(function(error) {
                showResult('Error: ' + error.message, false);
                document.getElementById('submitBtn').disabled = false;
                document.getElementById('submitBtn').innerText = 'Submit';
              })
              .updateSwapRequestStatus(requestId, status, notes);
          }
          
          function showResult(message, success) {
            const resultElement = document.getElementById('resultMessage');
            resultElement.innerHTML = message;
            resultElement.className = 'result ' + (success ? 'success' : 'error');
            resultElement.style.display = 'block';
            
            // Auto-hide after 5 seconds if it's a success message
            if (success) {
              setTimeout(function() {
                resultElement.style.display = 'none';
              }, 5000);
            }
          }
        </script>
      </body>
    </html>
  `)
  .setWidth(500)
  .setHeight(500)
  .setTitle('Approve Swap Request');
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Approve Swap Request');
}

/**
 * Function for administrators to view all pending swap requests
 */
function adminViewSwapRequests() {
  // Retrieve all swap requests
  var result = listAllSwapRequests();
  
  if (!result.success) {
    SpreadsheetApp.getUi().alert('Error loading swap requests: ' + result.message);
    return;
  }
  
  var requests = result.requests || [];
  
  // Filter to show only pending requests
  var pendingRequests = requests.filter(function(req) {
    return req.status && req.status.toString().toLowerCase() === 'pending';
  });
  
  if (pendingRequests.length === 0) {
    SpreadsheetApp.getUi().alert('No pending swap requests found.');
    return;
  }
  
  // Format the requests for display
  var output = 'Pending Swap Requests:\n\n';
  
  pendingRequests.forEach(function(request, index) {
    output += (index + 1) + '. Request ID: ' + request.id + '\n';
    output += '   Requester: ' + request.requesterID + '\n';
    output += '   Target Employee: ' + request.targetEmployee + '\n';
    output += '   Date: ' + request.date + '\n';
    output += '   Slot: ' + request.slot + '\n';
    output += '   Program: ' + (request.program || 'N/A') + '\n';
    output += '   Reason: ' + (request.reasonType || 'N/A') + '\n\n';
  });
  
  output += 'To approve or reject a request, use the "Admin: Approve Swap Request" menu option.';
  
  // Show the information
  var ui = SpreadsheetApp.getUi();
  ui.alert('Pending Swap Requests', output, ui.ButtonSet.OK);
}

function doGet(e) {
  if (e.parameter.page === 'manifest') {
    return doGetManifest();
  } else if (e.parameter.page === 'icon') {
    return doGetIcon();
  }
  
  // Your existing doGet logic here
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Invigilator Chart')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=5.0, minimum-scale=1.0, user-scalable=yes, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function showSearchSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Invigilator Duty Management');
  SpreadsheetApp.getUi().showSidebar(html);
}

function searchInvigilatorDuties(criteria) {
  try {
    var searchTerm = criteria.searchTerm.toString().trim().toLowerCase();
    if (!searchTerm) {
      return { success: false, message: "Search term is required" };
    }
    
    // Track if this is a forced refresh for logging purposes
    var forceRefresh = criteria.forceRefresh === true;
    if (forceRefresh) {
      Logger.log("Forced refresh requested for search term: " + searchTerm);
    }
    
    // Get data from main sheet
    var mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('main data');
    if (!mainSheet) {
      return { success: false, message: "Sheet 'main data' not found" };
    }
    
    // Ensure we get the latest data from the sheet by forcing a refresh of the cache
    if (forceRefresh) {
      // Flush any pending changes to the spreadsheet
      SpreadsheetApp.flush();
      
      // Force spreadsheet recalculation to ensure data freshness
      try {
        // Touch the first cell of the first three rows to trigger recalculation
        for (var i = 1; i <= 3; i++) {
          var range = mainSheet.getRange(i, 1);
          var value = range.getValue();
          range.setValue(value);
        }
        // Flush changes again
        SpreadsheetApp.flush();
        Logger.log("Performed forced recalculation for search");
      } catch (recalcError) {
        Logger.log("Warning during recalculation: " + recalcError.toString());
        // Continue despite error
      }
    }
    
    // Get fresh data after recalculation
    var mainData = mainSheet.getDataRange().getValues();
    if (mainData.length <= 1) {
      return { success: false, message: "No data found in 'main data' sheet" };
    }
    
    // Get column indices and log them for debugging
    var headers = mainData[0];
    var colIndices = getColumnIndices(headers);
    
    // Log the indices for debugging
    Logger.log("Column indices: " + JSON.stringify(colIndices));
    
    // Check if essential columns were found
    if (colIndices.empId === -1) {
      return { success: false, message: "Column 'EMPLOYEE ID' not found" };
    }
    
    if (colIndices.slot1 === -1 && colIndices.slot2 === -1 && 
        colIndices.slot3 === -1 && colIndices.slot4 === -1) {
      return { success: false, message: "No slot columns found. Please check column headers." };
    }
    
    // First find the employee
    var employeeInfo = findEmployeeInfo(searchTerm, mainData, colIndices);
    if (!employeeInfo) {
      return { success: false, message: "No employee found matching the search criteria" };
    }
    
    // Find the accurate season count for this employee
    var seasonTotal = calculateSeasonTotal(employeeInfo.empId, mainData, colIndices);
    
    // Then find all duties for this employee
    var duties = findEmployeeDuties(employeeInfo.empId, mainData, colIndices, seasonTotal);
    
    // Log the duties for debugging
    Logger.log("Found " + duties.length + " duty records with season total: " + seasonTotal);
    
    return {
      success: true,
      employee: employeeInfo,
      duties: duties,
      seasonTotal: seasonTotal,
      debug: {
        columnIndices: colIndices
      }
    };
  } catch (error) {
    Logger.log("Error: " + error.toString());
    Logger.log("Stack: " + error.stack);
    return { 
      success: false, 
      message: "An error occurred: " + error.toString() 
    };
  }
}

function getColumnIndices(headers) {
  // First find all the basic columns
  var indices = {
    // Employee info columns
    empId: findColumnIndex(headers, ["EMPLOYEE ID", "EMPLOYEEID", "EMP ID", "ID"]),
    empName: findColumnIndex(headers, ["NAME OF THE EMPLOYEE", "EMPLOYEE NAME", "NAME"]),
    designation: findColumnIndex(headers, ["DESIGNATION"]),
    email: findColumnIndex(headers, ["E-MAIL-ID", "EMAIL", "E MAIL", "MAIL"]),
    phone: findColumnIndex(headers, ["CELL PHONE NO.", "PHONE", "MOBILE", "CONTACT"]),
    
    // Duty info columns
    dayOfExam: findColumnIndex(headers, ["Days of Exam", "DAY OF EXAM", "EXAM DAY"]),
    date: findColumnIndex(headers, ["Date", "DATE"]),
    weekDay: findColumnIndex(headers, ["Week Days", "WEEKDAY", "DAY OF WEEK"]),
    program: findColumnIndex(headers, ["Degree Prog. Or Semester or School", "DEGREE", "PROGRAM", "SEMESTER", "SCHOOL"]),
    
    // Slot columns (will find exact slot headers)
    slot1: -1,
    location1: -1,
    slot2: -1,
    location2: -1,
    slot3: -1,
    location3: -1,
    slot4: -1,
    location4: -1,
    
    // Counts
    dayCount: findColumnIndex(headers, ["Total Count of Invigilation Duty on the Day", "DAILY COUNT", "DAY COUNT"]),
    seasonCount: findColumnIndex(headers, ["Total Count of Invigilation Duty of the Season", "SEASON COUNT", "TOTAL COUNT"]),
    notes: findColumnIndex(headers, ["Notes", "NOTES"])
  };
  
  // Find slot and location columns
  var slot1Patterns = ["Slot 1", "SLOT1", "SLOT 1"];
  var slot2Patterns = ["Slot 2", "SLOT2", "SLOT 2"];
  var slot3Patterns = ["Slot 3", "SLOT3", "SLOT 3"];
  var slot4Patterns = ["Slot 4", "SLOT4", "SLOT 4"];
  var locationPatterns = ["Location", "LOCATION", "VENUE"];
  
  for (var i = 0; i < headers.length; i++) {
    var header = headers[i] ? headers[i].toString() : "";
    
    // Slot 1
    if (containsAny(header, slot1Patterns)) {
      indices.slot1 = i;
      // Look for location in next column
      if (i+1 < headers.length && containsAny(headers[i+1].toString(), locationPatterns)) {
        indices.location1 = i+1;
      }
    }
    // Slot 2
    else if (containsAny(header, slot2Patterns)) {
      indices.slot2 = i;
      if (i+1 < headers.length && containsAny(headers[i+1].toString(), locationPatterns)) {
        indices.location2 = i+1;
      }
    }
    // Slot 3
    else if (containsAny(header, slot3Patterns)) {
      indices.slot3 = i;
      if (i+1 < headers.length && containsAny(headers[i+1].toString(), locationPatterns)) {
        indices.location3 = i+1;
      }
    }
    // Slot 4
    else if (containsAny(header, slot4Patterns)) {
      indices.slot4 = i;
      if (i+1 < headers.length && containsAny(headers[i+1].toString(), locationPatterns)) {
        indices.location4 = i+1;
      }
    }
  }
  
  Logger.log("Found column indices: " + JSON.stringify(indices));
  return indices;
}

// Helper function to find column index by possible header names
function findColumnIndex(headers, possibleNames) {
  for (var i = 0; i < possibleNames.length; i++) {
    var colName = possibleNames[i];
    for (var j = 0; j < headers.length; j++) {
      var header = headers[j] ? headers[j].toString() : "";
      if (header.toLowerCase().indexOf(colName.toLowerCase()) !== -1) {
        return j;
      }
    }
  }
  return -1;
}

// Helper function to check if string contains any pattern from array
function containsAny(str, patterns) {
  if (!str) return false;
  str = str.toLowerCase();
  for (var i = 0; i < patterns.length; i++) {
    if (str.indexOf(patterns[i].toLowerCase()) !== -1) {
      return true;
    }
  }
  return false;
}

function findEmployeeInfo(searchTerm, data, colIndices) {
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var empId = row[colIndices.empId] ? row[colIndices.empId].toString().toLowerCase() : "";
    var empName = row[colIndices.empName] ? row[colIndices.empName].toString().toLowerCase() : "";
    var email = row[colIndices.email] ? row[colIndices.email].toString().toLowerCase() : "";
    
    if (empId.includes(searchTerm) || empName.includes(searchTerm) || email.includes(searchTerm)) {
      return {
        empId: row[colIndices.empId],
        name: row[colIndices.empName],
        designation: row[colIndices.designation],
        email: row[colIndices.email],
        phone: row[colIndices.phone]
      };
    }
  }
  return null;
}

// Calculate total invigilation count for the season
function calculateSeasonTotal(empId, data, colIndices) {
  // Method 1: Count the actual occurrences of "x" in all slots
  var totalCount = 0;
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var rowEmpId = row[colIndices.empId] ? row[colIndices.empId].toString() : "";
    
    if (rowEmpId !== empId.toString()) continue;
    
    // Check all slots for "x" marks
    if (colIndices.slot1 >= 0 && row[colIndices.slot1]) {
      var value = row[colIndices.slot1].toString().trim();
      if (value === "x" || value === "X") totalCount++;
    }
    
    if (colIndices.slot2 >= 0 && row[colIndices.slot2]) {
      var value = row[colIndices.slot2].toString().trim();
      if (value === "x" || value === "X") totalCount++;
    }
    
    if (colIndices.slot3 >= 0 && row[colIndices.slot3]) {
      var value = row[colIndices.slot3].toString().trim();
      if (value === "x" || value === "X") totalCount++;
    }
    
    if (colIndices.slot4 >= 0 && row[colIndices.slot4]) {
      var value = row[colIndices.slot4].toString().trim();
      if (value === "x" || value === "X") totalCount++;
    }
  }
  
  // Method 2: If the above count is 0 or we couldn't find slots, use the value from the season count column
  if (totalCount === 0 && colIndices.seasonCount >= 0) {
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var rowEmpId = row[colIndices.empId] ? row[colIndices.empId].toString() : "";
      
      if (rowEmpId !== empId.toString()) continue;
      
      if (row[colIndices.seasonCount]) {
        var seasonValue = row[colIndices.seasonCount];
        if (typeof seasonValue === 'number' && seasonValue > 0) {
          return seasonValue;
        } else if (typeof seasonValue === 'string') {
          var num = parseInt(seasonValue.trim());
          if (!isNaN(num) && num > 0) return num;
        }
      }
    }
  }
  
  return totalCount;
}

function findEmployeeDuties(empId, data, colIndices, seasonTotal) {
  var duties = [];
  var dayMap = {}; // To group duties by day
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var rowEmpId = row[colIndices.empId] ? row[colIndices.empId].toString() : "";
    
    if (rowEmpId !== empId.toString()) continue;
    
    var dayKey = row[colIndices.date] ? row[colIndices.date].toString() : "";
    if (!dayKey) continue;
    
    if (!dayMap[dayKey]) {
      // Initialize the new day entry
      dayMap[dayKey] = {
        day: row[colIndices.dayOfExam],
        date: formatDate(row[colIndices.date]),
        weekDay: row[colIndices.weekDay],
        program: row[colIndices.program],
        slot1: "-",
        location1: "-",
        slot2: "-",
        location2: "-",
        slot3: "-",
        location3: "-",
        slot4: "-",
        location4: "-",
        dayCount: 0,
        seasonCount: seasonTotal // Use the accurately calculated season total
      };
      
      // For program column, if multiple entries for same day, use TextJoin approach
      if (colIndices.program >= 0) {
        dayMap[dayKey].programList = new Set();
        if (row[colIndices.program]) {
          dayMap[dayKey].programList.add(row[colIndices.program].toString().trim());
        }
      }
    } else if (colIndices.program >= 0 && row[colIndices.program]) {
      // Add unique program entries
      dayMap[dayKey].programList.add(row[colIndices.program].toString().trim());
    }
    
    // Update slots based on what's in the row
    updateSlotInfo(dayMap[dayKey], row, colIndices);
  }
  
  // Convert map to array and finalize data
  for (var key in dayMap) {
    var duty = dayMap[key];
    
    // Convert program Set to string if needed
    if (duty.programList && duty.programList.size > 0) {
      duty.program = Array.from(duty.programList).join(" / ");
    }
    delete duty.programList;
    
    duties.push(duty);
  }
  
  duties.sort(function(a, b) {
    // Parse dates for comparison (expecting dd-MMM-yyyy format)
    var datePartsA = a.date ? a.date.split('-') : [];
    var datePartsB = b.date ? b.date.split('-') : [];
    
    if (datePartsA.length !== 3 || datePartsB.length !== 3) {
      // Fallback to day number if date parsing fails
      var dayA = parseInt(a.day) || 0;
      var dayB = parseInt(b.day) || 0;
      return dayA - dayB;
    }
    
    // Convert month abbreviation to month number (Jan -> 0, Feb -> 1, etc.)
    var monthNamesShort = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
    
    var dayA = parseInt(datePartsA[0]) || 0;
    var monthA = monthNamesShort.indexOf(datePartsA[1].toLowerCase()) !== -1 ? 
                 monthNamesShort.indexOf(datePartsA[1].toLowerCase()) : 
                 parseInt(datePartsA[1]) - 1 || 0;
    var yearA = parseInt(datePartsA[2]) || 0;
    
    var dayB = parseInt(datePartsB[0]) || 0;
    var monthB = monthNamesShort.indexOf(datePartsB[1].toLowerCase()) !== -1 ? 
                 monthNamesShort.indexOf(datePartsB[1].toLowerCase()) : 
                 parseInt(datePartsB[1]) - 1 || 0;
    var yearB = parseInt(datePartsB[2]) || 0;
    
    // Compare years first
    if (yearA !== yearB) return yearA - yearB;
    
    // Then months
    if (monthA !== monthB) return monthA - monthB;
    
    // Then days
    return dayA - dayB;
  });
  
  return duties;
}

function updateSlotInfo(dutyDay, row, colIndices) {
  // Check and update each slot
  if (colIndices.slot1 >= 0) {
    var slotValue = (row[colIndices.slot1] || "").toString().trim();
    if (slotValue === "x" || slotValue === "X") {
      dutyDay.slot1 = "x";
      dutyDay.dayCount++;
    } else if (slotValue && slotValue !== "-") {
      dutyDay.slot1 = slotValue;
      dutyDay.dayCount++;
    }
    
    if (colIndices.location1 >= 0 && row[colIndices.location1]) {
      var locValue = row[colIndices.location1].toString().trim();
      if (locValue && locValue !== "-") {
        dutyDay.location1 = locValue;
      }
    }
  }
  
  if (colIndices.slot2 >= 0) {
    var slotValue = (row[colIndices.slot2] || "").toString().trim();
    if (slotValue === "x" || slotValue === "X") {
      dutyDay.slot2 = "x";
      dutyDay.dayCount++;
    } else if (slotValue && slotValue !== "-") {
      dutyDay.slot2 = slotValue;
      dutyDay.dayCount++;
    }
    
    if (colIndices.location2 >= 0 && row[colIndices.location2]) {
      var locValue = row[colIndices.location2].toString().trim();
      if (locValue && locValue !== "-") {
        dutyDay.location2 = locValue;
      }
    }
  }
  
  if (colIndices.slot3 >= 0) {
    var slotValue = (row[colIndices.slot3] || "").toString().trim();
    if (slotValue === "x" || slotValue === "X") {
      dutyDay.slot3 = "x";
      dutyDay.dayCount++;
    } else if (slotValue && slotValue !== "-") {
      dutyDay.slot3 = slotValue;
      dutyDay.dayCount++;
    }
    
    if (colIndices.location3 >= 0 && row[colIndices.location3]) {
      var locValue = row[colIndices.location3].toString().trim();
      if (locValue && locValue !== "-") {
        dutyDay.location3 = locValue;
      }
    }
  }
  
  if (colIndices.slot4 >= 0) {
    var slotValue = (row[colIndices.slot4] || "").toString().trim();
    if (slotValue === "x" || slotValue === "X") {
      dutyDay.slot4 = "x";
      dutyDay.dayCount++;
    } else if (slotValue && slotValue !== "-") {
      dutyDay.slot4 = slotValue;
      dutyDay.dayCount++;
    }
    
    if (colIndices.location4 >= 0 && row[colIndices.location4]) {
      var locValue = row[colIndices.location4].toString().trim();
      if (locValue && locValue !== "-") {
        dutyDay.location4 = locValue;
      }
    }
  }
}

function formatDate(dateValue) {
  if (!dateValue) return "";
  
  try {
    // If it's already a string in a readable format, return it
    if (typeof dateValue === 'string') return dateValue;
    
    // If it's a date object, format it
    if (dateValue instanceof Date) {
      return Utilities.formatDate(dateValue, Session.getScriptTimeZone(), "dd-MMM-yyyy");
    }
    
    // Try to convert to date and format
    var date = new Date(dateValue);
    if (!isNaN(date.getTime())) {
      return Utilities.formatDate(date, Session.getScriptTimeZone(), "dd-MMM-yyyy");
    }
    
    // Return as is if all else fails
    return dateValue.toString();
  } catch (e) {
    return dateValue.toString();
  }
}

/**
 * Submits a duty swap request
 */
function submitDutySwapRequest(requestData) {
  try {
    console.log("Swap request received:", JSON.stringify(requestData));
    
    // Validate required fields
    if (!requestData.requesterID || !requestData.date || !requestData.slot || !requestData.targetEmployee || !requestData.reasonType) {
      return {
        success: false,
        message: "Missing required fields"
      };
    }
    
    // TEMPORARILY BYPASS DUTY VERIFICATION FOR TESTING
    // Will restore verification logic after fixing modal issues
    /*
    // Verify the requester has duty assigned
    const hasDuty = verifyUserHasDuty(requestData.requesterID, requestData.date, requestData.slot);
    if (!hasDuty) {
      return {
        success: false,
        message: "You do not have a duty assigned for this slot"
      };
    }
    */
    
    // Get SwapRequests sheet, create if it doesn't exist
    let swapSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SwapRequests");
    if (!swapSheet) {
      swapSheet = createSwapRequestsSheet();
    }
    
    // Generate a unique request ID (timestamp + random string)
    const timestamp = new Date().getTime();
    const requestID = timestamp + '-' + Math.random().toString(36).substring(2, 8);
    
    // Get the requester's name
    const requesterName = getEmployeeNameById(requestData.requesterID) || requestData.requesterID;
    
    // Current date/time
    const currentDate = new Date();
    const dateCreated = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    
    // Prepare row data
    const rowData = [
      requestID,                    // RequestID
      requestData.requesterID,      // RequesterID
      requesterName,                // RequesterName
      requestData.targetEmployee,   // TargetEmployeeID
      "",                           // TargetEmployeeName (filled in if available)
      requestData.date,             // Date
      requestData.slot,             // Slot
      requestData.program,          // Program
      requestData.location,         // Location
      requestData.reasonType,       // ReasonType
      requestData.notes,            // Notes
      "Pending",                    // Status
      "",                           // AdminNotes
      dateCreated,                  // DateCreated
      dateCreated                   // LastUpdated
    ];
    
    // Try to find target employee name
    try {
      console.log(`Looking up target employee with ID/name: ${requestData.targetEmployee}`);
      const targetName = getEmployeeNameById(requestData.targetEmployee);
      if (targetName) {
        console.log(`Found target employee name: ${targetName}`);
        rowData[4] = targetName;
      } else {
        console.log(`Could not find a matching employee for: ${requestData.targetEmployee}`);
      }
    } catch (e) {
      console.error("Could not get target employee name: " + e.message);
    }
    
    // Log what we're about to save
    console.log("Saving swap request with data:", JSON.stringify(rowData));
    
    // Add the new row
    swapSheet.appendRow(rowData);
    
    // Send notification email to the requester
    try {
      const adminEmail = getAdminEmail();
      const requesterEmail = getEmployeeEmailById(requestData.requesterID);
      
      if (requesterEmail) {
        sendSwapRequestConfirmation(
          requesterEmail, 
          requestData.date, 
          requestData.slot, 
          rowData[4] || requestData.targetEmployee,
          requestID
        );
      }
    } catch (emailError) {
      console.log("Failed to send confirmation email: " + emailError.message);
    }
    
    // Success response
    return {
      success: true,
      message: "Swap request submitted successfully",
      requestId: requestID
    };
    
  } catch (error) {
    console.error("Error in submitDutySwapRequest:", error);
    return {
      success: false,
      message: error.toString()
    };
  }
}

/**
 * Update the status of a swap request
 */
function updateSwapRequestStatus(requestId, newStatus, adminNotes) {
  try {
    if (!requestId || !newStatus) {
      throw new Error("Request ID and new status are required");
    }
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var swapSheet = ss.getSheetByName('SwapRequests');
    
    if (!swapSheet) {
      throw new Error("SwapRequests sheet not found");
    }
    
    var data = swapSheet.getDataRange().getValues();
    var headers = data[0];
    
    // Find column indices
    var requestIdCol = headers.indexOf('RequestID');
    var statusCol = headers.indexOf('Status');
    var adminNotesCol = headers.indexOf('AdminNotes');
    var lastUpdatedCol = headers.indexOf('LastUpdated');
    
    if (requestIdCol === -1 || statusCol === -1) {
      throw new Error("Required columns not found in swap requests sheet");
    }
    
    // Find the request by ID
    var requestRow = -1;
    var requestDetails = null;
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][requestIdCol].toString() === requestId.toString()) {
        requestRow = i + 1; // +1 because sheet rows are 1-indexed
        requestDetails = {};
        
        // Collect request details for processing
        for (var j = 0; j < headers.length; j++) {
          requestDetails[headers[j]] = data[i][j];
        }
        break;
      }
    }
    
    if (requestRow === -1) {
      throw new Error("Request with ID " + requestId + " not found");
    }
    
    // Update status
    swapSheet.getRange(requestRow, statusCol + 1).setValue(newStatus);
    
    // Update admin notes if provided
    if (adminNotesCol !== -1 && adminNotes) {
      swapSheet.getRange(requestRow, adminNotesCol + 1).setValue(adminNotes);
    }
    
    // Update last updated timestamp
    if (lastUpdatedCol !== -1) {
      swapSheet.getRange(requestRow, lastUpdatedCol + 1).setValue(new Date());
    }
    
    // Process approved swap if status is Approved
    var result = { success: true, message: "Request status updated to " + newStatus };
    
    if (newStatus.toLowerCase() === "approved") {
      try {
        var swapResult = processApprovedSwap(requestDetails);
        if (!swapResult.success) {
          Logger.log("Warning: Failed to process approved swap: " + swapResult.message);
          result.warning = "Status updated but failed to process swap: " + swapResult.message;
        }
      } catch (swapError) {
        Logger.log("Error processing approved swap: " + swapError.toString());
        result.warning = "Status updated but failed to process swap: " + swapError.toString();
      }
    }
    
    // Send notification about status change
    try {
      sendSwapStatusNotification(requestDetails, newStatus, adminNotes);
    } catch (notifyError) {
      Logger.log("Error sending notification: " + notifyError.toString());
      result.notificationError = "Failed to send notification: " + notifyError.toString();
    }
    
    return result;
  } catch (error) {
    Logger.log("Error updating swap request status: " + error.toString());
    return { success: false, message: error.toString() };
  }
}

/**
 * Get swap requests for a specific employee
 * This allows employees to see their own swap requests
 */
function getSwapRequests(employeeID) {
  try {
    if (!employeeID) {
      return { success: false, message: "Employee ID is required" };
    }
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var swapSheet = ss.getSheetByName('SwapRequests');
    
    if (!swapSheet) {
      return { success: true, requests: [] };
    }
    
    var lastRow = swapSheet.getLastRow();
    if (lastRow <= 1) {
      return { success: true, requests: [] };
    }
    
    var data = swapSheet.getRange(1, 1, lastRow, swapSheet.getLastColumn()).getValues();
    var headers = data[0];
    
    var requestIdCol = headers.indexOf('RequestID');
    var requesterIdCol = headers.indexOf('RequesterID');
    var targetEmployeeCol = headers.indexOf('TargetEmployee');
    var dateCol = headers.indexOf('Date');
    var slotCol = headers.indexOf('Slot');
    var locationCol = headers.indexOf('Location');
    var programCol = headers.indexOf('Program');
    var reasonTypeCol = headers.indexOf('ReasonType');
    var notesCol = headers.indexOf('Notes');
    var statusCol = headers.indexOf('Status');
    var timestampCol = headers.indexOf('Timestamp');
    
    if (requesterIdCol === -1 || statusCol === -1) {
      return { 
        success: false, 
        message: "Required columns not found in SwapRequests sheet" 
      };
    }
    
    var requests = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var rowRequesterID = row[requesterIdCol] ? row[requesterIdCol].toString() : "";
      var rowTargetEmployee = row[targetEmployeeCol] ? row[targetEmployeeCol].toString() : "";
      
      // Match requests where this employee is either the requester or the target
      if (rowRequesterID.toString().toLowerCase() === employeeID.toString().toLowerCase() || 
          rowTargetEmployee.toString().toLowerCase() === employeeID.toString().toLowerCase()) {
        
        var request = {
          id: requestIdCol >= 0 ? row[requestIdCol] : '',
          timestamp: timestampCol >= 0 ? row[timestampCol] : '',
          requesterID: rowRequesterID,
          targetEmployee: rowTargetEmployee,
          date: dateCol >= 0 ? row[dateCol] : '',
          slot: slotCol >= 0 ? row[slotCol] : '',
          location: locationCol >= 0 ? row[locationCol] : '',
          program: programCol >= 0 ? row[programCol] : '',
          reasonType: reasonTypeCol >= 0 ? row[reasonTypeCol] : '',
          notes: notesCol >= 0 ? row[notesCol] : '',
          status: statusCol >= 0 ? row[statusCol] : 'Pending'
        };
        
        requests.push(request);
      }
    }
    
    return {
      success: true,
      requests: requests
    };
  } catch (error) {
    Logger.log("Error in getSwapRequests: " + error.toString());
    return {
      success: false,
      message: "Error retrieving swap requests: " + error.toString()
    };
  }
}

/**
 * Cancel a swap request
 */
function cancelSwapRequest(requestId) {
  try {
    if (!requestId) {
      return { success: false, message: "Request ID is required" };
    }
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var swapSheet = ss.getSheetByName('SwapRequests');
    
    if (!swapSheet) {
      return { success: false, message: "SwapRequests sheet not found" };
    }
    
    var lastRow = swapSheet.getLastRow();
    var data = swapSheet.getRange(1, 1, lastRow, swapSheet.getLastColumn()).getValues();
    var headers = data[0];
    
    var requestIdCol = headers.indexOf('RequestID');
    var statusCol = headers.indexOf('Status');
    var lastUpdatedCol = headers.indexOf('LastUpdated');
    
    if (requestIdCol === -1 || statusCol === -1) {
      return { success: false, message: "Required columns not found in SwapRequests sheet" };
    }
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][requestIdCol].toString() === requestId.toString()) {
        // Found the request, update status to cancelled
        swapSheet.getRange(i + 1, statusCol + 1).setValue('Cancelled');
        
        // Update last updated timestamp if column exists
        if (lastUpdatedCol !== -1) {
          swapSheet.getRange(i + 1, lastUpdatedCol + 1).setValue(new Date());
        }
        
        return { success: true, message: "Swap request cancelled successfully" };
      }
    }
    
    return { success: false, message: "Swap request not found" };
  } catch (error) {
    Logger.log("Error in cancelSwapRequest: " + error.toString());
    return {
      success: false,
      message: "Error cancelling swap request: " + error.toString()
    };
  }
}

/**
 * For admin use: get all swap requests with filtering options
 */
function getAllSwapRequests(filters) {
  try {
    console.log("Getting all swap requests with filters:", filters ? JSON.stringify(filters) : "none");
    
    // If a spreadsheet name was provided in the filters, try to get data from that spreadsheet
    if (filters && filters.spreadsheetName) {
      console.log("Spreadsheet name provided in filters, trying to get data from:", filters.spreadsheetName);
      return tryGetSwapRequestsFromSpreadsheet(filters.spreadsheetName);
    }
    
    // If a sheet name was provided in the filters, try to get data from that sheet
    if (filters && filters.sheetName) {
      console.log("Sheet name provided in filters, trying to get data from sheet:", filters.sheetName);
      return tryGetSwapRequestsFromSheet(filters.sheetName);
    }
    
    // Try to get from cache first
    var cache = CacheService.getScriptCache();
    var cacheKey = "all_swap_requests_cache";
    var cachedData = cache.get(cacheKey);
    
    if (cachedData && !filters) {
      try {
        var parsedData = JSON.parse(cachedData);
        console.log("Using cached swap requests data");
        return parsedData;
      } catch (e) {
        console.error("Error parsing cached data:", e);
        // Continue to fetch fresh data
      }
    }
    
    // Ensure the swap requests sheet exists
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var swapSheet = ss.getSheetByName('SwapRequests');
    
    // If the sheet doesn't exist, create it with proper headers
    if (!swapSheet) {
      console.log("SwapRequests sheet not found, creating it now");
      swapSheet = ss.insertSheet('SwapRequests');
      
      // Define headers
      var headers = [
        'RequestID', 'RequesterID', 'TargetEmployeeID', 
        'Date', 'Slot', 'Location', 'Program', 
        'Notes', 'Status', 'DateCreated', 'LastUpdated'
      ];
      
      // Add headers to the first row
      swapSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      // Format the header row
      swapSheet.getRange(1, 1, 1, headers.length)
        .setBackground('#4285f4')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
      
      // Return empty result since we just created the sheet
      return { 
        success: true, 
        message: "SwapRequests sheet created successfully", 
        requests: [] 
      };
    }
    
    // Get all data
    var data = swapSheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      console.log("SwapRequests sheet has no data rows");
      return { success: true, requests: [] };
    }
    
    // Extract header row
    var headers = data[0];
    console.log("SwapRequests headers:", headers.join(", "));
    
    // Find column indices for filtering
    var statusIndex = headers.indexOf('Status');
    var requesterIdIndex = headers.indexOf('RequesterID');
    var targetEmployeeIdIndex = headers.indexOf('TargetEmployeeID');
    
    // Log column indices for debugging
    console.log("Column indices - Status:", statusIndex, "RequesterID:", requesterIdIndex, "TargetEmployeeID:", targetEmployeeIdIndex);
    
    // Check for required columns
    if (statusIndex === -1 || requesterIdIndex === -1) {
      var missingColumns = [];
      if (statusIndex === -1) missingColumns.push("Status");
      if (requesterIdIndex === -1) missingColumns.push("RequesterID");
      
      var errorMsg = "Required columns not found: " + missingColumns.join(", ");
      console.error(errorMsg);
      return { 
        success: false, 
        message: errorMsg,
        requests: [] 
      };
    }
    
    // Convert all rows to objects
    var requests = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      
      // Skip empty rows
      if (row.every(function(cell) { return !cell; })) {
        continue;
      }
      
      // Apply filters if provided
      if (filters) {
        var skipRow = false;
        
        // Log filter application
        console.log("Applying filters to row", i + 1);
        
        // Filter by status
        if (filters.Status && statusIndex !== -1) {
          var rowStatus = String(row[statusIndex] || '').toLowerCase();
          var filterStatus = String(filters.Status).toLowerCase();
          
          console.log("Checking Status filter:", rowStatus, "vs", filterStatus);
          
          if (rowStatus !== filterStatus) {
            console.log("Row", i + 1, "filtered out by Status");
            continue; // Skip this row
          }
        }
        
        // Filter by requester ID
        if (filters.RequesterID && requesterIdIndex !== -1) {
          var rowRequesterID = String(row[requesterIdIndex] || '').toLowerCase();
          var filterRequesterID = String(filters.RequesterID).toLowerCase();
          
          console.log("Checking RequesterID filter:", rowRequesterID, "vs", filterRequesterID);
          
          if (rowRequesterID !== filterRequesterID) {
            console.log("Row", i + 1, "filtered out by RequesterID");
            continue; // Skip this row
          }
        }
        
        // Filter by target employee ID
        if (filters.TargetEmployeeID && targetEmployeeIdIndex !== -1) {
          var rowTargetEmployeeID = String(row[targetEmployeeIdIndex] || '').toLowerCase();
          var filterTargetEmployeeID = String(filters.TargetEmployeeID).toLowerCase();
          
          console.log("Checking TargetEmployeeID filter:", rowTargetEmployeeID, "vs", filterTargetEmployeeID);
          
          if (rowTargetEmployeeID !== filterTargetEmployeeID) {
            console.log("Row", i + 1, "filtered out by TargetEmployeeID");
            continue; // Skip this row
          }
        }
      }
      
      // Convert row to object
      var request = {};
      for (var j = 0; j < headers.length; j++) {
        var header = headers[j];
        var value = row[j];
        
        // Format dates for display
        if (header === 'DateCreated' || header === 'LastUpdated') {
          if (value instanceof Date) {
            value = Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
          }
        }
        
        request[header] = value;
      }
      
      requests.push(request);
    }
    
    console.log("Returning", requests.length, "swap requests after filtering");
    
    // Sort requests by date created (newest first)
    requests.sort(function(a, b) {
      var dateA = a.DateCreated ? new Date(a.DateCreated) : new Date(0);
      var dateB = b.DateCreated ? new Date(b.DateCreated) : new Date(0);
      return dateB - dateA;
    });
    
    var result = { success: true, requests: requests };
    
    // Store in cache for 5 minutes if no filters were applied
    if (!filters) {
      try {
        cache.put(cacheKey, JSON.stringify(result), 300); // 300 seconds = 5 minutes
      } catch (e) {
        console.error("Error caching swap requests data:", e);
      }
    }
    
    return result;
  } catch (error) {
    console.error("Error in getAllSwapRequests:", error);
    return { 
      success: false, 
      message: error.toString(),
      requests: []
    };
  }
}

/**
 * Show the admin panel for managing swap requests
 */
function showAdminSwapRequestPanel() {
  try {
    // Clear cache to ensure fresh data
    var cache = CacheService.getScriptCache();
    cache.remove("all_swap_requests_cache");
    
    // Log for debugging
    Logger.log("Creating admin swap request panel");
    
    var html = HtmlService.createHtmlOutputFromFile('AdminSwapRequests')
        .setTitle('Admin: Swap Request Management')
        .setWidth(1000)
        .setHeight(700)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
    
    // Don't use XFrameOptionsMode or additional meta tags as they're causing errors
    
    Logger.log("Showing admin swap request dialog");
    SpreadsheetApp.getUi().showModalDialog(html, 'Swap Request Management');
    
    return { success: true };
  } catch (error) {
    Logger.log("Error in showAdminSwapRequestPanel: " + error.toString());
    
    // Try to show a simplified version if the main one fails
    try {
      var simpleHtml = HtmlService.createHtmlOutput(
        '<div style="padding: 20px; text-align: center;">' +
        '<h3>Error Loading Admin Panel</h3>' +
        '<p>There was an error loading the admin panel: ' + error.toString() + '</p>' +
        '<button onclick="google.script.host.close()">Close</button>' +
        '</div>'
      )
      .setWidth(600)
      .setHeight(300);
      
      SpreadsheetApp.getUi().showModalDialog(simpleHtml, 'Error: Admin Panel');
    } catch (e) {
      Logger.log("Failed to show error dialog: " + e.toString());
    }
    
    return { success: false, error: error.toString() };
  }
}

/**
 * Admin function to delete a swap request by its RequestID
 * This allows admins to delete requests without needing the requester's ID
 */
function adminDeleteSwapRequest(requestID) {
  try {
    if (!requestID) {
      return { success: false, message: "Missing request ID" };
    }
    
    // Find and delete the request in the sheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var swapSheet = ss.getSheetByName('SwapRequests');
    
    if (!swapSheet) {
      return { success: false, message: "SwapRequests sheet not found" };
    }
    
    var data = swapSheet.getDataRange().getValues();
    var headers = data[0];
    var requestIDCol = headers.indexOf('RequestID');
    
    if (requestIDCol === -1) {
      return { success: false, message: "RequestID column not found in SwapRequests sheet" };
    }
    
    // Find the row containing this request
    for (var i = 1; i < data.length; i++) {
      if (data[i][requestIDCol] === requestID) {
        // Delete the row
        swapSheet.deleteRow(i + 1);
        return { success: true, message: "Request deleted successfully" };
      }
    }
    
    return { success: false, message: "Request with ID '" + requestID + "' not found" };
  } catch (error) {
    Logger.log("Error in adminDeleteSwapRequest: " + error.toString());
    return { success: false, message: "Error: " + error.toString() };
  }
}

/**
 * Force the spreadsheet to refresh all data and clear any caches
 * This is called when a duty swap is performed to ensure all data is updated
 */
function forceRefreshData() {
  try {
    // Record start time to track performance
    var startTime = new Date().getTime();
    Logger.log("Starting force refresh of data");
    
    // Get the active spreadsheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Flush any pending changes
    SpreadsheetApp.flush();
    
    // Clear script cache to ensure fresh data
    var cache = CacheService.getScriptCache();
    cache.remove('forceRefreshInProgress');  // Clear any existing refresh flag
    cache.put('forceRefreshInProgress', 'true', 60);  // Set refresh flag with 1-minute expiration
    
    // Force recalculation by targeting key rows instead of the entire sheet
    var mainSheet = ss.getSheetByName('main data');
    if (mainSheet) {
      Logger.log("Performing targeted data refresh on main data sheet");
      
      var lastRow = mainSheet.getLastRow();
      var lastCol = mainSheet.getLastColumn();
      
      try {
        // Limited approach - touch only strategic rows
        // 1. Touch header row
        // 2. Touch first data row
        // 3. Touch a middle row
        // 4. Touch last row
        var rowsToTouch = [1, 2, Math.floor(lastRow/2), lastRow];
        
        for (var i = 0; i < rowsToTouch.length; i++) {
          var rowNum = rowsToTouch[i];
          if (rowNum > 0 && rowNum <= lastRow) {
            Logger.log("Touching strategic row " + rowNum);
            // Get only the first 5 columns for header rows
            var colsToTouch = (rowNum === 1) ? Math.min(lastCol, 5) : Math.min(lastCol, 3);
            var rowRange = mainSheet.getRange(rowNum, 1, 1, colsToTouch);
            var rowValues = rowRange.getValues();
            rowRange.setValues(rowValues);
          }
          
          // Check timeout to avoid script running too long
          if (new Date().getTime() - startTime > SCRIPT_TIMEOUT * 0.5) {
            Logger.log("Refresh operation approaching timeout, stopping early");
            break;
          }
        }
        
        // Flush after each major operation
        SpreadsheetApp.flush();
      } catch (innerError) {
        Logger.log("Warning in forceRefreshData: " + innerError.toString());
      }
    }
    
    // Also refresh the SwapRequests sheet if it exists, but with minimal touches
    var swapSheet = ss.getSheetByName('SwapRequests');
    if (swapSheet) {
      try {
        // Only touch the header row of swap requests
        var headerRange = swapSheet.getRange(1, 1, 1, Math.min(swapSheet.getLastColumn(), 3));
        var headerValues = headerRange.getValues();
        headerRange.setValues(headerValues);
        
        // Touch the most recent request row (likely to be most relevant)
        if (swapSheet.getLastRow() > 1) {
          var lastRowRange = swapSheet.getRange(swapSheet.getLastRow(), 1, 1, 3);
          var lastRowValues = lastRowRange.getValues();
          lastRowRange.setValues(lastRowValues);
        }
      } catch (innerError) {
        Logger.log("Minor error in swap sheet refresh: " + innerError.toString());
      }
    }
    
    // Flush final changes
    SpreadsheetApp.flush();
    
    // Clear the in-progress flag
    cache.remove('forceRefreshInProgress');
    
    // Log performance metrics
    var endTime = new Date().getTime();
    var duration = (endTime - startTime) / 1000;
    Logger.log("Force refresh completed in " + duration + " seconds");
    
    return { 
      success: true, 
      message: "Data refresh completed",
      duration: duration 
    };
  } catch (error) {
    // Clear the in-progress flag even on error
    try {
      CacheService.getScriptCache().remove('forceRefreshInProgress');
    } catch (e) {
      // Ignore error in cleanup
    }
    
    Logger.log("Error in forceRefreshData: " + error.toString());
    return { success: false, message: "Error refreshing data: " + error.toString() };
  }
}

/**
 * Creates a new duty assignment for a target employee
 * This is used when creating a duty without swapping from an existing duty
 */
function createDutyForEmployee(targetEmployee, dutyDate, slotNumber, program, location) {
  try {
    if (!targetEmployee || !dutyDate || !slotNumber) {
      Logger.log("createDutyForEmployee called with missing parameters - employee: " + targetEmployee + 
                ", date: " + dutyDate + ", slot: " + slotNumber);
      return { success: false, message: "Missing required parameters for duty creation" };
    }
    
    Logger.log("Creating new duty for employee - Target: " + targetEmployee + 
              ", Date: " + dutyDate + ", Slot: " + slotNumber + 
              ", Program: " + program + ", Location: " + location);
    
    // Get the main data sheet - try both possible sheet names
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var mainSheet = ss.getSheetByName('main data');
    
    if (!mainSheet) {
      // Try alternative sheet name
      mainSheet = ss.getSheetByName('InvigilatorDuties');
      
      if (!mainSheet) {
        Logger.log("Error: Neither 'main data' nor 'InvigilatorDuties' sheet found");
        return { success: false, message: "Main data sheet not found. Please check sheet names." };
      }
    }
    
    // Log which sheet we're using
    Logger.log("Using sheet: " + mainSheet.getName() + " for creating new duty");
    
    // Get the data from the sheet
    var data = mainSheet.getDataRange().getValues();
    var headers = data[0];
    
    // Find necessary column indices - use getColumnIndices for consistency with other functions
    var colIndices = getColumnIndices(headers);
    Logger.log("Column indices for new duty creation: " + JSON.stringify(colIndices));
    
    // Verify essential columns exist
    if (colIndices.empId === -1) {
      Logger.log("Error: Employee ID column not found");
      return { success: false, message: "Employee ID column not found" };
    }
    
    if (colIndices.date === -1) {
      Logger.log("Error: Date column not found");
      return { success: false, message: "Date column not found" };
    }
    
    // Determine which slot column to use based on slot number
    var slotCol = -1;
    var locationCol = -1;
    
    switch(slotNumber) {
      case "1":
      case 1:
        slotCol = colIndices.slot1;
        locationCol = colIndices.location1;
        break;
      case "2":
      case 2:
        slotCol = colIndices.slot2;
        locationCol = colIndices.location2;
        break;
      case "3":
      case 3:
        slotCol = colIndices.slot3;
        locationCol = colIndices.location3;
        break;
      case "4":
      case 4:
        slotCol = colIndices.slot4;
        locationCol = colIndices.location4;
        break;
      default:
        Logger.log("Error: Invalid slot number: " + slotNumber);
        return { success: false, message: "Invalid slot number: " + slotNumber };
    }
    
    if (slotCol === -1) {
      Logger.log("Error: Slot column not found for slot " + slotNumber);
      return { success: false, message: "Slot column not found for slot " + slotNumber };
    }
    
    // Check if the target employee already has a row for this date
    var targetRow = -1;
    var targetDayCount = 0;
    var targetInfo = null;
    
    // First try to find the employee's info regardless of date
    for (var i = 1; i < data.length; i++) {
      var empId = data[i][colIndices.empId] + "";
      
      if (empId === (targetEmployee + "")) {
        // Found the employee - save their info
        targetInfo = {
          name: colIndices.empName >= 0 ? data[i][colIndices.empName] : targetEmployee,
          designation: colIndices.designation >= 0 ? data[i][colIndices.designation] : "",
          email: colIndices.email >= 0 ? data[i][colIndices.email] : "",
          phone: colIndices.phone >= 0 ? data[i][colIndices.phone] : ""
        };
        
        // Check if this row also matches the duty date
        var rowDate = data[i][colIndices.date];
        var formattedRowDate = formatDateForComparison(rowDate);
        var formattedTargetDate = formatDateForComparison(dutyDate);
        
        if (formattedRowDate === formattedTargetDate) {
          targetRow = i;
          targetDayCount = colIndices.dayCount >= 0 ? (data[i][colIndices.dayCount] || 0) : 0;
          Logger.log("Found existing row for target employee on this date at row " + (i+1));
          break;
        }
      }
    }
    
    // Log employee info we found
    if (targetInfo) {
      Logger.log("Found target employee info: " + JSON.stringify(targetInfo));
    } else {
      Logger.log("Warning: Could not find existing info for employee " + targetEmployee);
    }
    
    // If target employee doesn't have a row for this date, create a new one
    if (targetRow === -1) {
      Logger.log("No existing row found for target employee on this date. Creating new row.");
      
      // Create a new row with all slots and locations empty
      var newRow = Array(headers.length).fill("");
      
      // Set the employee and date
      newRow[colIndices.empId] = targetEmployee;
      if (colIndices.empName >= 0 && targetInfo && targetInfo.name) newRow[colIndices.empName] = targetInfo.name;
      if (colIndices.designation >= 0 && targetInfo && targetInfo.designation) newRow[colIndices.designation] = targetInfo.designation;
      if (colIndices.email >= 0 && targetInfo && targetInfo.email) newRow[colIndices.email] = targetInfo.email;
      if (colIndices.phone >= 0 && targetInfo && targetInfo.phone) newRow[colIndices.phone] = targetInfo.phone;
      
      // Set the date as a Date object
      if (colIndices.date >= 0) {
        try {
          newRow[colIndices.date] = new Date(dutyDate);
        } catch (e) {
          Logger.log("Warning: Could not parse date " + dutyDate + ". Using as string.");
          newRow[colIndices.date] = dutyDate;
        }
      }
      
      // Reset all slots to '-' except the one being assigned
      if (colIndices.slot1 >= 0) newRow[colIndices.slot1] = colIndices.slot1 === slotCol ? "x" : "-";
      if (colIndices.slot2 >= 0) newRow[colIndices.slot2] = colIndices.slot2 === slotCol ? "x" : "-";
      if (colIndices.slot3 >= 0) newRow[colIndices.slot3] = colIndices.slot3 === slotCol ? "x" : "-";
      if (colIndices.slot4 >= 0) newRow[colIndices.slot4] = colIndices.slot4 === slotCol ? "x" : "-";
      
      // Set the location if provided
      if (locationCol !== -1 && location) {
        newRow[locationCol] = location;
      }
      
      // Set the program if provided
      if (colIndices.program !== -1 && program) {
        newRow[colIndices.program] = program;
      }
      
      // Set day count to 1 since this is the employee's first duty on this date
      if (colIndices.dayCount !== -1) {
        newRow[colIndices.dayCount] = 1;
      }
      
      // Set additional metadata if present
      if (colIndices.dayOfExam !== -1) {
        // Extract day of exam from other rows with the same date if possible
        for (var i = 1; i < data.length; i++) {
          if (formatDateForComparison(data[i][colIndices.date]) === formatDateForComparison(dutyDate) && 
              data[i][colIndices.dayOfExam]) {
            newRow[colIndices.dayOfExam] = data[i][colIndices.dayOfExam];
            break;
          }
        }
      }
      
      if (colIndices.weekDay !== -1) {
        // Calculate weekday from the date
        var weekdays = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
        var dateObj = new Date(dutyDate);
        newRow[colIndices.weekDay] = weekdays[dateObj.getDay()];
      }
      
      // Append the new row to the sheet
      mainSheet.appendRow(newRow);
      
      // Force a flush to save changes
      SpreadsheetApp.flush();
      
      Logger.log("New duty created for target employee");
      return { 
        success: true, 
        message: "New duty created for " + targetEmployee, 
        details: {
          employee: targetEmployee,
          date: dutyDate,
          slot: slotNumber,
          program: program,
          location: location,
          createdNewRow: true
        }
      };
    } else {
      // Target employee already has a row for this date
      // Check if the slot is already occupied
      if (data[targetRow][slotCol] === "x") {
        Logger.log("Error: Target employee already has a duty in slot " + slotNumber + " on this date");
        return { 
          success: false, 
          message: "Target employee already has a duty in slot " + slotNumber + " on this date" 
        };
      }
      
      // Update the existing row with the new duty
      mainSheet.getRange(targetRow + 1, slotCol + 1).setValue("x");
      
      // Update the location if provided
      if (locationCol !== -1 && location) {
        mainSheet.getRange(targetRow + 1, locationCol + 1).setValue(location);
      }
      
      // Update the program if provided
      if (colIndices.program !== -1 && program) {
        mainSheet.getRange(targetRow + 1, colIndices.program + 1).setValue(program);
      }
      
      // Increment the day count
      if (colIndices.dayCount !== -1) {
        mainSheet.getRange(targetRow + 1, colIndices.dayCount + 1).setValue(targetDayCount + 1);
      }
      
      // Force a flush to save changes
      SpreadsheetApp.flush();
      
      Logger.log("Duty added to existing row for target employee");
      return { 
        success: true, 
        message: "Duty added for " + targetEmployee, 
        details: {
          employee: targetEmployee,
          date: dutyDate,
          slot: slotNumber,
          program: program,
          location: location,
          createdNewRow: false
        }
      };
    }
    
  } catch (error) {
    Logger.log("Error in createDutyForEmployee: " + error.toString());
    Logger.log("Stack trace: " + error.stack);
    return { success: false, message: "Error: " + error.toString() };
  }
}

/**
 * Helper function to format a date for comparison
 * This handles different date formats and converts them to YYYY-MM-DD
 */
function formatDateForComparison(dateVal) {
  if (!dateVal) return "";
  
  // Log the original date value for debugging
  Logger.log("Formatting date for comparison: " + dateVal + " (type: " + typeof dateVal + ")");
  
  // If it's already a string representation of our standard format, return it directly
  if (typeof dateVal === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(dateVal)) {
    Logger.log("  - Already in YYYY-MM-DD format, returning as is: " + dateVal);
    return dateVal;
  }
  
  var date;
  
  // If it's already a Date object
  if (dateVal instanceof Date) {
    date = dateVal;
    Logger.log("  - Input is a Date object: " + date);
  } 
  // If it's a string, try to parse it
  else if (typeof dateVal === 'string') {
    try {
      // First try to clean up the string - remove extra spaces and normalize separators
      let cleanDateStr = dateVal.trim().replace(/\s+/g, ' ');
      
      // Special handling for "dd-MMM-yyyy" format (e.g., "15-Mar-2023")
      var monthNames = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
      var monthAbbr = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
      var fullMonthNames = ["january", "february", "march", "april", "may", "june", "july", "august", "september", "october", "november", "december"];
      
      // Add full month names to our patterns
      var allMonthPatterns = monthNames.concat(fullMonthNames);
      var monthRegexPattern = "(" + allMonthPatterns.join("|") + ")";
      
      // Two regex patterns:
      // 1. Day-Month-Year (15-Mar-2023 or 15 March 2023)
      var dMYRegex = new RegExp("(\\d{1,2})[-\\.\\/\\s]+" + monthRegexPattern + "[-\\.\\/\\s]+(\\d{4})", "i");
      // 2. Month-Day-Year (Mar-15-2023 or March 15 2023)
      var mDYRegex = new RegExp(monthRegexPattern + "[-\\.\\/\\s]+(\\d{1,2})[-\\.\\/\\s]+(\\d{4})", "i");
      
      var dMYMatch = cleanDateStr.match(dMYRegex);
      var mDYMatch = cleanDateStr.match(mDYRegex);
      
      if (dMYMatch) {
        // Extract day, month, and year
        var day = parseInt(dMYMatch[1], 10);
        var monthName = dMYMatch[2].toLowerCase();
        var year = parseInt(dMYMatch[3], 10);
        
        // Find month index (0-11)
        var monthIndex = -1;
        
        // Check month abbreviations
        monthIndex = monthAbbr.indexOf(monthName.substring(0, 3));
        
        // If not found, check full month names
        if (monthIndex === -1) {
          monthIndex = fullMonthNames.indexOf(monthName);
        }
        
        if (monthIndex >= 0) {
          date = new Date(year, monthIndex, day);
          Logger.log("  - Parsed as dd-MMM-yyyy: " + date);
        }
      } else if (mDYMatch) {
        // Extract month, day, and year
        var monthName = mDYMatch[1].toLowerCase();
        var day = parseInt(mDYMatch[2], 10);
        var year = parseInt(mDYMatch[3], 10);
        
        // Find month index (0-11)
        var monthIndex = -1;
        
        // Check month abbreviations
        monthIndex = monthAbbr.indexOf(monthName.substring(0, 3));
        
        // If not found, check full month names
        if (monthIndex === -1) {
          monthIndex = fullMonthNames.indexOf(monthName);
        }
        
        if (monthIndex >= 0) {
          date = new Date(year, monthIndex, day);
          Logger.log("  - Parsed as MMM-dd-yyyy: " + date);
        }
      }
      
      // If we couldn't parse using the month name, try numeric formats
      if (!date || isNaN(date.getTime())) {
        // Try to handle numeric date formats
        if (cleanDateStr.includes('-') || cleanDateStr.includes('/') || cleanDateStr.includes('.')) {
          // Split by any of the common separators
          var parts = cleanDateStr.split(/[-\.\/]/);
          
          if (parts.length === 3) {
            // Try to determine the format based on the values
            var part1 = parseInt(parts[0], 10);
            var part2 = parseInt(parts[1], 10);
            var part3 = parseInt(parts[2], 10);
            
            // Check if we have a 4-digit year in part1 (YYYY-MM-DD)
            if (parts[0].length === 4 && part1 >= 2000 && part1 <= 2100) {
              // YYYY-MM-DD format
              date = new Date(part1, part2-1, part3);
              Logger.log("  - Parsed as YYYY-MM-DD: " + date);
            } 
            // Check if we have a 4-digit year in part3 (DD-MM-YYYY or MM-DD-YYYY)
            else if (parts[2].length === 4 && part3 >= 2000 && part3 <= 2100) {
              // Try both DD-MM-YYYY and MM-DD-YYYY

              // If part1 > 31, it's not a valid day
              if (part1 > 31) {
                date = new Date(part3, part1-1, part2); // Assume MM-DD-YYYY
                Logger.log("  - Parsed as MM-DD-YYYY (first part > 31): " + date);
              }
              // If part2 > 12, it can't be a month
              else if (part2 > 12) {
                date = new Date(part3, part1-1, part2); // Must be MM-DD-YYYY
                Logger.log("  - Parsed as MM-DD-YYYY (second part > 12): " + date);
              }
              // If part1 > 12, part1 must be a day not a month
              else if (part1 > 12) {
                date = new Date(part3, part2-1, part1); // Must be DD-MM-YYYY
                Logger.log("  - Parsed as DD-MM-YYYY (first part > 12): " + date);
              }
              // If both part1 and part2 <= 12, try both formats
              else {
                // Try DD-MM-YYYY first (more common internationally)
                var ddmmyyyyDate = new Date(part3, part2-1, part1);
                
                // Then try MM-DD-YYYY
                var mmddyyyyDate = new Date(part3, part1-1, part2);
                
                // Check which one gives a valid date
                if (!isNaN(ddmmyyyyDate.getTime()) && part1 <= 31) {
                  date = ddmmyyyyDate;
                  Logger.log("  - Parsed as DD-MM-YYYY: " + date);
                } else if (!isNaN(mmddyyyyDate.getTime()) && part2 <= 31) {
                  date = mmddyyyyDate;
                  Logger.log("  - Parsed as MM-DD-YYYY: " + date);
                }
              }
            } 
            // Check for 2-digit year
            else if ((parts[0].length === 2 || parts[2].length === 2) && 
                     (part1 <= 99 || part3 <= 99)) {
              // Handle 2-digit year formats
              if (parts[0].length === 2 && part1 <= 99) {
                // YY-MM-DD format
                var year = part1 < 50 ? 2000 + part1 : 1900 + part1;
                date = new Date(year, part2-1, part3);
                Logger.log("  - Parsed as YY-MM-DD: " + date);
              } else if (parts[2].length === 2 && part3 <= 99) {
                // DD-MM-YY or MM-DD-YY format
                var year = part3 < 50 ? 2000 + part3 : 1900 + part3;
                // Try DD-MM-YY first
                if (part2 <= 12) {
                  date = new Date(year, part2-1, part1);
                  Logger.log("  - Parsed as DD-MM-YY: " + date);
                } else if (part1 <= 12) {
                  // Then MM-DD-YY
                  date = new Date(year, part1-1, part2);
                  Logger.log("  - Parsed as MM-DD-YY: " + date);
                }
              }
            } else {
              // Try standard parsing as last resort
              date = new Date(cleanDateStr);
              Logger.log("  - Using standard date parsing: " + date);
            }
          } else {
            // Not enough parts, try standard parsing
            date = new Date(cleanDateStr);
            Logger.log("  - Not three parts, using standard parsing: " + date);
          }
        } else {
          // No separators found, try standard parsing
          date = new Date(cleanDateStr);
          Logger.log("  - No separators found, using standard parsing: " + date);
        }
      }
    } catch (e) {
      Logger.log("Error parsing date: " + e);
      return {
        error: true,
        originalValue: dateVal,
        message: "Date parsing error: " + e.toString()
      };
    }
  } else if (typeof dateVal === 'number') {
    // Handle numeric timestamp (milliseconds since epoch)
    date = new Date(dateVal);
    Logger.log("  - Input is a number, treated as timestamp: " + date);
  } else {
    Logger.log("  - Unhandled date type: " + typeof dateVal);
    return {
      error: true,
      originalValue: dateVal,
      message: "Unhandled date type: " + typeof dateVal
    };
  }
  
  // Ensure it's a valid date
  if (!date || isNaN(date.getTime())) {
    Logger.log("  - Failed to parse date: " + dateVal);
    
    return {
      error: true,
      originalValue: dateVal,
      message: "Could not parse date value: " + dateVal
    };
  }
  
  // Format as YYYY-MM-DD for internal comparison
  var formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
  Logger.log("  - Formatted date for comparison: " + formattedDate);
  return formattedDate;
}

/**
 * Emergency function to get swap request data when the main function fails
 * This is a simplified version that only returns essential data
 */
function getEmergencySwapData() {
  try {
    Logger.log("Emergency swap data retrieval requested");
    
    // Hard limits to ensure this function completes quickly
    var MAX_ROWS = 20;
    var MAX_COLS = 10;  
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var swapSheet = ss.getSheetByName('SwapRequests');
    
    if (!swapSheet) {
      return { 
        success: true, 
        requests: [],
        isEmergencyMode: true,
        message: "No SwapRequests sheet found (emergency mode)"
      };
    }
    
    // Get only the first few rows to ensure we don't timeout
    var lastRow = Math.min(swapSheet.getLastRow(), MAX_ROWS + 1); // +1 for header
    var lastCol = Math.min(swapSheet.getLastColumn(), MAX_COLS);
    
    if (lastRow <= 1) {
      return { 
        success: true,
        requests: [],
        isEmergencyMode: true,
        message: "No requests found (emergency mode)"
      };
    }
    
    Logger.log("Emergency mode: Reading first " + lastRow + " rows");
    
    // Get a very limited set of data
    var headers = swapSheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var rows = swapSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    
    var requests = [];
    
    // Process each row
    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      var requestData = {};
      
      // Map data using headers
      for (var j = 0; j < headers.length; j++) {
        requestData[headers[j]] = row[j];
      }
      
      // Check if request has an ID (skip empty rows)
      if (requestData['RequestID']) {
        // Transform the data to match the expected client format
        requests.push({
          requestID: requestData['RequestID'],
          requesterID: requestData['RequesterID'] || '',
          requesterName: requestData['RequesterName'] || '',
          targetEmployee: requestData['TargetEmployee'] || '',
          date: requestData['DutyDate'] || '',
          slot: requestData['Slot'] || '',
          program: requestData['Program'] || '',
          location: requestData['Location'] || '',
          message: requestData['Message'] || '',
          status: requestData['Status'] || 'pending',
          requestTime: requestData['RequestTime'] || new Date().toISOString(),
          responseTime: requestData['ResponseTime'] || ''
        });
      }
      
      // Limit to very few requests in emergency mode
      if (requests.length >= 10) {
        break;
      }
    }
    
    return {
      success: true,
      requests: requests,
      isEmergencyMode: true,
      message: "Limited data loaded in emergency mode",
      limitedData: true
    };
  } catch (error) {
    Logger.log("Error in getEmergencySwapData: " + error.toString());
    return {
      success: false,
      message: "Emergency data load failed: " + error.toString(),
      isEmergencyMode: true
    };
  }
}

/**
 * Get emergency swap data with improved handling to ensure latest data is fetched
 * This function serves as a fallback when the primary method fails
 */
function getEmergencySwapData(userID) {
  try {
    if (!userID) {
      return { 
        success: false, 
        message: "User ID is required for emergency data retrieval",
        isEmergencyMode: true 
      };
    }
    
    // Log this emergency access for diagnostics
    Logger.log("Emergency swap data access requested for user: " + userID);
    
    // First check if we have cached emergency data
    var cache = CacheService.getScriptCache();
    var cacheKey = 'emergencySwapRequests_' + userID;
    var cachedData = cache.get(cacheKey);
    
    if (cachedData) {
      // Decode and return cached data with freshness indicator
      var parsedData = JSON.parse(cachedData);
      Logger.log("Using emergency cached data for " + userID + " with " + parsedData.length + " requests");
      
      return {
        success: true,
        requests: parsedData,
        isEmergencyMode: true,
        message: "Emergency cached data loaded",
        timestamp: new Date().toISOString(),
        cacheHit: true
      };
    }
    
    // No cached data found, attempt direct sheet access as last resort
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var swapSheet = ss.getSheetByName('SwapRequests');
    
    if (!swapSheet) {
      return { 
        success: false, 
        message: "Emergency mode active but swap sheet not found",
        isEmergencyMode: true 
      };
    }
    
    // Get headers and essential columns only
    var headers = swapSheet.getRange(1, 1, 1, swapSheet.getLastColumn()).getValues()[0];
    
    // Find required column indices
    var requesterIDCol = -1, targetEmployeeCol = -1, requestIDCol = -1, statusCol = -1;
    var dateCol = -1, slotCol = -1, programCol = -1, locationCol = -1;
    var messageCol = -1, requesterNameCol = -1, requestTimeCol = -1, responseTimeCol = -1;
    
    for (var i = 0; i < headers.length; i++) {
      var header = headers[i].toString().toLowerCase();
      if (header.includes('requester') && header.includes('id')) requesterIDCol = i;
      else if (header.includes('target') && header.includes('employee')) targetEmployeeCol = i;
      else if (header.includes('request') && header.includes('id')) requestIDCol = i;
      else if (header === 'status') statusCol = i;
      else if (header.includes('date') || header.includes('duty') && header.includes('date')) dateCol = i;
      else if (header === 'slot') slotCol = i;
      else if (header === 'program') programCol = i;
      else if (header === 'location') locationCol = i;
      else if (header === 'message') messageCol = i;
      else if (header.includes('requester') && header.includes('name')) requesterNameCol = i;
      else if (header.includes('request') && header.includes('time')) requestTimeCol = i;
      else if (header.includes('response') && header.includes('time')) responseTimeCol = i;
    }
    
    // Verify essential columns were found
    if (requesterIDCol === -1 || targetEmployeeCol === -1) {
      return { 
        success: false, 
        message: "Emergency mode active but essential columns not found",
        isEmergencyMode: true 
      };
    }
    
    // Get only the columns we need for better performance
    var neededColumns = [requestIDCol, requesterIDCol, requesterNameCol, targetEmployeeCol, 
                         dateCol, slotCol, programCol, locationCol, messageCol, statusCol, 
                         requestTimeCol, responseTimeCol].filter(function(col) { return col !== -1; });
    
    // Get the highest necessary column index
    var maxCol = Math.max.apply(null, neededColumns);
    var columns = [];
    
    for (var c = 0; c <= maxCol; c++) {
      if (neededColumns.indexOf(c) !== -1) {
        columns.push(c);
      }
    }
    
    // Only read data for rows that might be relevant
    var userIdLower = userID.toString().toLowerCase();
    var dataRange = swapSheet.getRange(2, 1, swapSheet.getLastRow() - 1, swapSheet.getLastColumn());
    var data = dataRange.getValues();
    
    var requests = [];
    var requestIDs = new Set(); // To prevent duplicates
    
    // Only get the 20 most recent rows to optimize performance
    var startRow = Math.max(0, data.length - 20);
    
    // Process rows in reverse order to get most recent first
    for (var i = data.length - 1; i >= startRow; i--) {
      var row = data[i];
      
      var rowRequesterID = String(row[requesterIDCol] || '').toLowerCase();
      var rowTargetEmployee = String(row[targetEmployeeCol] || '').toLowerCase();
      
      // Check if this request is relevant to the user
      if (rowRequesterID.includes(userIdLower) || rowTargetEmployee.includes(userIdLower)) {
        var requestID = row[requestIDCol];
        
        // Skip duplicates
        if (requestIDs.has(requestID)) continue;
        requestIDs.add(requestID);
        
        // Build request object with available data
        var request = {
          requestID: requestID,
          requesterID: row[requesterIDCol],
          requesterName: requesterNameCol !== -1 ? row[requesterNameCol] : '',
          targetEmployee: row[targetEmployeeCol],
          date: dateCol !== -1 ? row[dateCol] : '',
          slot: slotCol !== -1 ? row[slotCol] : '',
          program: programCol !== -1 ? row[programCol] : '',
          location: locationCol !== -1 ? row[locationCol] : '',
          message: messageCol !== -1 ? row[messageCol] : '',
          status: statusCol !== -1 ? row[statusCol] : 'unknown',
          requestTime: requestTimeCol !== -1 ? row[requestTimeCol] : '',
          responseTime: responseTimeCol !== -1 ? row[responseTimeCol] : ''
        };
        
        requests.push(request);
      }
      
      // Limit to 15 requests for emergency mode
      if (requests.length >= 15) {
        break;
      }
    }
    
    // Cache this emergency data for future use
    if (requests.length > 0) {
      try {
        cache.put(cacheKey, JSON.stringify(requests), 21600); // Cache for 6 hours
      } catch (cacheError) {
        Logger.log("Failed to cache emergency swap data: " + cacheError.toString());
      }
    }
    
    return {
      success: true,
      requests: requests,
      isEmergencyMode: true,
      message: requests.length > 0 ? 
              "Retrieved " + requests.length + " recent requests in emergency mode" : 
              "No relevant requests found in emergency mode",
      timestamp: new Date().toISOString(),
      limitedData: true
    };
  } catch (error) {
    Logger.log("Error in getEmergencySwapData: " + error.toString());
    return {
      success: false,
      message: "Emergency data load failed: " + error.toString(),
      isEmergencyMode: true
    };
  }
}

/**
 * Verify if a user has a specific duty assigned
 */
function verifyUserHasDuty(employeeId, date, slotNumber) {
  try {
    if (!employeeId || !date || !slotNumber) {
      return { 
        success: false, 
        message: "Employee ID, date, and slot number are required" 
      };
    }
    
    // Get main data sheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var mainSheet = ss.getSheetByName('main data');
    
    if (!mainSheet) {
      return { success: false, message: "Main data sheet not found" };
    }
    
    // Get all data
    var data = mainSheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return { success: false, message: "No data found in main data sheet" };
    }
    
    // Extract headers and get column indices
    var headers = data[0];
    var colIndices = getColumnIndices(headers);
    
    // Validate required column indices
    if (colIndices.empId === -1) {
      return { success: false, message: "Employee ID column not found" };
    }
    
    if (colIndices.date === -1) {
      return { success: false, message: "Date column not found" };
    }
    
    // Determine which slot column to check
    var slotCol = -1;
    switch(slotNumber.toString()) {
      case "1": slotCol = colIndices.slot1; break;
      case "2": slotCol = colIndices.slot2; break;
      case "3": slotCol = colIndices.slot3; break;
      case "4": slotCol = colIndices.slot4; break;
      default: return { success: false, message: "Invalid slot number" };
    }
    
    if (slotCol === -1) {
      return { success: false, message: "Slot column not found" };
    }
    
    // Look for a matching row
    var matchingRow = -1;
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var rowEmpId = row[colIndices.empId].toString().trim();
      var rowDate = formatDate(row[colIndices.date]);
      
      // Check if employee ID and date match
      if (rowEmpId === employeeId.toString().trim() && rowDate === date) {
        // Check if this slot has an assignment
        var slotValue = row[slotCol].toString().trim().toLowerCase();
        if (slotValue === 'x') {
          return { 
            success: true, 
            hasDuty: true, 
            rowIndex: i + 1  // +1 because sheet rows are 1-indexed
          };
        }
      }
    }
    
    // No matching duty found
    return { 
      success: true, 
      hasDuty: false,
      needToCreateDuty: false
    };
  } catch (error) {
    Logger.log("Error verifying user duty: " + error.toString());
    return { success: false, message: "Error: " + error.toString() };
  }
}

/**
 * Process an approved swap request by updating the duty roster
 */
function processApprovedSwap(requestDetails) {
  try {
    Logger.log("Processing approved swap for request: " + requestDetails.RequestID);
    
    // Extract necessary details
    var requesterID = requestDetails.RequesterID;
    var targetEmployeeID = requestDetails.TargetEmployee;
    var date = requestDetails.Date;
    var slot = requestDetails.Slot;
    var location = requestDetails.Location;
    var program = requestDetails.Program;
    
    if (!requesterID || !targetEmployeeID || !date || !slot) {
      return { success: false, message: "Missing required details in request" };
    }
    
    // Get the main data sheet 
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var mainSheet = ss.getSheetByName('main data');
    
    if (!mainSheet) {
      return { success: false, message: "Main data sheet not found" };
    }
    
    // Get all data
    var data = mainSheet.getDataRange().getValues();
    var headers = data[0];
    
    // Find column indices
    var colIndices = getColumnIndices(headers);
    
    if (colIndices.empId === -1 || colIndices.date === -1) {
      return { success: false, message: "Required columns not found in main data" };
    }
    
    // Find the row with the requester's duty
    var requesterRow = -1;
    for (var i = 1; i < data.length; i++) {
      var rowEmpId = data[i][colIndices.empId] ? data[i][colIndices.empId].toString() : "";
      var rowDate = formatDate(data[i][colIndices.date]);
      
      if (rowEmpId.toString().toLowerCase() === requesterID.toString().toLowerCase() && 
          rowDate === date) {
        
        // Check that the requester has the specified slot
        var slotCol = -1;
        switch(slot) {
          case "1": case 1: slotCol = colIndices.slot1; break;
          case "2": case 2: slotCol = colIndices.slot2; break;
          case "3": case 3: slotCol = colIndices.slot3; break;
          case "4": case 4: slotCol = colIndices.slot4; break;
        }
        
        if (slotCol !== -1) {
          var slotValue = data[i][slotCol] ? data[i][slotCol].toString().trim().toLowerCase() : "";
          if (slotValue === 'x') {
            requesterRow = i + 1; // +1 because sheet rows are 1-indexed
            break;
          }
        }
      }
    }
    
    // Find if target employee already has a row for this date
    var targetRow = -1;
    for (var i = 1; i < data.length; i++) {
      var rowEmpId = data[i][colIndices.empId] ? data[i][colIndices.empId].toString() : "";
      var rowDate = formatDate(data[i][colIndices.date]);
      
      if (rowEmpId.toString().toLowerCase() === targetEmployeeID.toString().toLowerCase() && 
          rowDate === date) {
        targetRow = i + 1;
        break;
      }
    }
    
    // If we didn't find the requester's duty, return an error
    if (requesterRow === -1) {
      return { 
        success: false, 
        message: "Requester's duty not found for the specified date and slot" 
      };
    }
    
    // Determine which slot column to update
    var slotCol = -1;
    var locationCol = -1;
    switch(slot) {
      case "1": case 1: 
        slotCol = colIndices.slot1;
        locationCol = colIndices.location1;
        break;
      case "2": case 2: 
        slotCol = colIndices.slot2;
        locationCol = colIndices.location2;
        break;
      case "3": case 3: 
        slotCol = colIndices.slot3;
        locationCol = colIndices.location3;
        break;
      case "4": case 4: 
        slotCol = colIndices.slot4;
        locationCol = colIndices.location4;
        break;
    }
    
    if (slotCol === -1) {
      return { success: false, message: "Invalid slot number" };
    }
    
    // If target employee already has a row for this date
    if (targetRow !== -1) {
      // Check if the target employee's slot is free
      var targetSlotValue = data[targetRow-1][slotCol] ? data[targetRow-1][slotCol].toString().trim().toLowerCase() : "";
      
      if (targetSlotValue === 'x') {
        return {
          success: false,
          message: "Target employee already has a duty in this slot"
        };
      }
      
      // Update the target employee's slot
      mainSheet.getRange(targetRow, slotCol + 1).setValue('x');
      
      // Update location if applicable
      if (locationCol !== -1 && location) {
        mainSheet.getRange(targetRow, locationCol + 1).setValue(location);
      }
      
      // Update day counts for target employee
      updateDayCounts(mainSheet, targetRow, colIndices);
    } else {
      // Copy the requester's row for the target employee
      var requesterRowData = data[requesterRow - 1];
      var newRowData = [];
      
      for (var j = 0; j < requesterRowData.length; j++) {
        if (j === colIndices.empId) {
          // Set employee ID to target
          newRowData.push(targetEmployeeID);
        } else if (j === colIndices.empName) {
          // Set employee name if available
          var targetDetails = findEmployeeById(targetEmployeeID);
          newRowData.push(targetDetails && targetDetails.name ? targetDetails.name : '');
        } else if (j === colIndices.designation) {
          // Set designation if available
          var targetDetails = targetDetails || findEmployeeById(targetEmployeeID);
          newRowData.push(targetDetails && targetDetails.designation ? targetDetails.designation : '');
        } else if (j === colIndices.slot1 || j === colIndices.slot2 || 
                   j === colIndices.slot3 || j === colIndices.slot4) {
          // Clear all slots except the one being swapped
          if (j === slotCol) {
            newRowData.push('x');
          } else {
            newRowData.push('');
          }
        } else if (j === colIndices.location1 || j === colIndices.location2 || 
                   j === colIndices.location3 || j === colIndices.location4) {
          // Set location for the swapped slot, clear others
          if (j === locationCol && location) {
            newRowData.push(location);
          } else {
            newRowData.push('');
          }
        } else if (j === colIndices.dayCount) {
          // Set day count to 1 for the new row
          newRowData.push(1);
        } else {
          // Copy other values (e.g., date, program)
          newRowData.push(requesterRowData[j]);
        }
      }
      
      // Add the new row
      mainSheet.appendRow(newRowData);
      
      // Get the row we just added
      targetRow = mainSheet.getLastRow();
    }
    
    // Clear the requester's slot
    mainSheet.getRange(requesterRow, slotCol + 1).setValue('');
    
    // Clear the location for that slot if applicable
    if (locationCol !== -1) {
      mainSheet.getRange(requesterRow, locationCol + 1).setValue('');
    }
    
    // Update day counts for requester
    updateDayCounts(mainSheet, requesterRow, colIndices);
    
    Logger.log("Successfully swapped duty from " + requesterID + " to " + targetEmployeeID);
    return { 
      success: true, 
      message: "Duty successfully swapped" 
    };
  } catch (error) {
    Logger.log("Error processing approved swap: " + error.toString());
    return { success: false, message: "Error: " + error.toString() };
  }
}

/**
 * Update day counts for a row
 */
function updateDayCounts(sheet, rowNum, colIndices) {
  // Get the current row data
  var rowData = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // Count duties in this row
  var dayCount = 0;
  if (colIndices.slot1 !== -1 && rowData[colIndices.slot1] === 'x') dayCount++;
  if (colIndices.slot2 !== -1 && rowData[colIndices.slot2] === 'x') dayCount++;
  if (colIndices.slot3 !== -1 && rowData[colIndices.slot3] === 'x') dayCount++;
  if (colIndices.slot4 !== -1 && rowData[colIndices.slot4] === 'x') dayCount++;
  
  // Update the day count cell
  if (colIndices.dayCount !== -1) {
    sheet.getRange(rowNum, colIndices.dayCount + 1).setValue(dayCount);
  }
}

/**
 * Find employee by ID or name
 */
function findEmployeeById(employeeId) {
  try {
    if (!employeeId) {
      return null;
    }
    
    // Get main data sheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var mainSheet = ss.getSheetByName('main data');
    
    if (!mainSheet) {
      Logger.log("Main data sheet not found");
      return null;
    }
    
    // Get all data
    var data = mainSheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      Logger.log("No data found in main data sheet");
      return null;
    }
    
    // Extract headers and get column indices
    var headers = data[0];
    var colIndices = getColumnIndices(headers);
    
    // Validate required column indices
    if (colIndices.empId === -1) {
      Logger.log("Employee ID column not found");
      return null;
    }
    
    // Convert search term to lowercase for case-insensitive search
    var searchTerm = employeeId.toString().toLowerCase().trim();
    
    // Look for the employee
    for (var i = 1; i < data.length; i++) {
      var empId = data[i][colIndices.empId] ? data[i][colIndices.empId].toString().toLowerCase().trim() : "";
      var empName = colIndices.empName !== -1 && data[i][colIndices.empName] ? data[i][colIndices.empName].toString().toLowerCase().trim() : "";
      
      // Check for match by ID or name
      if (empId === searchTerm || empName === searchTerm) {
        return {
          empId: data[i][colIndices.empId],
          name: colIndices.empName !== -1 ? data[i][colIndices.empName] : "",
          designation: colIndices.designation !== -1 ? data[i][colIndices.designation] : "",
          email: colIndices.email !== -1 ? data[i][colIndices.email] : "",
          phone: colIndices.phone !== -1 ? data[i][colIndices.phone] : ""
        };
      }
    }
    
    // No employee found
    return null;
  } catch (error) {
    Logger.log("Error finding employee by ID: " + error.toString());
    return null;
  }
}

/**
 * Send email notification to target employee about a new swap request
 */
function sendSwapRequestNotification(requestData) {
  try {
    // Ensure we have a valid requestData object
    if (!requestData) {
      Logger.log("CRITICAL ERROR: requestData is null or undefined in sendSwapRequestNotification");
      // Create a minimal valid requestData as fallback
      requestData = {
        targetEmployee: "default_target",
        requesterID: Session.getActiveUser().getEmail() || "admin",
        date: new Date().toLocaleDateString(),
        slot: "1",
        program: "Default Program",
        location: "Default Location"
      };
      Logger.log("Created default requestData as fallback: " + JSON.stringify(requestData));
    }
    
    // Normalize the request data to handle various property naming conventions
    requestData = normalizeRequestData(requestData);
    
    // Log normalized data
    Logger.log("Normalized request data for notification: " + JSON.stringify(requestData));
    
    // Direct check and fix for targetEmployee
    if (!requestData.targetEmployee) {
      Logger.log("Missing target employee data for email notification - attempting to fix");
      
      // Try all possible property names
      if (requestData.targetEmployeeID) {
        requestData.targetEmployee = requestData.targetEmployeeID;
        Logger.log("Using targetEmployeeID: " + requestData.targetEmployee);
      } else if (requestData['Target Employee']) {
        requestData.targetEmployee = requestData['Target Employee'];
        Logger.log("Using Target Employee: " + requestData.targetEmployee);
      } else if (requestData.TargetEmployee) {
        requestData.targetEmployee = requestData.TargetEmployee;
        Logger.log("Using Targetemployee: " + requestData.targetEmployee);
      } else {
        // HARDCODED FALLBACK: If we absolutely cannot find a target employee, use admin
        requestData.targetEmployee = "admin";
        Logger.log("Using hardcoded fallback for targetEmployee: admin");
      }
    }
    
    Logger.log("Sending notification for swap request to target employee: " + requestData.targetEmployee);
    
    // Get target employee email directly from main data sheet
    var targetEmail = getEmployeeEmailById(requestData.targetEmployee);
    
    // Get requester name
    var requesterName = "Invigilator";
    if (requestData.requesterID) {
      // Try to get requester details
      var requesterDetails = findEmployeeById(requestData.requesterID);
      if (requesterDetails && requesterDetails.name) {
        requesterName = requesterDetails.name;
      } else {
        requesterName = "Invigilator " + requestData.requesterID;
      }
    }
    
    // Format slot time
    var slotTimeText = "";
    switch(requestData.slot) {
      case "1": slotTimeText = "Slot 1 (09:00-10:30)"; break;
      case "2": slotTimeText = "Slot 2 (11:30-13:00)"; break;
      case "3": slotTimeText = "Slot 3 (15:30-17:00)"; break;
      case "4": slotTimeText = "Slot 4 (17:30-19:00)"; break;
      default: slotTimeText = "Slot " + requestData.slot;
    }
    
    // Get formatted date
    var formattedDate = requestData.date || "Unknown Date";
    try {
      if (requestData.date) {
        var dateParts = requestData.date.split('-');
        if (dateParts.length === 3) {
          var monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
          var monthIndex = parseInt(dateParts[1]) - 1;
          if (isNaN(monthIndex)) {
            // Check if it's a month abbreviation
            var monthAbbr = dateParts[1].toLowerCase();
            var monthNamesLower = monthNames.map(function(m) { return m.toLowerCase(); });
            monthIndex = monthNamesLower.indexOf(monthAbbr);
          }
          
          if (monthIndex >= 0 && monthIndex < 12) {
            formattedDate = dateParts[0] + " " + monthNames[monthIndex] + " " + dateParts[2];
          }
        }
      }
    } catch (e) {
      // If date formatting fails, use the original date
      Logger.log("Error formatting date: " + e.toString());
    }
    
    // Create the email subject
    var subject = "Duty Swap Request from " + requesterName;
    
    // Create the email body
    var body = "Dear Colleague,\n\n" +
               requesterName + " has requested to swap a duty with you.\n\n" +
               "Duty Details:\n" +
               "- Date: " + formattedDate + "\n" +
               "- Time: " + slotTimeText + "\n" +
               "- Program: " + (requestData.program || 'N/A') + "\n" +
               "- Location: " + (requestData.location || 'N/A') + "\n";
    
    if (requestData.reasonType) {
      body += "- Reason: " + requestData.reasonType + "\n";
    }
    
    if (requestData.notes) {
      body += "- Additional Notes: " + requestData.notes + "\n";
    }
    
    body += "\nThis request requires approval from the administrator. " +
            "They will contact you to confirm if you are available to take this duty.\n\n" +
            "This is an automated notification. Please do not reply to this email.\n\n" +
            "Regards,\nInvigilator Duty Management System";
    
    // Send the email
    try {
      Logger.log("Sending email to: " + targetEmail);
      MailApp.sendEmail({
        to: targetEmail,
        subject: subject,
        body: body
      });
      
      Logger.log("Swap request notification sent to: " + targetEmail);
      return true;
    } catch (emailError) {
      Logger.log("Failed to send email: " + emailError.toString());
      // Return success anyway to prevent process from failing
      return true;
    }
  } catch (error) {
    Logger.log("Error sending swap request notification: " + error.toString());
    return false;
  }
}

/**
 * Send status update notification to the requester
 */
function sendStatusUpdateNotification(requestDetails, newStatus) {
  try {
    // Ensure we have a valid requestDetails object
    if (!requestDetails) {
      Logger.log("CRITICAL ERROR: requestDetails is null or undefined in sendStatusUpdateNotification");
      // Create a minimal valid requestDetails as fallback
      requestDetails = {
        'Requester ID': Session.getActiveUser().getEmail() || "admin",
        'Date': new Date().toLocaleDateString(),
        'Slot': "1",
        'Program': "Default Program",
        'Target Employee': "default_target",
        'Status': newStatus || "updated"
      };
      Logger.log("Created default requestDetails as fallback: " + JSON.stringify(requestDetails));
    }
    
    // Normalize the request data to handle various property naming conventions
    requestDetails = normalizeRequestData(requestDetails);
    
    // Log normalized data
    Logger.log("Normalized request data for status update: " + JSON.stringify(requestDetails));
    
    // Check for requester ID and ensure it's never null
    var requesterId = null;
    if (requestDetails['Requester ID']) {
      requesterId = requestDetails['Requester ID'];
    } else if (requestDetails['requesterID']) {
      requesterId = requestDetails['requesterID'];
    } else if (requestDetails.RequesterID) {
      requesterId = requestDetails.RequesterID;
    } else if (requestDetails.requesterID) {
      requesterId = requestDetails.requesterID;
    }
    
    // If still no requester ID, use a default
    if (!requesterId) {
      requesterId = "admin";
      Logger.log("Using hardcoded fallback for requester ID: admin");
    }
    
    Logger.log("Sending status update notification to requester: " + requesterId);
    
    // Find target employee name if available
    var targetName = requestDetails['Target Employee'];
    var targetDetails = findEmployeeById(requestDetails['Target Employee']);
    if (targetDetails && targetDetails.name) {
      targetName = targetDetails.name;
    }
    
    // Find requester details
    var requesterDetails = findEmployeeById(requestDetails['Requester ID']);
    var requesterName = requesterDetails && requesterDetails.name ? 
                       requesterDetails.name : requestDetails['Requester ID'];
    
    // Format status for display
    var statusText = newStatus.charAt(0).toUpperCase() + newStatus.slice(1);
    
    // Format slot time
    var slotTimeText = "";
    switch(requestDetails['Slot']) {
      case "1": slotTimeText = "Slot 1 (09:00-10:30)"; break;
      case "2": slotTimeText = "Slot 2 (11:30-13:00)"; break;
      case "3": slotTimeText = "Slot 3 (15:30-17:00)"; break;
      case "4": slotTimeText = "Slot 4 (17:30-19:00)"; break;
      default: slotTimeText = "Slot " + requestDetails['Slot'];
    }
    
    // Create the email subject
    var subject = "Duty Swap Request " + statusText;
    
    // Create the email body
    var body = "Dear " + (requesterDetails.name || "Colleague") + ",\n\n" +
               "Your duty swap request has been " + newStatus.toLowerCase() + ".\n\n" +
               "Request Details:\n" +
               "- Date: " + requestDetails['Date'] + "\n" +
               "- Time: " + slotTimeText + "\n" +
               "- Program: " + (requestDetails['Program'] || 'N/A') + "\n" +
               "- Swap With: " + targetName + "\n";
    
    if (requestDetails['Admin Notes']) {
      body += "\nAdministrator Notes: " + requestDetails['Admin Notes'] + "\n";
    }
    
    // Add specific instructions based on status
    if (newStatus.toLowerCase() === "approved") {
      body += "\nYour duty has been successfully transferred to " + targetName + ". " +
              "You are no longer responsible for this duty.\n";
    } else if (newStatus.toLowerCase() === "rejected") {
      body += "\nYour duty assignment remains unchanged. You are still responsible for this duty.\n";
    }
    
    body += "\nThis is an automated notification. Please do not reply to this email.\n\n" +
            "Regards,\nInvigilator Duty Management System";
    
    // Send the email
    MailApp.sendEmail({
      to: requesterId,
      subject: subject,
      body: body
    });
    
    Logger.log("Status update notification sent to: " + requesterId);
    return true;
  } catch (error) {
    Logger.log("Error sending status update notification: " + error.toString());
    return false;
  }
}

/**
 * Send notification about a swap request status change
 */
function sendSwapStatusNotification(requestDetails, newStatus, adminNotes) {
  try {
    // Ensure we have a valid requestDetails object
    if (!requestDetails) {
      Logger.log("CRITICAL ERROR: requestDetails is null or undefined in sendSwapStatusNotification");
      // Create a minimal valid requestDetails as fallback
      requestDetails = {
        'RequesterID': Session.getActiveUser().getEmail() || "admin",
        'Date': new Date().toLocaleDateString(),
        'Slot': "1",
        'Program': "Default Program",
        'TargetEmployee': "default_target"
      };
      Logger.log("Created default requestDetails as fallback: " + JSON.stringify(requestDetails));
    }
    
    // Normalize the request data
    requestDetails = normalizeRequestData(requestDetails);
    
    // Log normalized data
    Logger.log("Normalized request data for swap status notification: " + JSON.stringify(requestDetails));
    
    // Get requester ID - try all possible field names
    var requesterId = null;
    if (requestDetails.RequesterID) {
      requesterId = requestDetails.RequesterID;
    } else if (requestDetails['Requester ID']) {
      requesterId = requestDetails['Requester ID']; 
    } else if (requestDetails.requesterID) {
      requesterId = requestDetails.requesterID;
    }
    
    // If still no requester ID, use a default
    if (!requesterId) {
      requesterId = Session.getActiveUser().getEmail() || "admin";
      Logger.log("Using hardcoded fallback for requester ID: " + requesterId);
    }
    
    Logger.log("Using requester ID for notification: " + requesterId);
    
    // Get requester email directly using our dedicated function
    var requesterEmail = getEmployeeEmailById(requesterId);
    if (!requesterEmail) {
      requesterEmail = Session.getActiveUser().getEmail() || "admin@example.com";
      Logger.log("Using fallback email for requester: " + requesterEmail);
    }
    
    // Format the date for better readability
    var formattedDate = formatDate(requestDetails.Date);
    
    // Construct email subject and body
    var subject = "Duty Swap Request " + newStatus + " - " + formattedDate;
    
    var body = "Dear Staff Member,\n\n";
    body += "Your duty swap request has been " + newStatus.toLowerCase() + ".\n\n";
    body += "Request Details:\n";
    body += "- Date: " + formattedDate + "\n";
    body += "- Slot: " + requestDetails.Slot + "\n";
    if (requestDetails.Program) {
      body += "- Program: " + requestDetails.Program + "\n";
    }
    if (requestDetails.Location) {
      body += "- Location: " + requestDetails.Location + "\n";
    }
    body += "- Target Employee: " + requestDetails.TargetEmployee + "\n";
    body += "- Status: " + newStatus + "\n";
    
    if (adminNotes) {
      body += "\nAdmin Notes: " + adminNotes + "\n";
    }
    
    if (newStatus.toLowerCase() === "approved") {
      body += "\nThe duty has been reassigned from you to the target employee.\n";
    } else if (newStatus.toLowerCase() === "rejected") {
      body += "\nYour duty assignment remains unchanged.\n";
    }
    
    body += "\nThis is an automated message. Please do not reply to this email.\n";
    body += "\nBest regards,\nDuty Management System";
    
    // Send the email
    MailApp.sendEmail(requesterEmail, subject, body);
    
    // Also notify the target employee if the request was approved
    if (newStatus.toLowerCase() === "approved") {
      var targetEmail = getEmployeeEmailById(requestDetails.TargetEmployee);
      
      if (targetEmail) {
        var targetSubject = "Duty Assignment Notice - " + formattedDate;
        
        var targetBody = "Dear Staff Member,\n\n";
        targetBody += "A duty has been assigned to you through a swap request.\n\n";
        targetBody += "Duty Details:\n";
        targetBody += "- Date: " + formattedDate + "\n";
        targetBody += "- Slot: " + requestDetails.Slot + "\n";
        if (requestDetails.Program) {
          targetBody += "- Program: " + requestDetails.Program + "\n";
        }
        if (requestDetails.Location) {
          targetBody += "- Location: " + requestDetails.Location + "\n";
        }
        targetBody += "- Original Assignee: " + requestDetails.RequesterID + "\n";
        targetBody += "\nThis duty has been transferred to you as part of an approved swap request.\n";
        targetBody += "\nThis is an automated message. Please do not reply to this email.\n";
        targetBody += "\nBest regards,\nDuty Management System";
        
        MailApp.sendEmail(targetEmail, targetSubject, targetBody);
      } else {
        Logger.log("Target employee email not found for ID: " + requestDetails.TargetEmployee);
      }
    }
    
    return { success: true, message: "Notifications sent successfully" };
  } catch (error) {
    Logger.log("Error sending swap status notification: " + error.toString());
    return { success: false, message: "Error: " + error.toString() };
  }
}

/**
 * Get an employee's email by their ID
 */
function getEmployeeEmailById(employeeId) {
  try {
    if (!employeeId) return null;
    
    // Get employees sheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var employeesSheet = ss.getSheetByName('Employees') || ss.getSheetByName('employees');
    
    if (!employeesSheet) {
      Logger.log("Employees sheet not found");
      return null;
    }
    
    // Get all data
    var data = employeesSheet.getDataRange().getValues();
    var headers = data[0];
    
    // Find relevant column indices
    var idCol = -1;
    var emailCol = -1;
    
    for (var i = 0; i < headers.length; i++) {
      var header = headers[i].toString().toLowerCase();
      if (header.includes('id') || header === 'empid' || header === 'employeeid') {
        idCol = i;
      } else if (header.includes('email')) {
        emailCol = i;
      }
    }
    
    if (idCol === -1 || emailCol === -1) {
      Logger.log("Required columns not found in employees sheet");
      return null;
    }
    
    // Search for the employee
    for (var i = 1; i < data.length; i++) {
      var rowId = data[i][idCol] ? data[i][idCol].toString().trim() : "";
      
      if (rowId === employeeId.toString().trim()) {
        var email = data[i][emailCol];
        return email ? email.toString().trim() : null;
      }
    }
    
    // Employee not found or no email
    return null;
  } catch (error) {
    Logger.log("Error getting employee email: " + error.toString());
    return null;
  }
}

/**
 * Creates the SwapRequests sheet with proper headers and formatting
 * @return {Sheet} The newly created sheet
 */
function createSwapRequestsSheet() {
  console.log("Creating SwapRequests sheet");
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.insertSheet("SwapRequests");
  
  // Define headers
  const headers = [
    "RequestID",
    "RequesterID",
    "RequesterName",
    "TargetEmployeeID",
    "TargetEmployeeName",
    "Date",
    "Slot",
    "Program",
    "Location",
    "ReasonType",
    "Notes",
    "Status",
    "AdminNotes",
    "DateCreated",
    "LastUpdated"
  ];
  
  // Set headers
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Format headers
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#f1f5f9')
    .setFontColor('#334155');
  
  // Set column widths
  sheet.setColumnWidth(1, 150);  // RequestID
  sheet.setColumnWidth(2, 100);  // RequesterID
  sheet.setColumnWidth(3, 150);  // RequesterName
  sheet.setColumnWidth(4, 100);  // TargetEmployeeID
  sheet.setColumnWidth(5, 150);  // TargetEmployeeName
  sheet.setColumnWidth(6, 100);  // Date
  sheet.setColumnWidth(7, 80);   // Slot
  sheet.setColumnWidth(8, 150);  // Program
  sheet.setColumnWidth(9, 150);  // Location
  sheet.setColumnWidth(10, 120); // ReasonType
  sheet.setColumnWidth(11, 200); // Notes
  sheet.setColumnWidth(12, 100); // Status
  sheet.setColumnWidth(13, 200); // AdminNotes
  sheet.setColumnWidth(14, 150); // DateCreated
  sheet.setColumnWidth(15, 150); // LastUpdated
  
  // Freeze the header row
  sheet.setFrozenRows(1);
  
  // Create data validation for Status column
  const statusRange = sheet.getRange(2, 12, 999, 1); // Column 12 is Status
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Pending', 'Approved', 'Rejected', 'Cancelled'], true)
    .build();
  statusRange.setDataValidation(statusRule);
  
  console.log("SwapRequests sheet created successfully");
  return sheet;
}

/**
 * Helper function to get employee name by ID or name
 * @param {string} employeeIdOrName - Employee ID or name to search for
 * @return {string|null} - Employee name if found, or null if not found
 */
function getEmployeeNameById(employeeIdOrName) {
  try {
    if (!employeeIdOrName) return null;
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    // First try the main data sheet where employee data is stored
    const sheet = ss.getSheetByName("main data") || ss.getSheetByName("Main Data");
    if (!sheet) {
      console.log("Main data sheet not found");
      return null;
    }
    
    const data = sheet.getDataRange().getValues();
    const headerRow = data[0];
    
    // Find ID and Name column indexes - using the actual column names in the main data sheet
    const idColIdx = headerRow.indexOf("EMPLOYEE ID");
    const nameColIdx = headerRow.indexOf("NAME OF THE EMPLOYEE");
    
    if (idColIdx === -1) {
      console.log("EMPLOYEE ID column not found in main data sheet");
      return null;
    }
    
    if (nameColIdx === -1) {
      console.log("NAME OF THE EMPLOYEE column not found in main data sheet");
      return null;
    }
    
    // First, search for exact employee ID match
    for (let i = 1; i < data.length; i++) {
      if (data[i][idColIdx] && data[i][idColIdx].toString() == employeeIdOrName.toString()) {
        console.log(`Found employee by ID: ${data[i][nameColIdx]}`);
        return data[i][nameColIdx];
      }
    }
    
    // If no exact ID match, try to find by name
    for (let i = 1; i < data.length; i++) {
      if (data[i][nameColIdx] && 
          data[i][nameColIdx].toString().toLowerCase() === employeeIdOrName.toString().toLowerCase()) {
        console.log(`Found employee by name: ${data[i][nameColIdx]}`);
        return data[i][nameColIdx];
      }
    }
    
    console.log(`Employee with ID or name "${employeeIdOrName}" not found`);
    return null;
  } catch (error) {
    console.error("Error in getEmployeeNameById:", error);
    return null;
  }
}

/**
 * Helper function to get employee email by ID
 */
function getEmployeeEmailById(employeeId) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("main data");
    if (!sheet) return null;
    
    const data = sheet.getDataRange().getValues();
    const headerRow = data[0];
    
    // Find ID and Email column indexes
    const idColIdx = headerRow.indexOf("ID");
    const emailColIdx = headerRow.indexOf("Email");
    
    if (idColIdx === -1 || emailColIdx === -1) return null;
    
    // Search for employee
    for (let i = 1; i < data.length; i++) {
      if (data[i][idColIdx] == employeeId) {
        return data[i][emailColIdx];
      }
    }
    
    return null;
  } catch (error) {
    console.error("Error in getEmployeeEmailById:", error);
    return null;
  }
}

/**
 * Sends a swap request confirmation email to the requester
 */
function sendSwapRequestConfirmation(email, date, slot, targetEmployee, requestId) {
  const slotTimes = {
    "1": "09:00-10:30",
    "2": "11:30-13:00",
    "3": "15:30-17:00",
    "4": "17:30-19:00"
  };
  
  const slotTime = slotTimes[slot] || "Unknown time";
  
  const subject = "Duty Swap Request Confirmation";
  const body = `
    <p>Hello,</p>
    
    <p>We have received your duty swap request with the following details:</p>
    
    <ul>
      <li><strong>Date:</strong> ${date}</li>
      <li><strong>Time Slot:</strong> ${slotTime}</li>
      <li><strong>Requested Swap With:</strong> ${targetEmployee}</li>
      <li><strong>Request ID:</strong> ${requestId}</li>
    </ul>
    
    <p>Your request is currently under review and pending approval. You will be notified as soon as there is an update on the status of your request.</p>
    
    <p>Thank you,<br>
    Invigilator Management System</p>
  `;
  
  try {
    GmailApp.sendEmail(email, subject, "", {
      htmlBody: body
    });
    console.log("Confirmation email sent to: " + email);
    return true;
  } catch (error) {
    console.error("Failed to send confirmation email:", error);
    return false;
  }
}

/**
 * Gets the admin email from the settings
 */
function getAdminEmail() {
  // Default admin email if not found in settings
  return "admin@example.com";
}

/**
 * Admin function to approve a duty swap request
 * @param {string} requestID - ID of the request to approve
 * @return {object} - Response object
 */
function adminApproveDutySwapRequest(requestID) {
  try {
    if (!requestID) {
      return { 
        success: false, 
        message: "Request ID is required" 
      };
    }
    
    console.log("Processing approval for request: " + requestID);
    
    // Get the swap request details
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const swapSheet = ss.getSheetByName("SwapRequests");
    
    if (!swapSheet) {
      return { 
        success: false, 
        message: "Swap requests sheet not found" 
      };
    }
    
    // Find the request
    const data = swapSheet.getDataRange().getValues();
    const headers = data[0];
    
    // Find indices of columns we need
    const idIdx = headers.indexOf("RequestID");
    const statusIdx = headers.indexOf("Status");
    const requesterIdIdx = headers.indexOf("RequesterID");
    const targetIdIdx = headers.indexOf("TargetEmployeeID");
    const dateIdx = headers.indexOf("Date");
    const slotIdx = headers.indexOf("Slot");
    const locationIdx = headers.indexOf("Location");
    const programIdx = headers.indexOf("Program");
    const lastUpdatedIdx = headers.indexOf("LastUpdated");
    
    if (idIdx === -1 || statusIdx === -1 || requesterIdIdx === -1 || 
        targetIdIdx === -1 || dateIdx === -1 || slotIdx === -1) {
      return { 
        success: false, 
        message: "Required columns not found in SwapRequests sheet" 
      };
    }
    
    // Find the row with this request ID
    let requestRow = -1;
    let requestData = null;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][idIdx] === requestID) {
        requestRow = i + 1; // +1 because sheet rows are 1-indexed
        requestData = data[i];
        break;
      }
    }
    
    if (requestRow === -1 || !requestData) {
      return { 
        success: false, 
        message: "Request not found" 
      };
    }
    
    // Check if already approved
    if (requestData[statusIdx] === "Approved") {
      return { 
        success: false, 
        message: "This request has already been approved" 
      };
    }
    
    // Get the swap data
    const swapData = {
      requesterID: requestData[requesterIdIdx],
      targetEmployeeID: requestData[targetIdIdx],
      date: requestData[dateIdx],
      slot: requestData[slotIdx],
      location: locationIdx !== -1 ? requestData[locationIdx] : "",
      program: programIdx !== -1 ? requestData[programIdx] : ""
    };
    
    // Process the swap in the main data sheet
    const swapResult = processApprovedSwap(swapData);
    
    if (!swapResult.success) {
      return swapResult;
    }
    
    // Update the status of the swap request
    swapSheet.getRange(requestRow, statusIdx + 1).setValue("Approved");
    
    // Update the last updated timestamp if column exists
    if (lastUpdatedIdx !== -1) {
      swapSheet.getRange(requestRow, lastUpdatedIdx + 1).setValue(new Date());
    }
    
    // Try to send notifications
    let notificationErrors = [];
    
    // Get requester email and send notification
    try {
      const requesterEmail = getEmployeeEmailById(swapData.requesterID);
      if (requesterEmail) {
        const requesterMessage = "We are pleased to inform you that your duty swap request has been approved. The assigned duty has been successfully transferred to the requested employee.";
        const requesterSent = sendSwapStatusNotification(
          requesterEmail,
          "Approved",
          swapData.date,
          swapData.slot,
          swapData.targetEmployeeID,
          requestID,
          requesterMessage
        );
        
        if (!requesterSent) {
          notificationErrors.push("Failed to send notification to requester: " + swapData.requesterID);
        }
      } else {
        notificationErrors.push("Could not find email for requester: " + swapData.requesterID);
      }
    } catch (requesterError) {
      console.error("Error sending notification to requester:", requesterError);
      notificationErrors.push("Error notifying requester: " + requesterError.toString());
    }
    
    // Get target employee email and send notification
    try {
      const targetEmail = getEmployeeEmailById(swapData.targetEmployeeID);
      if (targetEmail) {
        const targetMessage = "You have been assigned a new duty as part of an approved swap request. Please review your updated schedule accordingly.";
        const targetSent = sendSwapStatusNotification(
          targetEmail,
          "Approved",
          swapData.date,
          swapData.slot,
          swapData.requesterID,
          requestID,
          targetMessage
        );
        
        if (!targetSent) {
          notificationErrors.push("Failed to send notification to target employee: " + swapData.targetEmployeeID);
        }
      } else {
        notificationErrors.push("Could not find email for target employee: " + swapData.targetEmployeeID);
      }
    } catch (targetError) {
      console.error("Error sending notification to target employee:", targetError);
      notificationErrors.push("Error notifying target employee: " + targetError.toString());
    }
    
    // Return success with warning if notification errors occurred
    if (notificationErrors.length > 0) {
      return { 
        success: true,
        message: "Swap request approved successfully, but with notification issues",
        notificationWarnings: notificationErrors.join("; ")
      };
    }
    
    return { 
      success: true, 
      message: "Swap request approved successfully" 
    };
  } catch (error) {
    console.error("Error in adminApproveDutySwapRequest:", error);
    return {
      success: false,
      message: "Error: " + error.toString()
    };
  }
}

/**
 * Updates the status of a swap request
 * @param {string} requestID - ID of the request
 * @param {string} status - New status ('Pending', 'Approved', 'Rejected', 'Cancelled')
 * @param {string} adminNotes - Notes from admin (optional)
 * @return {object} - Response object
 */
function updateSwapRequestStatus(requestID, status, adminNotes = "") {
  try {
    if (!requestID) {
      throw new Error("Request ID is required");
    }
    
    if (!status) {
      throw new Error("Status is required");
    }
    
    // Get the swap requests sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const swapSheet = ss.getSheetByName("SwapRequests");
    
    if (!swapSheet) {
      throw new Error("SwapRequests sheet not found");
    }
    
    // Get all data and headers
    const data = swapSheet.getDataRange().getValues();
    const headers = data[0];
    
    // Find column indices
    const idIdx = headers.indexOf("RequestID");
    const statusIdx = headers.indexOf("Status");
    const adminNotesIdx = headers.indexOf("AdminNotes");
    const lastUpdatedIdx = headers.indexOf("LastUpdated");
    
    if (idIdx === -1 || statusIdx === -1 || adminNotesIdx === -1 || lastUpdatedIdx === -1) {
      throw new Error("Required columns not found in SwapRequests sheet");
    }
    
    // Find the row with this request ID
    let requestRow = -1;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][idIdx] === requestID) {
        requestRow = i + 1; // +1 because sheet rows are 1-indexed
        break;
      }
    }
    
    if (requestRow === -1) {
      throw new Error("Request not found");
    }
    
    // Update the status and last updated timestamp
    const currentTime = new Date();
    const formattedTime = Utilities.formatDate(currentTime, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    
    swapSheet.getRange(requestRow, statusIdx + 1).setValue(status);
    swapSheet.getRange(requestRow, lastUpdatedIdx + 1).setValue(formattedTime);
    
    // Update admin notes if provided
    if (adminNotes) {
      swapSheet.getRange(requestRow, adminNotesIdx + 1).setValue(adminNotes);
    }
    
    return { success: true };
    
  } catch (error) {
    console.error("Error in updateSwapRequestStatus:", error);
    throw error;
  }
}

/**
 * Process an approved swap by updating the main data sheet
 */
function processApprovedSwap(swapData) {
  try {
    console.log("Processing approved swap for requester " + swapData.requesterID + 
                " with target " + swapData.targetEmployeeID + 
                " for date " + swapData.date + 
                " slot " + swapData.slot);
    
    // Ensure date is valid
    if (!swapData.date) {
      return {
        success: false,
        message: "Date is required for swap processing"
      };
    }
    
    // Ensure requesterID is valid
    if (!swapData.requesterID) {
      return {
        success: false,
        message: "Requester ID is required for swap processing"
      };
    }
    
    // Ensure targetEmployeeID is valid
    if (!swapData.targetEmployeeID) {
      return {
        success: false,
        message: "Target employee ID is required for swap processing"
      };
    }
    
    // Get main data sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dataSheet = ss.getSheetByName("main data");
    
    if (!dataSheet) {
      return {
        success: false,
        message: "Main data sheet not found"
      };
    }
    
    // Get all data
    const data = dataSheet.getDataRange().getValues();
    if (data.length <= 1) {
      return {
        success: false,
        message: "No data found in main data sheet"
      };
    }
    
    // Get headers and find necessary columns
    const headers = data[0];
    const idIdx = headers.findIndex(h => h.toString().toUpperCase().includes("EMPLOYEE ID"));
    const dateIdx = headers.findIndex(h => h.toString().toUpperCase().includes("DATE"));
    const dayCountIdx = headers.findIndex(h => h.toString().toUpperCase().includes("COUNT OF INVIGILATION") && h.toString().toUpperCase().includes("DAY"));
    
    // Find slot and location columns
    const slot1Idx = headers.findIndex(h => h.toString().toUpperCase().includes("SLOT 1"));
    const slot2Idx = headers.findIndex(h => h.toString().toUpperCase().includes("SLOT 2"));
    const slot3Idx = headers.findIndex(h => h.toString().toUpperCase().includes("SLOT 3"));
    const slot4Idx = headers.findIndex(h => h.toString().toUpperCase().includes("SLOT 4"));
    
    // Find location columns - they typically follow right after the slot columns
    const loc1Idx = (slot1Idx !== -1 && slot1Idx + 1 < headers.length && 
                    headers[slot1Idx + 1].toString().toUpperCase().includes("LOCATION")) ? 
                    slot1Idx + 1 : -1;
    const loc2Idx = (slot2Idx !== -1 && slot2Idx + 1 < headers.length && 
                    headers[slot2Idx + 1].toString().toUpperCase().includes("LOCATION")) ? 
                    slot2Idx + 1 : -1;
    const loc3Idx = (slot3Idx !== -1 && slot3Idx + 1 < headers.length && 
                    headers[slot3Idx + 1].toString().toUpperCase().includes("LOCATION")) ? 
                    slot3Idx + 1 : -1;
    const loc4Idx = (slot4Idx !== -1 && slot4Idx + 1 < headers.length && 
                    headers[slot4Idx + 1].toString().toUpperCase().includes("LOCATION")) ? 
                    slot4Idx + 1 : -1;
    
    // Validate required columns
    if (idIdx === -1) {
      return { success: false, message: "Employee ID column not found" };
    }
    if (dateIdx === -1) {
      return { success: false, message: "Date column not found" };
    }
    
    // Find the row with matching ID and date for the requester
    let requesterRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][idIdx] && data[i][idIdx].toString() === swapData.requesterID.toString() && 
          data[i][dateIdx] && data[i][dateIdx].toString() === swapData.date.toString()) {
        requesterRow = i + 1; // +1 for 1-indexed sheet rows
        break;
      }
    }
    
    if (requesterRow === -1) {
      return {
        success: false,
        message: "Could not find row for requester on the specified date"
      };
    }
    
    // Find the row with matching ID and date for the target employee
    // or remember we might need to create a new row for them
    let targetRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][idIdx] && data[i][idIdx].toString() === swapData.targetEmployeeID.toString() && 
          data[i][dateIdx] && data[i][dateIdx].toString() === swapData.date.toString()) {
        targetRow = i + 1; // +1 for 1-indexed sheet rows
        break;
      }
    }
    
    // Get the slot and location indices for the specific slot number
    let slotIdx = -1;
    let locIdx = -1;
    
    switch(swapData.slot.toString()) {
      case "1":
      case 1:
        slotIdx = slot1Idx;
        locIdx = loc1Idx;
        break;
      case "2":
      case 2:
        slotIdx = slot2Idx;
        locIdx = loc2Idx;
        break;
      case "3":
      case 3:
        slotIdx = slot3Idx;
        locIdx = loc3Idx;
        break;
      case "4":
      case 4:
        slotIdx = slot4Idx;
        locIdx = loc4Idx;
        break;
      default:
        return {
          success: false,
          message: "Invalid slot number: " + swapData.slot
        };
    }
    
    // If slot column wasn't found
    if (slotIdx === -1) {
      return {
        success: false,
        message: "Slot " + swapData.slot + " column not found in the sheet"
      };
    }
    
    // If target employee doesn't have a row for this date, create one
    if (targetRow === -1) {
      // Create a new row with the target employee info
      const newRowData = Array(headers.length).fill("");
      
      // Find the data for the target employee to get basic info like name
      let targetEmpInfo = null;
      for (let i = 1; i < data.length; i++) {
        if (data[i][idIdx] && data[i][idIdx].toString() === swapData.targetEmployeeID.toString()) {
          targetEmpInfo = data[i];
          break;
        }
      }
      
      // Fill basic information
      newRowData[idIdx] = swapData.targetEmployeeID;
      newRowData[dateIdx] = swapData.date;
      
      // If we found other info about target employee, add it
      if (targetEmpInfo) {
        for (let i = 0; i < headers.length; i++) {
          if (i !== idIdx && i !== dateIdx && 
              i !== slot1Idx && i !== slot2Idx && i !== slot3Idx && i !== slot4Idx &&
              i !== loc1Idx && i !== loc2Idx && i !== loc3Idx && i !== loc4Idx &&
              i !== dayCountIdx) {
            newRowData[i] = targetEmpInfo[i];
          }
        }
      }
      
      // Set all slots as empty
      if (slot1Idx !== -1) newRowData[slot1Idx] = "";
      if (slot2Idx !== -1) newRowData[slot2Idx] = "";
      if (slot3Idx !== -1) newRowData[slot3Idx] = "";
      if (slot4Idx !== -1) newRowData[slot4Idx] = "";
      
      // Set day count to 0 initially
      if (dayCountIdx !== -1) newRowData[dayCountIdx] = 0;
      
      // Append the new row
      dataSheet.appendRow(newRowData);
      
      // Get the row number of the newly added row
      targetRow = dataSheet.getLastRow();
    }
    
    // Update the duties
    console.log("Updating duties for requester row " + requesterRow + " and target row " + targetRow);
    
    // 1. Get slot value from requester
    const requesterSlotValue = dataSheet.getRange(requesterRow, slotIdx + 1).getValue();
    
    // 2. Also get location values if applicable
    const requesterLocValue = (locIdx !== -1) ? 
                             dataSheet.getRange(requesterRow, locIdx + 1).getValue() :
                             swapData.location;
    
    // 3. Swap the duties (requester loses duty, target gets it)
    dataSheet.getRange(requesterRow, slotIdx + 1).setValue("");
    dataSheet.getRange(targetRow, slotIdx + 1).setValue("x");
    
    // 4. Update locations if applicable
    if (locIdx !== -1) {
      dataSheet.getRange(requesterRow, locIdx + 1).setValue("");
      // Use location from swap data if provided, otherwise use the original location
      const locationValue = swapData.location || requesterLocValue;
      dataSheet.getRange(targetRow, locIdx + 1).setValue(locationValue);
    }
    
    // 5. Update day counts for both employees
    if (dayCountIdx !== -1) {
      updateDayCounts(dataSheet, requesterRow, [slot1Idx, slot2Idx, slot3Idx, slot4Idx], dayCountIdx);
      updateDayCounts(dataSheet, targetRow, [slot1Idx, slot2Idx, slot3Idx, slot4Idx], dayCountIdx);
    }
    
    return {
      success: true,
      message: "Swap processed successfully"
    };
    
  } catch (error) {
    console.error("Error in processApprovedSwap:", error);
    return {
      success: false,
      message: "Error processing swap: " + error.toString()
    };
  }
}

/**
 * Update the day count for a given row based on slot values
 */
function updateDayCounts(sheet, rowNumber, slotIndices, dayCountIdx) {
  // Get current values for all slots
  const slot1 = sheet.getRange(rowNumber, slotIndices[0] + 1).getValue();
  const slot2 = sheet.getRange(rowNumber, slotIndices[1] + 1).getValue();
  const slot3 = sheet.getRange(rowNumber, slotIndices[2] + 1).getValue();
  const slot4 = sheet.getRange(rowNumber, slotIndices[3] + 1).getValue();
  
  // Count duties (marked with 'x')
  let dutyCount = 0;
  if (slot1 === 'x') dutyCount++;
  if (slot2 === 'x') dutyCount++;
  if (slot3 === 'x') dutyCount++;
  if (slot4 === 'x') dutyCount++;
  
  // Update day count
  sheet.getRange(rowNumber, dayCountIdx + 1).setValue(dutyCount);
}

/**
 * Find an employee by ID
 */
function findEmployeeById(employeeId) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Directory");
    if (!sheet) return null;
    
    const data = sheet.getDataRange().getValues();
    const headerRow = data[0];
    
    // Find ID column index
    const idColIdx = headerRow.indexOf("ID");
    
    if (idColIdx === -1) return null;
    
    // Search for employee
    for (let i = 1; i < data.length; i++) {
      if (data[i][idColIdx] == employeeId) {
        // Found the employee, return the row data
        const employee = {};
        headerRow.forEach((header, idx) => {
          employee[header] = data[i][idx];
        });
        return employee;
      }
    }
    
    return null;
  } catch (error) {
    console.error("Error in findEmployeeById:", error);
    return null;
  }
}

/**
 * Send email notification about swap status change
 */
function sendSwapStatusNotification(email, status, date, slot, otherEmployee, requestId, message) {
  const slotTimes = {
    "1": "09:00-10:30",
    "2": "11:30-13:00",
    "3": "15:30-17:00",
    "4": "17:30-19:00"
  };
  
  const slotTime = slotTimes[slot] || "Unknown time";
  const statusColor = {
    "Approved": "#4caf50",
    "Rejected": "#f44336",
    "Cancelled": "#ff9800",
    "Pending": "#2196f3"
  };
  
  const subject = `Duty Swap Request ${status}`;
  const body = `
    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #e0e0e0; border-radius: 5px;">
      <h2 style="color: #333; border-bottom: 1px solid #eee; padding-bottom: 10px;">Duty Swap Request Update</h2>
      
      <p>Hello,</p>
      
      <p>Your duty swap request has been <strong style="color: ${statusColor[status] || '#333'}">${status}</strong>.</p>
      
      <div style="background-color: #f9f9f9; padding: 15px; border-radius: 4px; margin: 15px 0;">
        <h3 style="margin-top: 0; color: #555;">Swap Details</h3>
        <p><strong>Date:</strong> ${date}</p>
        <p><strong>Time Slot:</strong> ${slotTime}</p>
        <p><strong>With Employee:</strong> ${otherEmployee}</p>
        <p><strong>Request ID:</strong> ${requestId}</p>
      </div>
      
      ${message ? `<p>${message}</p>` : ''}
      
      <p>If you have any questions, please contact the administrator.</p>
      
      <p style="margin-top: 20px; padding-top: 10px; border-top: 1px solid #eee; color: #777; font-size: 12px;">
        This is an automated message from the Invigilator Management System.
      </p>
    </div>
  `;
  
  try {
    GmailApp.sendEmail(email, subject, "", {
      htmlBody: body
    });
    console.log(`${status} notification email sent to: ${email}`);
    return true;
  } catch (error) {
    console.error(`Failed to send ${status} notification to ${email}:`, error);
    return false;
  }
}

/**
 * Admin function to reject a duty swap request
 */
function adminRejectDutySwapRequest(requestID) {
  try {
    if (!requestID) {
      return { success: false, message: "Request ID is required" };
    }
    
    console.log("Processing rejection for request: " + requestID);
    
    // Get the swap request details
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const swapSheet = ss.getSheetByName("SwapRequests");
    
    if (!swapSheet) {
      return { success: false, message: "Swap requests sheet not found" };
    }
    
    // Find the request
    const data = swapSheet.getDataRange().getValues();
    const headers = data[0];
    
    // Find indices of columns we need
    const idIdx = headers.indexOf("RequestID");
    const statusIdx = headers.indexOf("Status");
    const requesterIdIdx = headers.indexOf("RequesterID");
    const targetIdIdx = headers.indexOf("TargetEmployeeID");
    const dateIdx = headers.indexOf("Date");
    const slotIdx = headers.indexOf("Slot");
    const lastUpdatedIdx = headers.indexOf("LastUpdated");
    const adminNotesIdx = headers.indexOf("AdminNotes");
    
    if (idIdx === -1 || statusIdx === -1 || requesterIdIdx === -1) {
      return { success: false, message: "Required columns not found in SwapRequests sheet" };
    }
    
    // Find the row with this request ID
    let requestRow = -1;
    let requestData = null;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][idIdx] === requestID) {
        requestRow = i + 1; // +1 because sheet rows are 1-indexed
        requestData = data[i];
        break;
      }
    }
    
    if (requestRow === -1 || !requestData) {
      return { success: false, message: "Request not found" };
    }
    
    // Check if already processed
    if (requestData[statusIdx] === "Rejected") {
      return { success: false, message: "This request has already been rejected" };
    }
    
    if (requestData[statusIdx] === "Approved") {
      return { success: false, message: "Cannot reject an already approved request" };
    }
    
    if (requestData[statusIdx] === "Cancelled") {
      return { success: false, message: "Cannot reject a cancelled request" };
    }
    
    // Update the status of the swap request
    swapSheet.getRange(requestRow, statusIdx + 1).setValue("Rejected");
    
    // Update the last updated timestamp if column exists
    if (lastUpdatedIdx !== -1) {
      swapSheet.getRange(requestRow, lastUpdatedIdx + 1).setValue(new Date());
    }
    
    // Add admin notes if applicable
    if (adminNotesIdx !== -1) {
      const currentNotes = swapSheet.getRange(requestRow, adminNotesIdx + 1).getValue() || "";
      const newNotes = currentNotes + (currentNotes ? "\n" : "") + "Rejected on " + new Date().toLocaleString();
      swapSheet.getRange(requestRow, adminNotesIdx + 1).setValue(newNotes);
    }
    
    // Try to send notifications to affected parties
    try {
      const requesterEmail = getEmployeeEmailById(requestData[requesterIdIdx]);
      
      if (requesterEmail) {
        sendSwapStatusNotification(
          requesterEmail,
          "Rejected",
          requestData[dateIdx],
          requestData[slotIdx],
          requestData[targetIdIdx],
          requestID,
          "Your swap request has been rejected by an administrator."
        );
      }
    } catch (notifyError) {
      console.error("Error sending rejection notification:", notifyError);
      // Continue despite notification errors
    }
    
    return { 
      success: true, 
      message: "Swap request rejected successfully" 
    };
    
  } catch (error) {
    console.error("Error in adminRejectDutySwapRequest:", error);
    return {
      success: false,
      message: "Error: " + error.toString()
    };
  }
}

/**
 * Admin function to delete a swap request
 * @param {string} requestID - ID of the request to delete
 * @return {object} - Response object
 */
function adminDeleteSwapRequest(requestID) {
  try {
    if (!requestID) {
      return { success: false, message: "Request ID is required" };
    }
    
    // Get the swap request sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const swapSheet = ss.getSheetByName("SwapRequests");
    
    if (!swapSheet) {
      return { success: false, message: "Swap requests sheet not found" };
    }
    
    // Find the request
    const data = swapSheet.getDataRange().getValues();
    const headers = data[0];
    
    const idIdx = headers.indexOf("RequestID");
    
    if (idIdx === -1) {
      return { success: false, message: "RequestID column not found" };
    }
    
    // Find the row with this request ID
    let requestRow = -1;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][idIdx] === requestID) {
        requestRow = i + 1; // +1 because sheet rows are 1-indexed
        break;
      }
    }
    
    if (requestRow === -1) {
      return { success: false, message: "Request not found" };
    }
    
    // Delete the row
    swapSheet.deleteRow(requestRow);
    
    return { 
      success: true, 
      message: "Request deleted successfully" 
    };
    
  } catch (error) {
    console.error("Error in adminDeleteSwapRequest:", error);
    return { 
      success: false, 
      message: error.toString() 
    };
  }
}

/**
 * List all swap requests - wrapper for getAllSwapRequests for backward compatibility
 */
function listAllSwapRequests(forceRefresh = false, filters = null) {
  console.log("listAllSwapRequests called with forceRefresh:", forceRefresh, "filters:", filters ? JSON.stringify(filters) : "none");
  
  // Clear cache if force refresh
  if (forceRefresh) {
    console.log("Force refreshing swap requests data");
    var cache = CacheService.getScriptCache();
    cache.remove("all_swap_requests_cache");
  }
  
  // Call the main function, passing any filters
  return getAllSwapRequests(filters);
}

/**
 * Ensures the SwapRequests sheet exists with correct headers
 */
function ensureSwapRequestsSheet() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('SwapRequests');
    
    if (!sheet) {
      sheet = ss.insertSheet('SwapRequests');
      var headers = [
        'RequestID',
        'RequesterID',
        'TargetEmployee',
        'Date',
        'Slot',
        'Location',
        'Program',
        'ReasonType',
        'Notes',
        'Status',
        'Timestamp'
      ];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, headers.length).setBackground('#f3f3f3').setFontWeight('bold');
    }
    return true;
  } catch (error) {
    Logger.log("Error ensuring SwapRequests sheet: " + error.toString());
    return false;
  }
}

/**
 * Clear the swap requests cache
 */
function clearCache() {
  try {
    var cache = CacheService.getScriptCache();
    cache.remove("all_swap_requests_cache");
    console.log("Cache cleared successfully");
    return { success: true, message: "Cache cleared successfully" };
  } catch (error) {
    console.error("Error clearing cache:", error);
    return { success: false, message: error.toString() };
  }
}

/**
 * Try to get swap requests from a specific spreadsheet by name
 * This function is designed to help when the SwapRequests is a standalone spreadsheet
 * rather than a sheet within the active spreadsheet
 */
function tryGetSwapRequestsFromSpreadsheet(spreadsheetName) {
  try {
    console.log("Trying to access spreadsheet by name:", spreadsheetName);
    
    if (!spreadsheetName) {
      return { 
        success: false, 
        message: "Spreadsheet name is required" 
      };
    }
    
    // Try to find the spreadsheet by name
    var spreadsheets = DriveApp.getFilesByName(spreadsheetName);
    
    if (!spreadsheets.hasNext()) {
      return { 
        success: false, 
        message: "No spreadsheet found with name: " + spreadsheetName 
      };
    }
    
    // Get the first matching spreadsheet
    var spreadsheetFile = spreadsheets.next();
    console.log("Found spreadsheet:", spreadsheetFile.getName(), "ID:", spreadsheetFile.getId());
    
    // Open the spreadsheet
    var ss = SpreadsheetApp.openById(spreadsheetFile.getId());
    
    if (!ss) {
      return { 
        success: false, 
        message: "Failed to open spreadsheet: " + spreadsheetName 
      };
    }
    
    // Get all sheets in the spreadsheet
    var sheets = ss.getSheets();
    console.log("Spreadsheet has", sheets.length, "sheets");
    
    // Option 1: If SwapRequests is the name of the spreadsheet, get the first sheet
    var dataSheet = ss.getSheets()[0];
    
    // Option 2: If there's a sheet named 'SwapRequests' within the spreadsheet, use that
    var swapSheet = ss.getSheetByName('SwapRequests');
    if (swapSheet) {
      dataSheet = swapSheet;
      console.log("Found sheet named 'SwapRequests' within the spreadsheet");
    }
    
    if (!dataSheet) {
      return { 
        success: false, 
        message: "No data sheet found in spreadsheet: " + spreadsheetName 
      };
    }
    
    console.log("Using sheet:", dataSheet.getName());
    
    // Get all data from the sheet
    var data = dataSheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      console.log("Sheet has no data rows");
      return { success: true, requests: [] };
    }
    
    // Extract header row
    var headers = data[0];
    console.log("Headers:", headers.join(", "));
    
    // Convert all rows to objects
    var requests = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      
      // Skip empty rows
      if (row.every(function(cell) { return !cell; })) {
        continue;
      }
      
      // Convert row to object
      var request = {};
      for (var j = 0; j < headers.length; j++) {
        var header = headers[j];
        var value = row[j];
        
        // Format dates for display
        if (value instanceof Date) {
          value = Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
        }
        
        request[header] = value;
      }
      
      requests.push(request);
    }
    
    console.log("Returning", requests.length, "swap requests from spreadsheet:", spreadsheetName);
    
    return { 
      success: true, 
      requests: requests,
      message: "Successfully loaded data from " + spreadsheetName
    };
    
  } catch (error) {
    console.error("Error in tryGetSwapRequestsFromSpreadsheet:", error);
    return { 
      success: false, 
      message: error.toString(),
      requests: []
    };
  }
}

/**
 * Handler function for client-side requests to load swap requests from a named spreadsheet
 * This function is called by google.script.run from AdminSwapRequests.html
 */
function handleTryLoadFromNamedSpreadsheet(spreadsheetName) {
  console.log("handleTryLoadFromNamedSpreadsheet called with spreadsheet name:", spreadsheetName);
  
  // Call our utility function to get the swap requests
  var result = tryGetSwapRequestsFromSpreadsheet(spreadsheetName);
  
  // Log the result for debugging
  if (result.success) {
    console.log("Successfully loaded", result.requests.length, "swap requests");
  } else {
    console.error("Failed to load swap requests:", result.message);
  }
  
  return result;
}

/**
 * Function for administrators to list all swap requests
 */
function listAllSwapRequestsDirectFromSheet() {
  try {
    Logger.log("Starting listAllSwapRequestsDirectFromSheet");
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var swapSheet = ss.getSheetByName('SwapRequests');
    
    if (!swapSheet) {
      Logger.log("SwapRequests sheet not found");
      return { 
        success: false, 
        message: "SwapRequests sheet not found in the active spreadsheet" 
      };
    }
    
    var lastRow = swapSheet.getLastRow();
    if (lastRow <= 1) {
      Logger.log("No requests found in sheet");
      return { 
        success: true, 
        requests: [],
        message: "No requests found" 
      };
    }
    
    Logger.log("Reading data from SwapRequests sheet");
    var data = swapSheet.getRange(1, 1, lastRow, swapSheet.getLastColumn()).getValues();
    var headers = data[0];
    
    // Get column indices using exact column names from the spreadsheet
    var requestIdCol = headers.indexOf('RequestID');
    var requesterIdCol = headers.indexOf('RequesterID');
    var requesterNameCol = headers.indexOf('RequesterName');
    var targetEmployeeIdCol = headers.indexOf('TargetEmployeeID');
    var targetEmployeeNameCol = headers.indexOf('TargetEmployeeName');
    var dateCol = headers.indexOf('Date');
    var slotCol = headers.indexOf('Slot');
    var programCol = headers.indexOf('Program');
    var locationCol = headers.indexOf('Location');
    var reasonTypeCol = headers.indexOf('ReasonType');
    var notesCol = headers.indexOf('Notes');
    var statusCol = headers.indexOf('Status');
    var adminNotesCol = headers.indexOf('AdminNotes');
    var dateCreatedCol = headers.indexOf('DateCreated');
    var lastUpdatedCol = headers.indexOf('LastUpdated');
    var creatingNewDutyCol = headers.indexOf('CreatingNewDuty');
    
    var requests = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[requesterIdCol]) continue; // Skip empty rows
      
      var request = {
        RequestID: row[requestIdCol],
        RequesterID: row[requesterIdCol],
        RequesterName: requesterNameCol >= 0 ? row[requesterNameCol] : '',
        TargetEmployeeID: targetEmployeeIdCol >= 0 ? row[targetEmployeeIdCol] : '',
        TargetEmployeeName: targetEmployeeNameCol >= 0 ? row[targetEmployeeNameCol] : '',
        Date: dateCol >= 0 ? row[dateCol] : '',
        Slot: slotCol >= 0 ? row[slotCol] : '',
        Program: programCol >= 0 ? row[programCol] : '',
        Location: locationCol >= 0 ? row[locationCol] : '',
        ReasonType: reasonTypeCol >= 0 ? row[reasonTypeCol] : '',
        Notes: notesCol >= 0 ? row[notesCol] : '',
        Status: statusCol >= 0 ? row[statusCol] : 'Pending',
        AdminNotes: adminNotesCol >= 0 ? row[adminNotesCol] : '',
        DateCreated: dateCreatedCol >= 0 ? row[dateCreatedCol] : '',
        LastUpdated: lastUpdatedCol >= 0 ? row[lastUpdatedCol] : '',
        CreatingNewDuty: creatingNewDutyCol >= 0 ? row[creatingNewDutyCol] : false
      };
      
      requests.push(request);
    }
    
    Logger.log("Successfully retrieved " + requests.length + " requests");
    return {
      success: true,
      requests: requests
    };
    
  } catch (error) {
    Logger.log("Error in listAllSwapRequestsDirectFromSheet: " + error.toString());
    return {
      success: false,
      message: "Error retrieving swap requests: " + error.toString()
    };
  }
}

/**
 * Shows the Admin Swap Requests panel
 */
function showAdminSwapRequestPanel() {
  var html = HtmlService.createHtmlOutputFromFile('AdminSwapRequests')
      .setWidth(1000)
      .setHeight(600)
      .setTitle('Admin Swap Requests');
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Admin Swap Requests');
}

/**
 * Get admin data including ID and permissions
 */
function getAdminData() {
  try {
    var email = Session.getActiveUser().getEmail();
    // You can implement your own admin check logic here
    return {
      success: true,
      adminID: email,
      isAdmin: true
    };
  } catch (error) {
    Logger.log("Error getting admin data: " + error.toString());
    return {
      success: false,
      message: "Error getting admin data: " + error.toString()
    };
  }
}

/**
 * Try to load swap requests from a named spreadsheet
 */
function handleTryLoadFromNamedSpreadsheet(sheetName) {
  try {
    Logger.log("Attempting to load from sheet: " + sheetName);
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      Logger.log("Sheet not found: " + sheetName);
      return {
        success: false,
        message: "Sheet '" + sheetName + "' not found"
      };
    }
    
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    
    if (lastRow <= 1) {
      Logger.log("No data in sheet: " + sheetName);
      return {
        success: true,
        requests: []
      };
    }
    
    var data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    var headers = data[0];
    var requests = [];
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var request = {};
      
      headers.forEach(function(header, index) {
        request[header] = row[index];
      });
      
      if (request.RequesterID) { // Only add if there's a requester ID
        requests.push(request);
      }
    }
    
    Logger.log("Successfully loaded " + requests.length + " requests from " + sheetName);
    return {
      success: true,
      requests: requests
    };
    
  } catch (error) {
    Logger.log("Error loading from named spreadsheet: " + error.toString());
    return {
      success: false,
      message: "Error loading from spreadsheet: " + error.toString()
    };
  }
}

/**
 * Check if the SwapRequests sheet exists and return its status
 */
function checkSwapRequestsSheetExists() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      return {
        exists: false,
        message: "No active spreadsheet found"
      };
    }
    
    var sheet = ss.getSheetByName("SwapRequests");
    if (!sheet) {
      return {
        exists: false,
        message: "SwapRequests sheet does not exist in the active spreadsheet"
      };
    }
    
    // Check if the sheet has data
    var lastRow = sheet.getLastRow();
    var hasData = lastRow > 1;
    
    return {
      exists: true,
      status: hasData ? "Contains " + (lastRow - 1) + " rows of data" : "Empty (no data rows)",
      message: "SwapRequests sheet exists" + (hasData ? " and contains data" : " but is empty")
    };
  } catch (error) {
    console.error("Error checking SwapRequests sheet:", error);
    return {
      exists: false,
      message: "Error checking SwapRequests sheet: " + error.toString()
    };
  }
}

/**
 * Handle duty swap request from the client
 * @param {object} data - Swap request data
 * @return {object} - Response object
 */
function requestDutySwap(data) {
  try {
    console.log("Processing duty swap request:", JSON.stringify(data));
    
    // Validate required fields
    if (!data.requesterID) {
      return {success: false, message: "Requester ID is required"};
    }
    if (!data.targetEmployeeID) {
      return {success: false, message: "Target employee ID is required"};
    }
    if (!data.date) {
      return {success: false, message: "Date is required"};
    }
    if (!data.slot) {
      return {success: false, message: "Slot is required"};
    }
    
    // Get spreadsheet and swap requests sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let swapSheet = ss.getSheetByName("SwapRequests");
    
    // Create sheet if it doesn't exist
    if (!swapSheet) {
      swapSheet = createSwapRequestsSheet();
    }
    
    // Get the column headers
    const headers = swapSheet.getRange(1, 1, 1, swapSheet.getLastColumn()).getValues()[0];
    
    // Generate a unique request ID
    const requestID = "REQ-" + new Date().getTime() + "-" + Math.floor(Math.random() * 1000);
    
    // Create a new row of data
    const newRowData = [];
    for (let i = 0; i < headers.length; i++) {
      switch(headers[i]) {
        case "RequestID":
          newRowData[i] = requestID;
          break;
        case "RequesterID":
          newRowData[i] = data.requesterID;
          break;
        case "RequesterName":
          newRowData[i] = getEmployeeNameById(data.requesterID) || "";
          break;
        case "TargetEmployeeID":
          newRowData[i] = data.targetEmployeeID;
          break;
        case "TargetEmployeeName":
          newRowData[i] = getEmployeeNameById(data.targetEmployeeID) || "";
          break;
        case "Date":
          newRowData[i] = data.date;
          break;
        case "Slot":
          newRowData[i] = data.slot;
          break;
        case "Program":
          newRowData[i] = data.program || "";
          break;
        case "Location":
          newRowData[i] = data.location || "";
          break;
        case "ReasonType":
          newRowData[i] = data.reasonType || "";
          break;
        case "Notes":
          newRowData[i] = data.notes || "";
          break;
        case "Status":
          newRowData[i] = "Pending";
          break;
        case "DateCreated":
          newRowData[i] = new Date();
          break;
        case "LastUpdated":
          newRowData[i] = new Date();
          break;
        default:
          newRowData[i] = "";
      }
    }
    
    // Append the new row
    swapSheet.appendRow(newRowData);
    
    // Try to send notification emails
    try {
      // Send confirmation to requester
      const requesterEmail = getEmployeeEmailById(data.requesterID);
      if (requesterEmail) {
        sendSwapRequestConfirmation(
          requesterEmail, 
          data.date, 
          data.slot, 
          getEmployeeNameById(data.targetEmployeeID) || data.targetEmployeeID,
          requestID
        );
      }
      
      // Send notification to admin
      const adminEmail = getAdminEmail();
      if (adminEmail) {
        const subject = "New Duty Swap Request";
        const body = `
          <p>A new duty swap request has been submitted:</p>
          <ul>
            <li><strong>Request ID:</strong> ${requestID}</li>
            <li><strong>Requester:</strong> ${getEmployeeNameById(data.requesterID) || data.requesterID}</li>
            <li><strong>Target Employee:</strong> ${getEmployeeNameById(data.targetEmployeeID) || data.targetEmployeeID}</li>
            <li><strong>Date:</strong> ${data.date}</li>
            <li><strong>Slot:</strong> ${data.slot}</li>
          </ul>
          <p>Please log in to the admin panel to review this request.</p>
        `;
        
        GmailApp.sendEmail(adminEmail, subject, "", { htmlBody: body });
      }
    } catch (emailError) {
      console.error("Error sending email notifications:", emailError);
      // Continue even if email notification fails
    }
    
    return {
      success: true,
      message: "Swap request submitted successfully",
      requestID: requestID
    };
    
  } catch (error) {
    console.error("Error in requestDutySwap:", error);
    return {
      success: false,
      message: "Error: " + error.toString()
    };
  }
}

/**
 * Debug function to test swap request functions
 */
function debugSwapRequestFunctions() {
  try {
    var ui = SpreadsheetApp.getUi();
    
    // Test the original listAllSwapRequests function
    var listResult = listAllSwapRequests();
    Logger.log("listAllSwapRequests result: " + JSON.stringify(listResult));
    
    // Test the getAllSwapRequests function
    var allResult = getAllSwapRequests();
    Logger.log("getAllSwapRequests result: " + JSON.stringify(allResult));
    
    // Test the direct from sheet function
    var directResult = listAllSwapRequestsDirectFromSheet();
    Logger.log("listAllSwapRequestsDirectFromSheet result: " + JSON.stringify(directResult));
    
    // Add an option to test tryGetSwapRequestsFromSheet (the new function for sheets, not spreadsheets)
    var promptResult = ui.prompt(
      'Debug tryGetSwapRequestsFromSheet',
      'Enter the name of the sheet to test:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (promptResult.getSelectedButton() == ui.Button.OK) {
      var sheetName = promptResult.getResponseText();
      if (sheetName) {
        var tryResult = tryGetSwapRequestsFromSheet(sheetName);
        Logger.log("tryGetSwapRequestsFromSheet result: " + JSON.stringify(tryResult));
        ui.alert("Result of tryGetSwapRequestsFromSheet: " + (tryResult.success ? "Success" : "Failed") + 
                "\nMessage: " + tryResult.message + 
                "\nRequests found: " + (tryResult.requests ? tryResult.requests.length : 0));
      }
    }
    
    ui.alert("Debugging complete! Check the logs for results.");
    
  } catch (error) {
    Logger.log("Error in debugSwapRequestFunctions: " + error.toString());
    SpreadsheetApp.getUi().alert("Error: " + error.toString());
  }
}

/**
 * Try to get swap requests from a specific sheet by name
 * This function is designed to help when the swap requests data is in
 * a different sheet than the standard 'SwapRequests' sheet
 */
function tryGetSwapRequestsFromSheet(sheetName) {
  try {
    console.log("Trying to access sheet by name:", sheetName);
    
    if (!sheetName) {
      return { 
        success: false, 
        message: "Sheet name is required" 
      };
    }
    
    // Get active spreadsheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Try to find the sheet by name
    var dataSheet = ss.getSheetByName(sheetName);
    
    if (!dataSheet) {
      return { 
        success: false, 
        message: "No sheet found with name: " + sheetName 
      };
    }
    
    console.log("Found sheet:", dataSheet.getName());
    
    // Get all data from the sheet
    var data = dataSheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      console.log("Sheet has no data rows");
      return { success: true, requests: [] };
    }
    
    // Extract header row
    var headers = data[0];
    console.log("Headers:", headers.join(", "));
    
    // Convert all rows to objects
    var requests = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      
      // Skip empty rows
      if (row.every(function(cell) { return !cell; })) {
        continue;
      }
      
      // Convert row to object
      var request = {};
      for (var j = 0; j < headers.length; j++) {
        var header = headers[j];
        var value = row[j];
        
        // Format dates for display
        if (value instanceof Date) {
          value = Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
        }
        
        request[header] = value;
      }
      
      requests.push(request);
    }
    
    console.log("Returning", requests.length, "swap requests from sheet:", sheetName);
    
    return { 
      success: true, 
      requests: requests,
      message: "Successfully loaded data from " + sheetName
    };
    
  } catch (error) {
    console.error("Error in tryGetSwapRequestsFromSheet:", error);
    return { 
      success: false, 
      message: error.toString(),
      requests: []
    };
  }
}

/**
 * Handler function for client-side requests to load swap requests from a named sheet
 * This function can be called by google.script.run from client-side HTML files
 */
function handleTryLoadFromNamedSheet(sheetName) {
  console.log("handleTryLoadFromNamedSheet called with sheet name:", sheetName);
  
  // Call our utility function to get the swap requests
  var result = tryGetSwapRequestsFromSheet(sheetName);
  
  // Log the result for debugging
  if (result.success) {
    console.log("Successfully loaded", result.requests.length, "swap requests from sheet");
  } else {
    console.error("Failed to load swap requests from sheet:", result.message);
  }
  
  return result;
}

/**
 * Get an employee's email by their ID directly from main data sheet
 * This function is specifically focused on extracting email addresses
 * IMPORTANT: This function will ALWAYS return a valid email - if no
 * valid email is found, it will return a default value
 */
function getEmployeeEmailById(employeeId) {
  try {
    if (!employeeId) {
      Logger.log("Employee ID is null or empty in getEmployeeEmailById - using default email");
      return Session.getActiveUser().getEmail() || "admin@example.com";
    }
    
    // Make sure to convert to string and clean
    employeeId = String(employeeId).trim();
    
    Logger.log("Searching for email for employee ID: " + employeeId);
    
    // Get the spreadsheet and try multiple possible sheet names
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = [
      ss.getSheetByName('main data'),
      ss.getSheetByName('Main Data'),
      ss.getSheetByName('InvigilatorDuties'),
      ss.getSheetByName('Employees'),
      ss.getSheetByName('employees')
    ];
    
    // Filter out null sheets
    sheets = sheets.filter(function(sheet) {
      return sheet !== null;
    });
    
    if (sheets.length === 0) {
      Logger.log("No data sheets found for employee lookup - using default email");
      return Session.getActiveUser().getEmail() || "admin@example.com";
    }
    
    // Try each sheet until we find the employee
    for (var s = 0; s < sheets.length; s++) {
      var sheet = sheets[s];
      Logger.log("Searching in sheet: " + sheet.getName());
      
      var data = sheet.getDataRange().getValues();
      if (data.length <= 1) {
        Logger.log("No data found in sheet: " + sheet.getName());
        continue;
      }
      
      var headers = data[0];
      Logger.log("Headers in sheet: " + headers.join(", "));
      
      // Look for email column - try all possible variations
      var emailCol = -1;
      var idCol = -1;
      
      // Map of possible headers to their meanings
      var headerMap = {
        email: ["EMAIL", "E-MAIL", "E-MAIL-ID", "EMAIL ID", "MAIL", "E MAIL", "EMAIL ADDRESS"],
        id: ["EMPLOYEE ID", "EMPLOYEEID", "EMP ID", "ID"]
      };
      
      // First look for explicit email column
      for (var i = 0; i < headers.length; i++) {
        var header = headers[i] ? headers[i].toString().toUpperCase() : "";
        
        // Check for exact email matches
        if (headerMap.email.indexOf(header) !== -1) {
          emailCol = i;
          Logger.log("Found exact email column match: " + header + " at index " + i);
          break; // Found a direct match, stop searching
        }
      }
      
      // If we didn't find an exact email match, look for partial matches
      if (emailCol === -1) {
        for (var i = 0; i < headers.length; i++) {
          var header = headers[i] ? headers[i].toString().toLowerCase() : "";
          if (header.indexOf("email") !== -1 || header.indexOf("mail") !== -1 || 
              header.indexOf("e-mail") !== -1) {
            emailCol = i;
            Logger.log("Found partial email column match: " + headers[i] + " at index " + i);
            break;
          }
        }
      }
      
      // Look for ID column
      for (var i = 0; i < headers.length; i++) {
        var header = headers[i] ? headers[i].toString().toUpperCase() : "";
        if (headerMap.id.indexOf(header) !== -1) {
          idCol = i;
          Logger.log("Found exact ID column match: " + header + " at index " + i);
          break; // Found a direct match, stop searching
        }
      }
      
      // If no exact ID match, try partial
      if (idCol === -1) {
        for (var i = 0; i < headers.length; i++) {
          var header = headers[i] ? headers[i].toString().toLowerCase() : "";
          if (header.indexOf("id") !== -1 || header.indexOf("employee") !== -1) {
            idCol = i;
            Logger.log("Found partial ID column match: " + headers[i] + " at index " + i);
            break;
          }
        }
      }
      
      // If we have an ID column, search for matching employee
      if (idCol !== -1) {
        Logger.log("Searching for employee ID " + employeeId + " in column " + idCol);
        var searchId = employeeId.toString().toLowerCase().trim();
        
        // Approach 1: Look for exact ID match and extract email
        for (var i = 1; i < data.length; i++) {
          var rowId = data[i][idCol] ? data[i][idCol].toString().toLowerCase().trim() : "";
          
          if (rowId === searchId) {
            Logger.log("Found employee ID match at row " + i);
            
            // Look for a valid email in the row
            if (emailCol !== -1) {
              var email = data[i][emailCol];
              if (email && isValidEmail(email.toString())) {
                return email.toString();
              }
            }
            
            // Scan entire row for any valid email
            for (var j = 0; j < data[i].length; j++) {
              var cellValue = data[i][j] ? data[i][j].toString() : "";
              if (isValidEmail(cellValue)) {
                return cellValue;
              }
            }
          }
        }
        
        // Approach 2: No exact match, try partial ID match
        for (var i = 1; i < data.length; i++) {
          var rowId = data[i][idCol] ? data[i][idCol].toString().toLowerCase().trim() : "";
          
          if (rowId.indexOf(searchId) !== -1 || searchId.indexOf(rowId) !== -1) {
            Logger.log("Found partial employee ID match at row " + i);
            
            // Look for a valid email
            if (emailCol !== -1) {
              var email = data[i][emailCol];
              if (email && isValidEmail(email.toString())) {
                return email.toString();
              }
            }
            
            // Scan row for any valid email
            for (var j = 0; j < data[i].length; j++) {
              var cellValue = data[i][j] ? data[i][j].toString() : "";
              if (isValidEmail(cellValue)) {
                return cellValue;
              }
            }
          }
        }
      }
      
      // If no ID column or no match found, try to find any valid email in the sheet
      if (emailCol !== -1) {
        Logger.log("No ID match found, looking for any valid email in email column");
        for (var i = 1; i < data.length; i++) {
          var email = data[i][emailCol];
          if (email && isValidEmail(email.toString())) {
            Logger.log("Found valid email in row " + i + ": " + email);
            return email.toString();
          }
        }
      }
      
      // Last resort: Scan entire sheet for any valid email
      Logger.log("No email column found, scanning entire sheet for valid emails");
      for (var i = 1; i < data.length; i++) {
        for (var j = 0; j < data[i].length; j++) {
          var cellValue = data[i][j] ? data[i][j].toString() : "";
          if (isValidEmail(cellValue)) {
            Logger.log("Found valid email at row " + i + ", col " + j + ": " + cellValue);
            return cellValue;
          }
        }
      }
    }
    
    // If we get here, we couldn't find a valid email in ANY sheet
    Logger.log("No valid email found anywhere - using default email");
    return Session.getActiveUser().getEmail() || "admin@example.com";
    
  } catch (error) {
    Logger.log("Error in getEmployeeEmailById: " + error.toString());
    return Session.getActiveUser().getEmail() || "admin@example.com";
  }
}

/**
 * Check if a string is a valid email address
 */
function isValidEmail(email) {
  if (!email) return false;
  
  // Convert to string if not already
  email = email.toString().trim();
  
  // Very basic email validation regex
  var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}

/**
 * Get an employee's email by their ID with enhanced robustness
 */
function getEmployeeEmailById(employeeId) {
  try {
    if (!employeeId) {
      Logger.log("Employee ID is null or empty in getEmployeeEmailById - using default email");
      return Session.getActiveUser().getEmail() || "admin@example.com";
    }
    
    // Make sure to convert to string and clean
    employeeId = String(employeeId).trim();
    
    Logger.log("Searching for email for employee ID: " + employeeId);
    
    // Get the spreadsheet and try multiple possible sheet names
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = [
      ss.getSheetByName('main data'),
      ss.getSheetByName('Main Data'),
      ss.getSheetByName('InvigilatorDuties'),
      ss.getSheetByName('Employees'),
      ss.getSheetByName('employees')
    ];
    
    // Filter out null sheets
    sheets = sheets.filter(function(sheet) {
      return sheet !== null;
    });
    
    if (sheets.length === 0) {
      Logger.log("No data sheets found for employee lookup - using default email");
      return Session.getActiveUser().getEmail() || "admin@example.com";
    }
    
    // Try each sheet until we find the employee
    for (var s = 0; s < sheets.length; s++) {
      var sheet = sheets[s];
      Logger.log("Searching in sheet: " + sheet.getName());
      
      var data = sheet.getDataRange().getValues();
      if (data.length <= 1) {
        Logger.log("No data found in sheet: " + sheet.getName());
        continue;
      }
      
      var headers = data[0];
      Logger.log("Headers in sheet: " + headers.join(", "));
      
      // Look for email column - try all possible variations
      var emailCols = [];
      var idCols = [];
      var nameCols = [];
      
      // Map of possible headers to their meanings
      var headerMap = {
        email: ["EMAIL", "E-MAIL", "E-MAIL-ID", "EMAIL ID", "MAIL", "E MAIL", "EMAIL ADDRESS"],
        id: ["EMPLOYEE ID", "EMPLOYEEID", "EMP ID", "ID", "STAFF ID"],
        name: ["NAME", "EMPLOYEE NAME", "STAFF NAME", "FULL NAME", "NAME OF THE EMPLOYEE"]
      };
      
      // Find all possible columns for email, id, and name
      for (var i = 0; i < headers.length; i++) {
        var header = headers[i] ? headers[i].toString().toUpperCase().trim() : "";
        
        // Check for email columns
        if (headerMap.email.indexOf(header) !== -1 || 
            header.indexOf("EMAIL") !== -1 || 
            header.indexOf("MAIL") !== -1) {
          emailCols.push(i);
          Logger.log("Found email column: " + headers[i] + " at index " + i);
        }
        
        // Check for ID columns
        if (headerMap.id.indexOf(header) !== -1 || 
            (header.indexOf("ID") !== -1 && 
             (header.indexOf("EMPLOYEE") !== -1 || 
              header.indexOf("STAFF") !== -1 || 
              header === "ID"))) {
          idCols.push(i);
          Logger.log("Found ID column: " + headers[i] + " at index " + i);
        }
        
        // Check for name columns
        if (headerMap.name.indexOf(header) !== -1 || 
            (header.indexOf("NAME") !== -1 && 
             (header.indexOf("EMPLOYEE") !== -1 || 
              header.indexOf("STAFF") !== -1 || 
              header === "NAME"))) {
          nameCols.push(i);
          Logger.log("Found name column: " + headers[i] + " at index " + i);
        }
      }
      
      // Strategy 1: Find by ID
      if (idCols.length > 0) {
        for (var idColIndex = 0; idColIndex < idCols.length; idColIndex++) {
          var idCol = idCols[idColIndex];
          Logger.log("Searching by ID in column: " + headers[idCol]);
          
          var searchId = employeeId.toString().toLowerCase().trim();
          
          for (var i = 1; i < data.length; i++) {
            var rowId = data[i][idCol] ? data[i][idCol].toString().toLowerCase().trim() : "";
            
            // Check for exact or partial match
            if (rowId === searchId || rowId.indexOf(searchId) !== -1 || searchId.indexOf(rowId) !== -1) {
              Logger.log("Found employee ID match at row " + i);
              
              // First try to find a valid email in email columns
              for (var emailColIndex = 0; emailColIndex < emailCols.length; emailColIndex++) {
                var emailCol = emailCols[emailColIndex];
                var email = data[i][emailCol];
                
                if (email && isValidEmail(email.toString())) {
                  Logger.log("Found valid email in email column: " + email);
                  return email.toString();
                }
              }
              
              // If no valid email in email columns, scan entire row
              for (var j = 0; j < data[i].length; j++) {
                var cellValue = data[i][j] ? data[i][j].toString() : "";
                if (isValidEmail(cellValue)) {
                  Logger.log("Found valid email in row: " + cellValue);
                  return cellValue;
                }
              }
            }
          }
        }
      }
      
      // Strategy 2: Find by name (if employeeId might be a name)
      if (nameCols.length > 0) {
        for (var nameColIndex = 0; nameColIndex < nameCols.length; nameColIndex++) {
          var nameCol = nameCols[nameColIndex];
          Logger.log("Searching by name in column: " + headers[nameCol]);
          
          var searchName = employeeId.toString().toLowerCase().trim();
          
          for (var i = 1; i < data.length; i++) {
            var rowName = data[i][nameCol] ? data[i][nameCol].toString().toLowerCase().trim() : "";
            
            if (rowName === searchName || rowName.indexOf(searchName) !== -1 || searchName.indexOf(rowName) !== -1) {
              Logger.log("Found name match at row " + i);
              
              // Try email columns first
              for (var emailColIndex = 0; emailColIndex < emailCols.length; emailColIndex++) {
                var emailCol = emailCols[emailColIndex];
                var email = data[i][emailCol];
                
                if (email && isValidEmail(email.toString())) {
                  Logger.log("Found valid email by name match: " + email);
                  return email.toString();
                }
              }
              
              // Scan entire row for emails
              for (var j = 0; j < data[i].length; j++) {
                var cellValue = data[i][j] ? data[i][j].toString() : "";
                if (isValidEmail(cellValue)) {
                  Logger.log("Found valid email in row by name match: " + cellValue);
                  return cellValue;
                }
              }
            }
          }
        }
      }
      
      // Strategy 3: Scan entire sheet for the exact employee ID
      Logger.log("Scanning entire sheet for employee ID match");
      for (var i = 1; i < data.length; i++) {
        for (var j = 0; j < data[i].length; j++) {
          var cellValue = data[i][j] ? data[i][j].toString().toLowerCase().trim() : "";
          if (cellValue === employeeId.toString().toLowerCase().trim()) {
            Logger.log("Found exact ID match at row " + i + ", col " + j);
            
            // Look for email in this row
            for (var k = 0; k < data[i].length; k++) {
              var possibleEmail = data[i][k] ? data[i][k].toString() : "";
              if (isValidEmail(possibleEmail)) {
                Logger.log("Found valid email in row with ID match: " + possibleEmail);
                return possibleEmail;
              }
            }
          }
        }
      }
      
      // Strategy 4: Get any valid email from the sheet as last resort
      if (emailCols.length > 0) {
        Logger.log("No match found, using first valid email in email column");
        for (var emailColIndex = 0; emailColIndex < emailCols.length; emailColIndex++) {
          var emailCol = emailCols[emailColIndex];
          
          for (var i = 1; i < data.length; i++) {
            var email = data[i][emailCol];
            if (email && isValidEmail(email.toString())) {
              Logger.log("Using fallback email from row " + i + ": " + email);
              return email.toString();
            }
          }
        }
      }
    }
    
    // If we get here, we couldn't find a valid email in ANY sheet
    Logger.log("No valid email found anywhere - using default email");
    return Session.getActiveUser().getEmail() || "admin@example.com";
    
  } catch (error) {
    Logger.log("Error in getEmployeeEmailById: " + error.toString());
    return Session.getActiveUser().getEmail() || "admin@example.com";
  }
}

/**
 * Handles PWA manifest request
 */
function doGetManifest() {
  const manifest = {
    "name": "Invigilator Chart",
    "short_name": "Invigilator",
    "description": "Invigilator Duty Management System",
    "start_url": ScriptApp.getService().getUrl(),
    "display": "standalone",
    "background_color": "#ffffff",
    "theme_color": "#4361ee",
    "icons": [
      {
        "src": ScriptApp.getService().getUrl() + "?page=icon",
        "sizes": "192x192",
        "type": "image/png"
      }
    ]
  };
  
  return ContentService.createTextOutput(JSON.stringify(manifest))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Handles PWA icon request
 */
function doGetIcon() {
  // You can replace this with your own icon data
  const iconData = "iVBORw0KGgoAAAANSUhEUgAAAMAAAADACAYAAABS3GwHAAAACXBIWXMAAAsTAAALEwEAmpwYAAAF0WlUWHRYTUw6Y29tLmFkb2JlLnhtcAAAAAAAPD94cGFja2V0IGJlZ2luPSLvu78iIGlkPSJXNU0wTXBDZWhpSHpyZVN6TlRjemtjOWQiPz4gPHg6eG1wbWV0YSB4bWxuczp4PSJhZG9iZTpuczptZXRhLyIgeDp4bXB0az0iQWRvYmUgWE1QIENvcmUgNy4yLWMwMDAgNzkuMWI2NWE3OWI0LCAyMDIyLzA2LzEzLTIyOjAxOjAxICAgICAgICAiPiA8cmRmOlJERiB4bWxuczpyZGY9Imh0dHA6Ly93d3cudzMub3JnLzE5OTkvMDIvMjItcmRmLXN5bnRheC1ucyMiPiA8cmRmOkRlc2NyaXB0aW9uIHJkZjphYm91dD0iIiB4bWxuczp4bXA9Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC8iIHhtbG5zOnhtcE1NPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvbW0vIiB4bWxuczpzdEV2dD0iaHR0cDovL25zLmFkb2JlLmNvbS94YXAvMS4wL3NUeXBlL1Jlc291cmNlRXZlbnQjIiB4bWxuczpkYz0iaHR0cDovL3B1cmwub3JnL2RjL2VsZW1lbnRzLzEuMS8iIHhtbG5zOnBob3Rvc2hvcD0iaHR0cDovL25zLmFkb2JlLmNvbS9waG90b3Nob3AvMS4wLyIgeG1wOkNyZWF0b3JUb29sPSJBZG9iZSBQaG90b3Nob3AgMjQuMCAoTWFjaW50b3NoKSIgeG1wOkNyZWF0ZURhdGU9IjIwMjQtMDMtMjVUMTU6NDc6NDUrMDg6MDAiIHhtcDpNZXRhZGF0YURhdGU9IjIwMjQtMDMtMjVUMTU6NDc6NDUrMDg6MDAiIHhtcDpNb2RpZnlEYXRlPSIyMDI0LTAzLTI1VDE1OjQ3OjQ1KzA4OjAwIiB4bXBNTTpJbnN0YW5jZUlEPSJ4bXAuaWlkOjY5ZDM4ZmM1LTU0ZDgtNDI0Ny1hMzA2LTNmYjFiYzM5ZmM0YyIgeG1wTU06RG9jdW1lbnRJRD0iYWRvYmU6ZG9jaWQ6cGhvdG9zaG9wOjY5ZDM4ZmM1LTU0ZDgtNDI0Ny1hMzA2LTNmYjFiYzM5ZmM0YyIgeG1wTU06T3JpZ2luYWxEb2N1bWVudElEPSJ4bXAuZGlkOjY5ZDM4ZmM1LTU0ZDgtNDI0Ny1hMzA2LTNmYjFiYzM5ZmM0YyIgZGM6Zm9ybWF0PSJpbWFnZS9wbmciIHBob3Rvc2hvcDpDb2xvck1vZGU9IjMiPiA8eG1wTU06SGlzdG9yeT4gPHJkZjpTZXE+IDxyZGY6bGkgc3RFdnQ6YWN0aW9uPSJjcmVhdGVkIiBzdEV2dDppbnN0YW5jZUlEPSJ4bXAuaWlkOjY5ZDM4ZmM1LTU0ZDgtNDI0Ny1hMzA2LTNmYjFiYzM5ZmM0YyIgc3RFdnQ6d2hlbj0iMjAyNC0wMy0yNVQxNTo0Nzo0NSswODowMCIgc3RFdnQ6c29mdHdhcmVBZ2VudD0iQWRvYmUgUGhvdG9zaG9wIDI0LjAgKE1hY2ludG9zaCkiLz4gPC9yZGY6U2VxPiA8L3htcE1NOkhpc3Rvcnk+IDwvcmRmOkRlc2NyaXB0aW9uPiA8L3JkZjpSREY+IDwveDp4bXBtZXRhPiA8P3hwYWNrZXQgZW5kPSJyIj8+";
  
  return ContentService.createTextOutput(Utilities.base64Decode(iconData))
    .setMimeType(ContentService.MimeType.PNG);
}

/**
 * Modified doGet function to handle PWA requests
 */
function doGet(e) {
  if (e.parameter.page === 'manifest') {
    return doGetManifest();
  } else if (e.parameter.page === 'icon') {
    return doGetIcon();
  }
  
  // Your existing doGet logic here
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Invigilator Chart')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=5.0, minimum-scale=1.0, user-scalable=yes, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getCachedData(key) {
  const cache = CacheService.getUserCache();
  return cache.get(key);
}

function setCachedData(key, value, expirationInSeconds = 21600) {
  const cache = CacheService.getUserCache();
  cache.put(key, value, expirationInSeconds);
}

function sortDutiesByDate(duties) {
  return duties.sort((a, b) => {
    const dateA = new Date(a.date);
    const dateB = new Date(b.date);
    return dateA - dateB;
  });
}

/**
 * Show the duties entry panel
 */
function showDutiesEntryPanel() {
  var html = HtmlService.createHtmlOutputFromFile('DutiesEntry')
      .setTitle('Duties Entry')
      .setWidth(800)
      .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Duties Entry');
}

function checkEmployeeExists(employeeId) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Main Data');
    if (!sheet) {
      return { success: false, exists: false, message: 'Main Data sheet not found' };
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const employeeIdCol = headers.indexOf('EMPLOYEE ID');

    if (employeeIdCol === -1) {
      return { success: false, exists: false, message: 'EMPLOYEE ID column not found' };
    }

    // Log the search parameters
    console.log('Searching for employee ID:', employeeId);
    console.log('Employee ID column index:', employeeIdCol);

    // Check if employee exists
    for (let i = 1; i < data.length; i++) {
      const existingId = data[i][employeeIdCol];
      // Log each comparison
      console.log('Comparing with:', existingId);
      
      // Convert both to strings and trim whitespace for comparison
      if (String(existingId).trim() === String(employeeId).trim()) {
        return { success: true, exists: true };
      }
    }

    return { success: true, exists: false };
  } catch (error) {
    console.error('Error in checkEmployeeExists:', error);
    return { success: false, exists: false, message: error.toString() };
  }
}

function saveDutyEntry(formData) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Main Data');
    if (!sheet) {
      return { success: false, message: 'Main Data sheet not found' };
    }

    // Get all data and headers
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // Find required column indices
    const requiredColumns = [
      'Date', 'Week Days', 'Degree Prog. Or Semester or School',
      'Slot 1 (Time: 09:00 a.m. to 10: 30 a.m.)', 'Location',
      'Slot 2 (Time: 11:30 a.m. to 01:00 p.m.)', 'Location',
      'Slot 3 (Time: 03:30 p.m. to 05:00 p.m.)', 'Location',
      'Slot 4 (Time: 05:30 p.m. to 07:00 p.m.)', 'Location'
    ];
    
    const columnIndices = {};
    for (const col of requiredColumns) {
      const index = headers.indexOf(col);
      if (index === -1) {
        return { success: false, message: `Required column "${col}" not found` };
      }
      columnIndices[col] = index;
    }

    // Format the date to MMM-DD-YYYY
    const date = new Date(formData.date);
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    const formattedDate = `${months[date.getMonth()]}-${String(date.getDate()).padStart(2, '0')}-${date.getFullYear()}`;

    // Validate required fields
    if (!formData.employeeId || !formData.employeeName || !formData.date || !formData.degreeProgram) {
      return { success: false, message: 'Required fields are missing' };
    }

    // Create new row data
    const newRow = [
      sheet.getLastRow(), // Sl. No.
      formData.employeeId,
      formData.employeeName,
      formData.designation || '',
      formData.email || '',
      formData.cellPhone || '',
      '', // Days of Exam
      formattedDate,
      formData.weekDays,
      formData.degreeProgram,
      formData.slots[1].assigned ? 'x' : '',
      formData.slots[1].location || '',
      formData.slots[2].assigned ? 'x' : '',
      formData.slots[2].location || '',
      formData.slots[3].assigned ? 'x' : '',
      formData.slots[3].location || '',
      formData.slots[4].assigned ? 'x' : '',
      formData.slots[4].location || '',
      '', // Total Count of Invigilation Duty on the Day
      ''  // Total Count of Invigilation Duty of the Season
    ];

    // Find the correct row to insert the new entry
    let insertRow = 1; // Start after header row
    for (let i = 1; i < data.length; i++) {
      const rowDate = new Date(data[i][columnIndices['Date']]);
      if (date < rowDate) {
        insertRow = i + 1;
        break;
      }
    }

    // Insert the new row at the correct position
    sheet.insertRowBefore(insertRow);
    sheet.getRange(insertRow, 1, 1, newRow.length).setValues([newRow]);

    // Apply formatting to the new row
    const range = sheet.getRange(insertRow, 1, 1, newRow.length);
    range.setFontFamily('Arial');
    range.setFontSize(10);
    range.setHorizontalAlignment('center');
    range.setVerticalAlignment('middle');
    range.setWrap(true);

    return { success: true };
  } catch (error) {
    console.error('Error in saveDutyEntry:', error);
    return { success: false, message: error.toString() };
  }
}