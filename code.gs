function onInstall() {
  onOpen();
}
function onOpen() {
  var ui = SpreadsheetApp.getUi();
    
    ui.createMenu("🌐 DataMate")
    .addSubMenu(
      ui.createMenu("🌐 DataMate")
        .addItem("💾 Save Record", "save")
        .addItem("🔄 Reset Input", "copyInput1")
        .addItem("👁️ Reset View/Print", "view")
        .addItem("📄 New Dataset", "newfile")
        .addItem("📄 Add Input Sheet", "duplicateAndRenameSheet")
        .addItem('📷 SnapSync', 'showCameraSidebar')
    )
    .addSeparator()
    .addSubMenu(
      ui.createMenu("📝 FormBuilder")
        .addItem("👀 Preview Form", "previewForm")
        .addItem("🛠️ Form Builder", "showFormBuilder")
        .addItem('📋 Dynamic Entry Form', 'showDynamicForm')
    )
    .addSeparator()
    .addSubMenu(
      ui.createMenu("📇 AddressBlock")
        .addItem("📧 Mail It", "showMailItSidebar")       
        .addItem("📋 Add Contact Sheets", "contacts")
        .addItem("📥 Import Gmail™ Contacts", "showUploadDialog")
        .addItem("➕ New Contact", "newcontact")
        .addItem("✏️ Edit Name", "EditAddressSheet")
        .addItem("🏢 Edit Company", "EditAddressSheet1")
    )
    
    .addSeparator()
    .addSubMenu(
      ui.createMenu("📋 Templates")
        .addItem("📦 Inventory", "setup")
        .addItem("🔧 Save Invoice", "saveAndUpdate")
        .addItem("⏰ Weekly Timesheets", "setupTS")
        .addItem("💸 Submit Timesheet", "saveAndCopyWithTempRename")
        .addItem("🛒 Purchase Order", "setupPO")
    )
    .addToUi();
}




function showDynamicForm() {
  const html = HtmlService.createHtmlOutputFromFile('DynamicForm')
    .setTitle('Dynamic Data Entry Form');
  SpreadsheetApp.getUi().showSidebar(html);
  createDropdownSheet();
}

function doGet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const user = Session.getActiveUser().getEmail();
    const owner = ss.getOwner().getEmail();
    const editors = ss.getEditors().map(e => e.getEmail());
    if (user !== owner && !editors.includes(user)) {
      Logger.log(`User ${user} lacks edit access to spreadsheet`);
      return HtmlService.createHtmlOutput(`
        <h3>Access Denied</h3>
        <p>You do not have permission to edit this spreadsheet. Please contact the owner (${owner}) to request edit access.</p>
      `).setTitle('Dynamic Data Entry Form');
    }
  } catch (e) {
    Logger.log(`Error checking permissions: ${e.message}`);
    return HtmlService.createHtmlOutput(`
      <h3>Error</h3>
      <p>Unable to verify permissions: ${e.message}. Please ensure you are logged in and have edit access to the spreadsheet.</p>
    `).setTitle('Dynamic Data Entry Form');
  }

  const htmlContent = `
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body { font-family: Arial, sans-serif; padding: 20px; }
        label { display: block; margin: 10px 0 5px; }
        input, select { width: 100%; padding: 8px; margin-bottom: 10px; }
        button { padding: 10px; margin: 5px; }
        #message { color: green; margin-top: 10px; }
        .error { color: red; }
        #spinner { display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.5); }
        #spinner div { position: absolute; top: 50%; left: 50%; transform: translate(-50%,-50%); color: white; }
        #searchContainer { margin-bottom: 20px; }
      </style>
    </head>
    <body>
      <div id="searchContainer">
        <label for="searchId">Search by ID</label>
        <select id="searchId">
          <option value="">Select ID</option>
        </select>
        <button type="button" onclick="searchRecord()">Search</button>
      </div>
      <form id="dynamicForm">
        <div id="formFields"></div>
        <button type="button" onclick="saveRecord()">Save</button>
        <button type="button" onclick="clearForm()">New</button>
        <button type="button" onclick="navigate('prev')">Previous</button>
        <button type="button" onclick="navigate('next')">Next</button>
      </form>
      <div id="message"></div>
      <div id="spinner"><div>Loading...</div></div>
      <script>
        let headers = [];
        let records = [];
        let currentIndex = -1;
        let isNewRecord = false;
        let currentRowNumber = null;

        google.script.run
          .withSuccessHandler(populateForm)
          .withFailureHandler(err => {
            console.error('Error loading headers: ' + err.message);
            showMessage('Error loading form fields: ' + err.message, 'error');
          })
          .getSheetInfo();
        google.script.run
          .withSuccessHandler(loadRecords)
          .withFailureHandler(err => {
            console.error('Error loading records: ' + err.message);
            showMessage('Error loading records: ' + err.message, 'error');
          })
          .getVisibleRecords();
        google.script.run
          .withSuccessHandler(populateIdDropdown)
          .withFailureHandler(err => {
            console.error('Error loading ID dropdown: ' + err.message);
            showMessage('Error loading ID dropdown: ' + err.message, 'error');
          })
          .getColumnAValues();

        function populateForm(headerData) {
          if (!headerData || headerData.length === 0) {
            console.error('No headers received from getSheetInfo');
            showMessage('No fields found in the target sheet. Please check the sheet headers.', 'error');
            return;
          }
          headers = headerData;
          const formFields = document.getElementById('formFields');
          formFields.innerHTML = headers.map(header => {
            if (header.name === 'ID') {
              return '<label for="' + header.name + '">' + header.name + '</label>' +
                     '<input type="number" id="' + header.name + '" readonly>';
            } else if (header.type === 'select') {
              if (!header.options || header.options.length === 0) {
                console.warn(\`No options available for dropdown \${header.name}\`);
                showMessage(\`Warning: No options for \${header.name} dropdown. Check sheet configuration.\`, 'error');
                return '<label for="' + header.name + '">' + header.name + (header.required ? ' *' : '') + '</label>' +
                       '<select id="' + header.name + '" ' + (header.required ? 'required' : '') + ' onchange="onDropdownChange()">' +
                       '<option value="">Select ' + header.name + '</option>' +
                       '</select>';
              }
              console.log(\`Populating dropdown \${header.name} with options: \${header.options.join(', ')}\`);
              return '<label for="' + header.name + '">' + header.name + (header.required ? ' *' : '') + '</label>' +
                     '<select id="' + header.name + '" ' + (header.required ? 'required' : '') + ' onchange="onDropdownChange()">' +
                     '<option value="">Select ' + header.name + '</option>' +
                     header.options.map(opt => '<option value="' + opt + '">' + opt + '</option>').join('') +
                     '</select>';
            } else {
              return '<label for="' + header.name + '">' + header.name + (header.required ? ' *' : '') + '</label>' +
                     '<input type="' + header.type + '" id="' + header.name + '" ' + (header.required ? 'required' : '') + '>';
            }
          }).join('');
          console.log('Form fields populated with headers: ' + headers.map(h => h.name).join(', '));
          showMessage('Form fields loaded successfully.', '');
        }

        function populateIdDropdown(idValues) {
          const searchIdDropdown = document.getElementById('searchId');
          searchIdDropdown.innerHTML = '<option value="">Select ID</option>';
          if (idValues.length === 0) {
            console.warn('No IDs found in column A');
            showMessage('No IDs found in column A of the target sheet.', 'error');
          } else {
            idValues.forEach(id => {
              if (id !== '') {
                const option = document.createElement('option');
                option.value = id;
                option.textContent = id;
                searchIdDropdown.appendChild(option);
              }
            });
            console.log('ID dropdown populated with: ' + idValues.join(', '));
          }
        }

        function loadRecords(data) {
          records = data;
          if (records.length > 0) {
            currentIndex = 0;
            currentRowNumber = currentIndex + 2;
            isNewRecord = false;
            displayRecord();
          } else {
            clearForm();
            console.warn('No records loaded');
            showMessage('No records found in the target sheet.', 'error');
          }
          google.script.run.withSuccessHandler(populateIdDropdown).getColumnAValues();
        }

        function displayRecord() {
          isNewRecord = false;
          if (currentIndex >= 0 && currentIndex < records.length) {
            headers.forEach(header => {
              const field = document.getElementById(header.name);
              if (!field) {
                console.error(\`Field \${header.name} not found in form\`);
                showMessage(\`Error: Field \${header.name} not found in form.\`, 'error');
                return;
              }
              const value = records[currentIndex][header.name] != null ? records[currentIndex][header.name] : '';
              if (field.tagName === 'SELECT') {
                let exists = Array.from(field.options).some(opt => opt.value === String(value));
                if (!exists && value !== '') {
                  const opt = document.createElement('option');
                  opt.value = String(value);
                  opt.textContent = String(value);
                  field.appendChild(opt);
                  console.log(\`Added option \${value} to dropdown \${header.name}\`);
                }
                field.value = String(value);
              } else {
                field.value = value;
              }
            });
            const currentId = records[currentIndex]['ID'] != null ? String(records[currentIndex]['ID']) : '';
            document.getElementById('searchId').value = currentId;
            console.log(\`Displayed record at index \${currentIndex}: ID=\${currentId}, Row=\${currentRowNumber}, Data=\${JSON.stringify(records[currentIndex])}\`);
          } else {
            console.error(\`Invalid currentIndex: \${currentIndex}, records length: \${records.length}\`);
            showMessage('Error: No record to display.', 'error');
          }
        }

        function searchRecord() {
          const searchId = document.getElementById('searchId').value;
          if (!searchId) {
            console.warn('No ID selected for search');
            showMessage('Please select an ID to search.', 'error');
            return;
          }
          console.log(\`Initiating search for ID: \${searchId}\`);
          document.getElementById('spinner').style.display = 'block';
          google.script.run
            .withSuccessHandler(result => {
              document.getElementById('spinner').style.display = 'none';
              if (result && result.record && Object.keys(result.record).length > 0) {
                console.log(\`Record found for ID \${searchId}: \${JSON.stringify(result.record)}, Row=\${result.rowNumber}\`);
                records = records.filter(r => String(r.ID) !== String(searchId));
                records.push(result.record);
                currentIndex = records.length - 1;
                currentRowNumber = result.rowNumber;
                isNewRecord = false;
                displayRecord();
                showMessage('Record found and displayed.', '');
              } else {
                console.error(\`No record found for ID: \${searchId}\`);
                showMessage(\`No record found with ID: \${searchId}\`, 'error');
              }
            })
            .withFailureHandler(err => {
              document.getElementById('spinner').style.display = 'none';
              console.error(\`Search failed for ID \${searchId}: \${err.message}\`);
              showMessage('Error searching record: ' + err.message, 'error');
            })
            .getRecordById(searchId);
        }

        function saveRecord() {
          console.log(\`Saving record: isNewRecord=\${isNewRecord}, currentRowNumber=\${currentRowNumber}\`);
          document.getElementById('spinner').style.display = 'block';
          const formData = {};
          headers.forEach(header => {
            const field = document.getElementById(header.name);
            if (!field) {
              console.error(\`Field \${header.name} not found during save\`);
              showMessage(\`Error: Field \${header.name} not found.\`, 'error');
              document.getElementById('spinner').style.display = 'none';
              return;
            }
            formData[header.name] = field.value.trim();
          });

          for (const header of headers) {
            if (header.required && !formData[header.name]) {
              console.warn(\`Required field \${header.name} is empty\`);
              showMessage('Please fill all required fields.', 'error');
              document.getElementById('spinner').style.display = 'none';
              return;
            }
          }

          if (isNewRecord) {
            console.log('Calling addRecord for new record');
            google.script.run
              .withSuccessHandler(result => onSave(result, null))
              .withFailureHandler(onError)
              .addRecord(formData);
          } else if (currentRowNumber && currentRowNumber > 1) {
            console.log(\`Calling updateRecord for row \${currentRowNumber}\`);
            formData._rowNumber = currentRowNumber;
            google.script.run
              .withSuccessHandler(result => onSave(result, currentRowNumber))
              .withFailureHandler(onError)
              .updateRecord(formData);
          } else {
            console.error('No valid row number for updating record');
            showMessage('Error: Cannot update record, no valid row number available.', 'error');
            document.getElementById('spinner').style.display = 'none';
          }
        }

        function clearForm() {
          isNewRecord = true;
          currentRowNumber = null;
          document.getElementById('dynamicForm').reset();
          headers.forEach(header => {
            if (header.name === 'ID') return;
            const field = document.getElementById(header.name);
            if (field) field.value = '';
          });
          document.getElementById('searchId').value = '';
          console.log('Form cleared for new record');
          showMessage('Ready for new record.', '');
        }

        function navigate(direction) {
          if (records.length === 0) {
            console.warn('No records to navigate');
            return;
          }
          if (direction === 'prev' && currentIndex > 0) {
            currentIndex--;
            currentRowNumber = currentIndex + 2;
          } else if (direction === 'next' && currentIndex < records.length - 1) {
            currentIndex++;
            currentRowNumber = currentIndex + 2;
          }
          console.log(\`Navigating to \${direction}, new index: \${currentIndex}, row: \${currentRowNumber}\`);
          displayRecord();
        }

        function onDropdownChange() {
          if (isNewRecord) return;
          const currentID = document.getElementById('ID').value;
          if (!currentID) {
            console.warn('No ID for dropdown change refresh');
            return;
          }
          console.log(\`Refreshing record for ID: \${currentID} due to dropdown change\`);
          document.getElementById('spinner').style.display = 'block';
          google.script.run
            .withSuccessHandler(result => {
              document.getElementById('spinner').style.display = 'none';
              if (result && result.record) {
                headers.forEach(header => {
                  const field = document.getElementById(header.name);
                  if (!field) {
                    console.error(\`Field \${header.name} not found on dropdown change\`);
                    return;
                  }
                  const val = result.record[header.name] != null ? result.record[header.name] : '';
                  if (field.tagName === 'SELECT') {
                    let exists = Array.from(field.options).some(opt => opt.value === String(val));
                    if (!exists && val !== '') {
                      const opt = document.createElement('option');
                      opt.value = String(val);
                      opt.textContent = String(val);
                      field.appendChild(opt);
                      console.log(\`Added option \${val} to dropdown \${header.name}\`);
                    }
                    field.value = String(val);
                  } else {
                    field.value = val;
                  }
                });
                document.getElementById('searchId').value = result.record['ID'] != null ? String(result.record['ID']) : '';
                currentRowNumber = result.rowNumber;
                console.log(\`Refreshed record for ID \${currentID}: \${JSON.stringify(result.record)}, Row=\${result.rowNumber}\`);
                showMessage('Record refreshed with formula updates.', '');
              } else {
                console.error(\`No record found for ID \${currentID} on refresh\`);
                showMessage('Record not found on reload.', 'error');
              }
            })
            .withFailureHandler(err => {
              document.getElementById('spinner').style.display = 'none';
              console.error(\`Error refreshing record for ID \${currentID}: \${err.message}\`);
              showMessage('Error refreshing record: ' + err.message, 'error');
            })
            .getRecordById(currentID);
        }

        function onSave(result, existingRow) {
          document.getElementById('spinner').style.display = 'none';
          if (result.status === 'success') {
            console.log(\`Record saved: \${JSON.stringify(result)}\`);
            showMessage('Record saved successfully.', '');
            google.script.run
              .withSuccessHandler(data => {
                records = data;
                if (existingRow) {
                  const updatedID = document.getElementById('ID').value;
                  const idx = records.findIndex(r => String(r.ID) === String(updatedID));
                  if (idx >= 0) {
                    currentIndex = idx;
                    currentRowNumber = idx + 2;
                    displayRecord();
                  } else {
                    currentIndex = records.length - 1;
                    currentRowNumber = currentIndex + 2;
                    displayRecord();
                  }
                } else {
                  currentIndex = records.length - 1;
                  currentRowNumber = currentIndex + 2;
                  displayRecord();
                }
                google.script.run.withSuccessHandler(populateIdDropdown).getColumnAValues();
              })
              .withFailureHandler(err => {
                console.error('Error reloading records after save: ' + err.message);
                showMessage('Error reloading records: ' + err.message, 'error');
              })
              .getVisibleRecords();
            isNewRecord = false;
            document.getElementById('searchId').value = '';
          } else {
            console.error(\`Save failed: \${result.message}\`);
            showMessage(result.message || 'Error saving record.', 'error');
          }
        }

        function onError(error) {
          document.getElementById('spinner').style.display = 'none';
          console.error('Operation error: ' + error.message);
          showMessage('Error: ' + error.message, 'error');
        }

        function showMessage(message, className) {
          const msgDiv = document.getElementById('message');
          msgDiv.textContent = message;
          msgDiv.className = className;
          setTimeout(() => msgDiv.textContent = '', 3000);
        }
      </script>
    </body>
    </html>
  `;
  return HtmlService.createHtmlOutput(htmlContent)
    .setTitle('Dynamic Data Entry Form')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function createDropdownSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName("Dropdowns")) {
    const newSheet = ss.insertSheet("Dropdowns");
    newSheet.getRange("A1").setValue("Dropdown");
    newSheet.getRange("B1").setValue("Options");
    newSheet.getRange("C1").setValue("Target Sheet");
    newSheet.getRange("C2").setValue(ss.getActiveSheet().getName());
    Logger.log("Created Dropdowns sheet with default target: " + ss.getActiveSheet().getName());
  }
}

function getDropdownOptions(dropdownsSheet) {
  const data = dropdownsSheet.getDataRange().getValues();
  const options = {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  Logger.log(`Processing Dropdowns sheet with ${data.length - 1} entries`);
  for (let i = 1; i < data.length; i++) {
    const key = data[i][0]?.toString().trim();
    const value = data[i][1]?.toString().trim();
    if (!key || !value) {
      Logger.log(`Skipping invalid dropdown entry at row ${i + 1}: key=${key}, value=${value}`);
      continue;
    }
    Logger.log(`Processing dropdown: ${key}, options: ${value}`);
    if (value.includes('!')) {
      try {
        const [sheetName, colRange] = value.split('!');
        Logger.log(`Attempting to access sheet: ${sheetName}, range: ${colRange}`);
        const sourceSheet = ss.getSheetByName(sheetName);
        if (!sourceSheet) {
          Logger.log(`Error: Source sheet ${sheetName} for dropdown ${key} not found`);
          continue;
        }
        const lastRow = sourceSheet.getLastRow();
        if (lastRow < 1) {
          Logger.log(`Error: No data in sheet ${sheetName} for dropdown ${key}`);
          continue;
        }
        const range = sourceSheet.getRange(colRange);
        if (!range || range.isBlank()) {
          Logger.log(`Error: Range ${colRange} in sheet ${sheetName} is invalid or empty for dropdown ${key}`);
          continue;
        }
        const values = sourceSheet.getRange(1, range.getColumn(), lastRow, 1).getValues().flat().filter(v => v != null && v.toString().trim() !== '');
        if (values.length === 0) {
          Logger.log(`Warning: No valid values in ${sheetName}!${colRange} for dropdown ${key}`);
          continue;
        }
        options[key] = [...new Set(values.map(v => v.toString().trim()))];
        Logger.log(`Dropdown options for ${key} from ${sheetName}!${colRange}: ${JSON.stringify(options[key])}`);
      } catch (e) {
        Logger.log(`Error parsing dropdown options for ${key} from ${value}: ${e.message}`);
        continue;
      }
    } else {
      options[key] = value.split(',').map(opt => opt.trim()).filter(opt => opt !== '');
      Logger.log(`Dropdown options for ${key}: ${JSON.stringify(options[key])}`);
    }
  }
  Logger.log(`Final dropdown options: ${JSON.stringify(options)}`);
  return options;
}

function protectAllFormulaCells() {
  const sheet = getTargetSheet();
  const range = sheet.getDataRange();
  const formulas = range.getFormulas();

  for (let r = 0; r < formulas.length; r++) {
    for (let c = 0; c < formulas[r].length; c++) {
      if (formulas[r][c]) {
        const cell = sheet.getRange(r + 1, c + 1);
        try {
          const protection = cell.protect();
          protection.setDescription('Formula cell - do not edit');
          protection.removeEditors(protection.getEditors());
          Logger.log(`Protected formula cell at ${sheet.getName()}!${cell.getA1Notation()}`);
        } catch (e) {
          Logger.log(`Error protecting cell ${sheet.getName()}!${cell.getA1Notation()}: ${e.message}`);
        }
      }
    }
  }
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'All formula cells have been protected.',
    'Done',
    3
  );
}

function getTargetSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dropdownsSheet = ss.getSheetByName("Dropdowns");
  let targetSheetName = ss.getActiveSheet().getName();
  if (dropdownsSheet) {
    const targetCell = dropdownsSheet.getRange("C2").getValue()?.toString().trim();
    if (targetCell && ss.getSheetByName(targetCell)) {
      targetSheetName = targetCell;
      Logger.log(`Target sheet set to: ${targetSheetName}`);
    } else {
      Logger.log(`Invalid or missing sheet name in Dropdowns!C2: ${targetCell}. Falling back to active sheet: ${targetSheetName}`);
    }
  } else {
    Logger.log("Dropdowns sheet not found. Falling back to active sheet: " + targetSheetName);
  }
  const targetSheet = ss.getSheetByName(targetSheetName);
  if (!targetSheet) {
    Logger.log(`Error: Target sheet ${targetSheetName} does not exist`);
    throw new Error(`Target sheet ${targetSheetName} does not exist`);
  }
  return targetSheet;
}

function getSheetInfo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getTargetSheet();
  const lastColumn = sheet.getLastColumn();
  if (lastColumn === 0) {
    Logger.log(`Error: No headers found in target sheet ${sheet.getName()}`);
    throw new Error(`No headers found in target sheet ${sheet.getName()}`);
  }
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].filter(h => h != null && h.toString().trim() !== '');
  const validations = lastColumn > 0 ? sheet.getRange(2, 1, 1, lastColumn).getDataValidations()[0] : [];
  const dropdownsSheet = ss.getSheetByName('Dropdowns');
  const dropdownOptions = dropdownsSheet ? getDropdownOptions(dropdownsSheet) : {};

  Logger.log(`Headers found in ${sheet.getName()}: ${JSON.stringify(headers)}`);
  Logger.log(`Dropdown options from Dropdowns sheet: ${JSON.stringify(dropdownOptions)}`);

  if (headers.length === 0) {
    Logger.log(`Error: No valid headers found in target sheet ${sheet.getName()}`);
    throw new Error(`No valid headers found in target sheet ${sheet.getName()}`);
  }

  return headers.map((header, index) => {
    const headerStr = header.toString().trim();
    const validation = validations[index] || null;
    let type = 'text';
    let options = [];

    if (validation && validation.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
      type = 'select';
      options = validation.getCriteriaValues().filter(v => v != null && v.toString().trim() !== '');
      Logger.log(`Data validation for ${headerStr}: ${JSON.stringify(options)}`);
    }
    const dropdownKey = Object.keys(dropdownOptions).find(key => key.toLowerCase() === headerStr.toLowerCase());
    if (dropdownKey && dropdownOptions[dropdownKey]) {
      type = 'select';
      options = dropdownOptions[dropdownKey];
      Logger.log(`Dropdown options for ${headerStr} from Dropdowns sheet: ${JSON.stringify(options)}`);
    }
    if (headerStr === 'ID') {
      type = 'number';
      options = [];
    }
    Logger.log(`Header ${headerStr}: type=${type}, options=${JSON.stringify(options)}`);
    return {
      name: headerStr,
      type: type,
      options: options,
      required: headerStr !== 'ID',
      columnIndex: index + 1
    };
  });
}

function getColumnAValues() {
  const sheet = getTargetSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    Logger.log(`No data rows in ${sheet.getName()}, returning empty array for column A`);
    return [];
  }
  try {
    const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    const uniqueValues = [...new Set(values.filter(v => v != null && v.toString().trim() !== '').map(v => v.toString().trim()))];
    Logger.log(`Column A values in ${sheet.getName()}: ${JSON.stringify(uniqueValues)}`);
    return uniqueValues;
  } catch (e) {
    Logger.log(`Error getting column A values: ${e.message}`);
    return [];
  }
}

function getVisibleRecords() {
  const sheet = getTargetSheet();
  const filter = sheet.getFilter();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const records = [];

  if (headers.length === 0) {
    Logger.log(`No headers in ${sheet.getName()}, returning empty records`);
    return records;
  }

  if (filter) {
    for (let i = 1; i < data.length; i++) {
      if (!sheet.isRowHiddenByFilter(i + 1)) {
        records.push(data[i]);
      }
    }
  } else {
    records.push(...data.slice(1));
  }
  Logger.log(`Visible records in ${sheet.getName()}: ${records.length}`);
  return records.map((row, index) => {
    return headers.reduce((obj, header, i) => {
      obj[header] = row[i];
      return obj;
    }, {});
  });
}

function addRecord(formData) {
  const sheet = getTargetSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  try {
    const lastId = sheet.getLastRow() > 1 ? Number(sheet.getRange(sheet.getLastRow(), 1).getValue()) || 0 : 0;
    const newId = lastId + 1;

    const row = headers.map(header => header === 'ID' ? newId : formData[header] || '');
    sheet.appendRow(row);
    Logger.log(`Added record with ID ${newId} to ${sheet.getName()}`);
    return { status: 'success', id: newId };
  } catch (e) {
    Logger.log(`Error adding record: ${e.message}`);
    return { status: 'error', message: `Failed to add record: ${e.message}` };
  }
}

function updateRecord(formData) {
  const sheet = getTargetSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  if (!formData._rowNumber || formData._rowNumber <= 1) {
    Logger.log(`Error: Invalid row number ${formData._rowNumber} for update in ${sheet.getName()}`);
    return { status: 'error', message: 'Invalid row number' };
  }

  try {
    const existingRowValues = sheet.getRange(formData._rowNumber, 1, 1, headers.length).getValues()[0];
    const existingRowFormulas = sheet.getRange(formData._rowNumber, 1, 1, headers.length).getFormulas()[0];

    const updatedRow = headers.map((header, idx) => {
      if (existingRowFormulas[idx]) {
        return existingRowFormulas[idx];
      } else {
        return (formData[header] !== '' && formData[header] !== undefined)
          ? formData[header]
          : existingRowValues[idx];
      }
    });

    sheet.getRange(formData._rowNumber, 1, 1, headers.length).setValues([updatedRow]);
    Logger.log(`Updated record at row ${formData._rowNumber} in ${sheet.getName()}`);
    return { status: 'success', row: formData._rowNumber };
  } catch (e) {
    Logger.log(`Error updating record at row ${formData._rowNumber}: ${e.message}`);
    return { status: 'error', message: `Failed to update record: ${e.message}` };
  }
}

function deleteRecord(id) {
  const sheet = getTargetSheet();
  const data = sheet.getDataRange().getValues();

  try {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.deleteRow(i + 1);
        Logger.log(`Deleted record with ID ${id} from ${sheet.getName()}`);
        return { status: 'success' };
      }
    }
    Logger.log(`No record found to delete with ID ${id} in ${sheet.getName()}`);
    return { status: 'error', message: 'Record not found' };
  } catch (e) {
    Logger.log(`Error deleting record with ID ${id}: ${e.message}`);
    return { status: 'error', message: `Failed to delete record: ${e.message}` };
  }
}

function getRecordById(id) {
  const sheet = getTargetSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const data = sheet.getDataRange().getValues();

  Logger.log(`Searching for ID ${id} in ${sheet.getName()}`);
  for (let i = 1; i < data.length; i++) {
    const sheetId = String(data[i][0]).trim();
    const searchId = String(id).trim();
    if (sheetId === searchId) {
      const record = {};
      headers.forEach((header, idx) => {
        record[header] = data[i][idx];
      });
      Logger.log(`Record found for ID ${id}: ${JSON.stringify(record)}, Row=${i + 1}`);
      return { record: record, rowNumber: i + 1 };
    }
  }
  Logger.log(`No record found for ID ${id} in ${sheet.getName()}`);
  return { record: null, rowNumber: null };
}

function createDropdownSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName("Dropdowns")) {
    const newSheet = ss.insertSheet("Dropdowns");
    newSheet.getRange("A1").setValue("Dropdown");
    newSheet.getRange("B1").setValue("Options");
    newSheet.getRange("C1").setValue("Target Sheet");
    newSheet.getRange("C2").setValue(ss.getActiveSheet().getName());
  }
}





// Show the Mail It sidebar
function showMailItSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('MailIt')
    .setTitle('Mail It');
  SpreadsheetApp.getUi().showSidebar(html);
}

// Get sheet names and contact emails for the form
function getSpreadsheetData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    const contactsSheet = ss.getSheetByName("contacts");
    
    if (!contactsSheet) {
      throw new Error("Sheet 'contacts' not found.");
    }

    // Get sheet names
    const sheetNames = sheets.map(sheet => sheet.getName());

    // Get contact emails and names
    const lastRow = contactsSheet.getLastRow();
    let contacts = [];
    if (lastRow >= 2) {
      const range = contactsSheet.getRange("A2:P" + lastRow);
      const values = range.getValues();
      contacts = values
        .map(row => ({
          name: row[0], // Column A (Full Name)
          email: row[15] // Column P (E-mail Address)
        }))
        .filter(contact => contact.email && contact.email.match(/^[^\s@]+@[^\s@]+\.[^\s@]+$/))
        .map(contact => ({ name: contact.name, email: contact.email }));
    }

    return { sheetNames, contacts };
  } catch (error) {
    throw new Error(`Error fetching data: ${error.message}`);
  }
}

// Process the email form submission
function processEmailForm(formData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const { sheetName, rangeAddress, contactSelection, selectedEmails, subject } = formData;

    // Validate inputs
    if (!sheetName || !rangeAddress || !subject) {
      throw new Error("All fields are required.");
    }

    if (!contactSelection || (contactSelection === "specific" && (!selectedEmails || selectedEmails.length === 0))) {
      throw new Error("No contacts selected.");
    }

    // Validate range format
    if (!rangeAddress.match(/^[A-Z]+[0-9]+:[A-Z]+[0-9]+$/)) {
      throw new Error("Invalid range format. Use format like 'A1:G48'.");
    }

    const targetSheet = ss.getSheetByName(sheetName);
    if (!targetSheet) {
      throw new Error(`Sheet '${sheetName}' not found.`);
    }

    // Get the range
    const range = targetSheet.getRange(rangeAddress);
    const values = range.getValues();
    const backgrounds = range.getBackgrounds();
    const fontWeights = range.getFontWeights();
    const fontColors = range.getFontColors();
    const fontStyles = range.getFontStyles();
    const numRows = values.length;
    const numCols = values[0].length;

    // Track merged cells
    const mergedRanges = range.getMergedRanges();
    const mergeMap = {};

    mergedRanges.forEach(mr => {
      const startRow = mr.getRow() - range.getRow() + 1;
      const startCol = mr.getColumn() - range.getColumn() + 1;
      const rowspan = mr.getNumRows();
      const colspan = mr.getNumColumns();
      mergeMap[`${startRow},${startCol}`] = { rowspan, colspan };

      for (let r = startRow; r < startRow + rowspan; r++) {
        for (let c = startCol; c < startCol + colspan; c++) {
          if (!(r === startRow && c === startCol)) {
            mergeMap[`${r},${c}`] = "skip";
          }
        }
      }
    });

    // Build formatted HTML table
    let htmlBody = '<table style="border-collapse: collapse; border: none;">';
    for (let r = 0; r < numRows; r++) {
      htmlBody += '<tr>';
      for (let c = 0; c < numCols; c++) {
        const key = `${r + 1},${c + 1}`;
        if (mergeMap[key] === "skip") continue;

        let cellValue = values[r][c];
        const bgColor = backgrounds[r][c];
        const fontWeight = fontWeights[r][c];
        const fontColor = fontColors[r][c];
        const fontStyle = fontStyles[r][c];

        // Check if cellValue is a Date object and format it
        if (cellValue instanceof Date) {
          cellValue = cellValue.toLocaleDateString('en-US', {
            month: 'long',
            day: 'numeric',
            year: 'numeric'
          });
        }

        let style = `
          padding: 5px;
          border: none;
          background-color: ${bgColor};
          font-weight: ${fontWeight};
          color: ${fontColor};
          font-style: ${fontStyle};
        `;

        let attrs = '';
        if (mergeMap[key]) {
          const { rowspan, colspan } = mergeMap[key];
          if (rowspan > 1) attrs += ` rowspan="${rowspan}"`;
          if (colspan > 1) attrs += ` colspan="${colspan}"`;
        }

        const isHeader = r === 0;
        const tag = isHeader ? 'th' : 'td';
        const finalStyle = isHeader ? style + ' font-weight: bold;' : style;

        htmlBody += `<${tag} style="${finalStyle}"${attrs}>${cellValue}</${tag}>`;
      }
      htmlBody += '</tr>';
    }
    htmlBody += '</table>';

    // Determine recipients
    let recipients = [];
    if (contactSelection === "all") {
      const contactsSheet = ss.getSheetByName("contacts");
      const lastRow = contactsSheet.getLastRow();
      if (lastRow < 2) {
        throw new Error("No contacts found in the contacts sheet.");
      }
      const emailRange = contactsSheet.getRange("P2:P" + lastRow);
      recipients = emailRange.getValues().flat().filter(email => email && email.match(/^[^\s@]+@[^\s@]+\.[^\s@]+$/));
    } else {
      recipients = selectedEmails.filter(email => email && email.match(/^[^\s@]+@[^\s@]+\.[^\s@]+$/));
    }

    if (recipients.length === 0) {
      throw new Error("No valid email addresses selected.");
    }

    // Send emails
    let count = 0;
    recipients.forEach(email => {
      GmailApp.sendEmail(email, subject, '', { htmlBody });
      count++;
    });

    return `Successfully sent ${count} email${count > 1 ? "s" : ""}!`;
  } catch (error) {
    throw new Error(`Error: ${error.message}`);
  }
}

function saveWithTempRename() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = ss.getActiveSheet();
  var originalName = activeSheet.getName(); // Store the original name
  
  // Only rename to "Input" if the sheet is not already named "Input"
  if (originalName !== "Input") {
    activeSheet.setName("Input");
  }
  
  try {
    // Run the save function
    save();
  } catch (e) {
    Logger.log("Error occurred: " + e);
  } finally {
    // Restore the original sheet name only if it was changed
    if (originalName !== "Input") {
      activeSheet.setName(originalName);
    }
  }
}

function saveAndCopyWithTempRename() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = ss.getActiveSheet();
  var originalName = activeSheet.getName(); // Store the original name
  
  // Rename the active sheet to "Input"
  activeSheet.setName("Input");
  
  try {
    // Run the provided functions
    save();
    copyToCodeTotals();
  } catch (e) {
    Logger.log("Error occurred: " + e);
  } finally {
    // Restore the original sheet name
    activeSheet.setName(originalName);
  }
}

function duplicateAndRenameSheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var inputSheet = spreadsheet.getSheetByName('Input') || spreadsheet.getSheetByName('Input Template');
  
  if (!inputSheet) {
    SpreadsheetApp.getUi().alert('No Input or Input Template sheet found!');
    return;
  }
  
  // Prompt user for new sheet name
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Enter new sheet name:', ui.ButtonSet.OK_CANCEL);
  
  // Check if user clicked OK and provided a name
  if (response.getSelectedButton() == ui.Button.OK) {
    var newSheetName = response.getResponseText().trim();
    
    // Validate the new sheet name
    if (newSheetName === '') {
      ui.alert('Please enter a valid sheet name!');
      return;
    }
    
    if (spreadsheet.getSheetByName(newSheetName)) {
      ui.alert('Sheet name already exists! Please choose a different name.');
      return;
    }
    
    // Check if Dashboard sheet exists, create and format if not
    var dashboardSheet = spreadsheet.getSheetByName('Dashboard');
    if (!dashboardSheet) {
      dashboardSheet = spreadsheet.insertSheet('Dashboard');
      
      // Format Dashboard sheet
      // Set title in A1
      dashboardSheet.getRange('A1').setValue('Dashboard to Added Sheets')
        .setFontSize(16)
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
      
      // Merge A1 across multiple columns for better appearance
      dashboardSheet.getRange('A1:C1').merge();
      
      // Set column widths
      dashboardSheet.setColumnWidth(1, 200); // Wider column for sheet names
      dashboardSheet.setColumnWidths(2, dashboardSheet.getMaxColumns() - 1, 100);
      
      // Add header for links in A2
      dashboardSheet.getRange('A2').setValue('Sheet Links')
        .setFontWeight('bold')
        .setBackground('#4A90E2') // Blue background
        .setFontColor('#FFFFFF') // White text
        .setHorizontalAlignment('center');
      
      // Apply alternating row colors starting from A3
      var maxRows = dashboardSheet.getMaxRows();
      if (maxRows > 2) {
        var range = dashboardSheet.getRange('A3:A' + maxRows);
        range.setBackground(null); // Clear existing backgrounds
        for (var i = 3; i <= maxRows; i += 2) {
          dashboardSheet.getRange('A' + i).setBackground('#F5F6F5'); // Light gray
        }
      }
      
      // Add border to header
      dashboardSheet.getRange('A2').setBorder(true, true, true, true, false, false, '#000000', SpreadsheetApp.BorderStyle.SOLID);
      
      // Freeze the first two rows
      dashboardSheet.setFrozenRows(2);
    }
    
    // Duplicate the sheet and rename it
    var newSheet = inputSheet.copyTo(spreadsheet);
    newSheet.setName(newSheetName);
    
    // Add link to Dashboard in cell R1 of the new sheet
    newSheet.getRange('R1').setFormula(`=HYPERLINK("#gid=${dashboardSheet.getSheetId()}","Dashboard")`)
      .setFontSize(12)
      .setFontColor('#4A90E2')
      .setHorizontalAlignment('center');
    
    // Find the next available cell in Dashboard (starting at A3, then A5, A7, etc.)
    var lastRow = dashboardSheet.getLastRow();
    var targetRow = lastRow < 3 ? 3 : lastRow + (lastRow % 2 === 0 ? 1 : 2);
    
    // Add link to the new sheet in the Dashboard
    dashboardSheet.getRange('A' + targetRow)
      .setFormula(`=HYPERLINK("#gid=${newSheet.getSheetId()}","${newSheetName}")`)
      .setFontSize(11)
      .setFontColor('#4A90E2')
      .setHorizontalAlignment('left');
    
    // If original sheet is named 'Input', rename it to 'Input Template'
    if (inputSheet.getName() === 'Input') {
      inputSheet.setName('Input Template');
    }
  }
}

function showUploadDialog() {
  const html = HtmlService.createHtmlOutputFromFile('UploadCSV')
    .setWidth(400)
    .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'Upload CSV File');
}

function processCSV(csvContent) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let importSheet = ss.getSheetByName("Import");

  if (!importSheet) {
    importSheet = ss.insertSheet("Import");
  }

  importSheet.clear();

  const csvData = Utilities.parseCsv(csvContent);
  if (!csvData || csvData.length === 0) {
    throw new Error('The uploaded file is empty or invalid.');
  }

  // Paste the CSV data into the Import sheet
  importSheet.getRange(1, 1, csvData.length, csvData[0].length).setValues(csvData);

  // Define the mapping logic
  const contactsSheet = ss.getSheetByName("contacts");
  if (!contactsSheet) {
    throw new Error("The 'contacts' sheet does not exist.");
  }

  // Copy specific ranges based on the VBA mapping logic
  const mappings = [
    { source: "A1:A2000", target: "B1:B2000" },
    { source: "B1:B2000", target: "C1:C2000" },
    { source: "C1:C2000", target: "D1:D2000" },
    { source: "N1:N2000", target: "T1:T2000" },
    { source: "J1:J2000", target: "P1:P2000" },
    { source: "AM1:AM2000", target: "E1:E2000" },
    { source: "AH1:AH2000", target: "AN1:AN2000" },
    { source: "P1:P2000", target: "V1:V2000" },
    { source: "AJ1:AJ2000", target: "AP1:AP2000" },
    { source: "AL1:AL2000", target: "AR1:AR2000" },
    { source: "AP1:AP2000", target: "AZ1:AZ2000" },
    { source: "AT1:AT2000", target: "BD1:BD2000" },
    { source: "AU1:AU2000", target: "BE1:BE2000" },
    { source: "AV1:AV2000", target: "BF1:BF2000" },
    { source: "T1:T2000", target: "Z1:Z2000" },
    { source: "X1:X2000", target: "AD1:AD2000" },
    { source: "Y1:Y2000", target: "AE1:AE2000" },
    { source: "Z1:Z2000", target: "AF1:AF2000" },
    { source: "BA1:BA2000", target: "BK1:BK2000" },
    { source: "BE1:BE2000", target: "BO1:BO2000" },
    { source: "BF1:BF2000", target: "BP1:BP2000" },
    { source: "BG1:BG2000", target: "BQ1:BQ2000" }
  ];

  // Apply the mappings
  mappings.forEach(({ source, target }) => {
    const sourceRange = importSheet.getRange(source);
    const targetRange = contactsSheet.getRange(target);
    targetRange.setValues(sourceRange.getValues());
  });

  // Copy formulas from column A in Import sheet to A2:A2000 in Contacts sheet
  contactsSheet.getRange('A2:A2000').activate();
  contactsSheet.getRange('A1').copyTo(contactsSheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);


  // Notify user of success
  SpreadsheetApp.getUi().alert("CSV data has been successfully imported and mapped into the 'contacts' sheet.");
}



function edit() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const addressSheet = ss.getSheetByName("Address");

  addressSheet.getRange('F2').activate();
  var lookupValue = addressSheet.getRange('F1').getValue();
  addressSheet.getCurrentCell().setFormula('=VLOOKUP(' + lookupValue + ', contacts!A:CJ, 44, FALSE)');
  addressSheet.getRange('F3').activate();
  var lookupValue = addressSheet.getRange('F1').getValue();
  addressSheet.getCurrentCell().setFormula('=VLOOKUP(' + lookupValue + ', contacts!A:CJ, 52, FALSE)');
  addressSheet.getRange('F4').activate();
  var lookupValue = addressSheet.getRange('F1').getValue();
  addressSheet.getCurrentCell().setFormula('=VLOOKUP(' + lookupValue + ', contacts!A:CJ, 56, FALSE) & ", " & VLOOKUP(' + lookupValue + ', contacts!A:CJ, 57, FALSE) & " " & VLOOKUP(' + lookupValue + ', contacts!A:CJ, 58, FALSE)');

  addressSheet.getRange('F5').activate();
  var lookupValue = addressSheet.getRange('F1').getValue();
  addressSheet.getRange("F5").setFormula('=HYPERLINK(VLOOKUP(' + lookupValue + ', contacts!A:CJ, 16, FALSE))');

  addressSheet.getRange('F6').activate();
  var lookupValue = addressSheet.getRange('F1').getValue();
  addressSheet.getCurrentCell().setFormula('=VLOOKUP(' + lookupValue + ', contacts!A:CJ, 40, FALSE)');
  addressSheet.getRange('F7').activate();
  var lookupValue = addressSheet.getRange('F1').getValue();
  addressSheet.getCurrentCell().setFormula('=VLOOKUP(' + lookupValue + ', contacts!A:CJ, 42, FALSE)');
  addressSheet.getRange('F8').activate();
  var lookupValue = addressSheet.getRange('F1').getValue();
  addressSheet.getCurrentCell().setFormula('=VLOOKUP(' + lookupValue + ', contacts!A:CJ, 26, FALSE)');
  addressSheet.getRange('F9').activate();
  var lookupValue = addressSheet.getRange('F1').getValue();
  addressSheet.getCurrentCell().setFormula('=VLOOKUP(' + lookupValue + ', contacts!A:CJ, 30, FALSE) & ", " & VLOOKUP(' + lookupValue + ', contacts!A:CJ, 31, FALSE) & " " & VLOOKUP(' + lookupValue + ', contacts!A:CJ, 32, FALSE)');
  addressSheet.getRange('F10').activate();
  var lookupValue = addressSheet.getRange('F1').getValue();
  addressSheet.getCurrentCell().setFormula('=VLOOKUP(' + lookupValue + ', contacts!A:CJ, 63, FALSE)');
  addressSheet.getRange('F11').activate();
  var lookupValue = addressSheet.getRange('F1').getValue();
  addressSheet.getCurrentCell().setFormula('=VLOOKUP(' + lookupValue + ', contacts!A:CJ, 67, FALSE) & ", " & VLOOKUP(' + lookupValue + ', contacts!A:CJ, 68, FALSE) & " " & VLOOKUP(' + lookupValue + ', contacts!A:CJ, 69, FALSE)');
  addressSheet.getRange('F12').activate();
  var lookupValue = addressSheet.getRange('F1').getValue();
  addressSheet.getCurrentCell().setFormula('=VLOOKUP(' + lookupValue + ', contacts!A:CJ, 5, FALSE)');
  addressSheet.getRange('F13').activate();
  var lookupValue = addressSheet.getRange('F1').getValue();
  addressSheet.getCurrentCell().setFormula('=VLOOKUP(' + lookupValue + ', contacts!A:CJ, 20, FALSE)');
  addressSheet.getRange('F14').activate();
  var lookupValue = addressSheet.getRange('F1').getValue();
  addressSheet.getCurrentCell().setFormula('=VLOOKUP(' + lookupValue + ', contacts!A:CJ, 22, FALSE)');
}



function copyInput2() {
var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName("Sheet1");
  var targetSheet = ss.getSheetByName("View_Print");
  
  // Define the range to copy
  var copyRange = sourceSheet.getRange("A3:Q48");
  var targetRange = targetSheet.getRange("A3:Q48");
  
  // Copy everything from source to target
  copyRange.copyTo(targetRange); // This will copy values, formats, and formulas
  
  // Copy column widths
  var sourceColWidths = [];
  var lastColumnSource = sourceSheet.getLastColumn();
  var lastColumnTarget = targetSheet.getLastColumn();
  
  // Ensure we only consider columns up to the last column in both sheets
  var columnsToCopy = Math.min(lastColumnSource, lastColumnTarget, 17); // A to Q = 17 columns
  
  for (var i = 1; i <= columnsToCopy; i++) {
    sourceColWidths.push(sourceSheet.getColumnWidth(i));
  }
  
  // Set column widths in target sheet, but only for existing columns
  for (var j = 1; j <= columnsToCopy; j++) {
    targetSheet.setColumnWidth(j, sourceColWidths[j - 1]);
  }
  
  // Select cell C4 in the target sheet
  targetSheet.activate();
  targetSheet.getRange("A1").activate();

 }


function newfile() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheetNames = [
    "Input",
    "View_Print",
    "Log",
    "Update",
    "Data",
  ];

  sheetNames.forEach((sheetName) => {
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      // Insert a new sheet with the specified name
      ss.insertSheet(sheetName);
    }
  });
  const inputSheet = ss.getSheetByName("Input");
  const dataSheet = ss.getSheetByName("Data");
  const viewPrintSheet = ss.getSheetByName("View_Print");
  const updateSheet = ss.getSheetByName("Update");
  const logSheet = ss.getSheetByName("Log");

  inputSheet.getRange("A1:Q1").activate();
  inputSheet.getActiveRangeList().setBackground("#a4c2f4");
  inputSheet.getRange("P2:Q2").activate();
  inputSheet.getActiveRangeList().setBackground("#a4c2f4");
  inputSheet.getRange("A1").activate();
  inputSheet.getCurrentCell().setValue("Log 1");
  inputSheet.getRange("B1").activate();
  inputSheet.getCurrentCell().setValue("Log 2");
  inputSheet.getRange("C1").activate();
  inputSheet.getCurrentCell().setValue("Log 3");
  inputSheet.getRange("D1").activate();
  inputSheet.getCurrentCell().setValue("Log 4");
  inputSheet.getRange("E1").activate();
  inputSheet.getCurrentCell().setValue("Log 5");
  inputSheet.getRange("F1").activate();
  inputSheet.getCurrentCell().setValue("Log 6");
  inputSheet.getRange("G1").activate();
  inputSheet.getCurrentCell().setValue("Log 7");
  inputSheet.getRange("H1").activate();
  inputSheet.getCurrentCell().setValue("Log 8");
  inputSheet.getRange("I1").activate();
  inputSheet.getCurrentCell().setValue("Log 9");
  inputSheet.getRange("J1").activate();
  inputSheet.getCurrentCell().setValue("Log 10");
  inputSheet.getRange("K1").activate();
  inputSheet.getCurrentCell().setValue("Log 11");
  inputSheet.getRange("L1").activate();
  inputSheet.getCurrentCell().setValue("Log 12");
  inputSheet.getRange("M1").activate();
  inputSheet.getCurrentCell().setValue("Update 1");
  inputSheet.getRange("N1").activate();
  inputSheet.getCurrentCell().setValue("Update 2");
  inputSheet.getRange("O1").activate();
  inputSheet.getCurrentCell().setValue("Update 3");
  inputSheet.getRange("P1:Q2").merge();
  inputSheet.getRange("P1").setFormula('=HYPERLINK("https://datamateapp.github.io/help.html", "Help")');
  const cell = inputSheet.getRange("P1");
  cell.setFontWeight("bold");
  cell.setFontSize(16);
  cell.setFontColor("#FF0000");
  cell.setHorizontalAlignment("center");
  cell.setVerticalAlignment("middle");

  inputSheet.getRange("A3:Q48").activate();
  inputSheet.setCurrentCell(inputSheet.getRange("Q48"));
  inputSheet
    .getActiveRangeList()
    .setBorder(false, false, false, false, false, false)
    .setBorder(
      true,
      true,
      true,
      true,
      null,
      null,
      "#000000",
      SpreadsheetApp.BorderStyle.SOLID
    );
  inputSheet.getRange("A1:O2").activate();
  inputSheet.setCurrentCell(inputSheet.getRange("O1"));
  inputSheet
    .getActiveRangeList()
    .setBorder(
      true,
      true,
      true,
      true,
      true,
      true,
      "#000000",
      SpreadsheetApp.BorderStyle.SOLID
    );


  viewPrintSheet.getRange("A1").activate();
  viewPrintSheet.getRange("A1:Q1").activate();
  viewPrintSheet.getActiveRangeList().setBackground("#a4c2f4");
  viewPrintSheet.getRange("A2").activate();
  viewPrintSheet.getActiveRangeList().setBackground("#a4c2f4");
  viewPrintSheet.getRange("P2:Q2").activate();
  viewPrintSheet.getActiveRangeList().setBackground("#a4c2f4");

  viewPrintSheet.getRange("M1").setFormula("=Input!M1");
  viewPrintSheet.getRange("N1").setFormula("=Input!N1");
  viewPrintSheet.getRange("O1").setFormula("=Input!O1");

  viewPrintSheet.getRange('A3:Q48').activate();
viewPrintSheet.setCurrentCell(viewPrintSheet.getRange('Q48'));
viewPrintSheet.getActiveRangeList().setBorder(false, false, false, false, false, false)
  .setBorder(false, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID); // Top border is set to 'false'
viewPrintSheet.setHiddenGridlines(true);

  viewPrintSheet.getRange("B2:L2").activate();
  viewPrintSheet.setCurrentCell(viewPrintSheet.getRange("L2"));
  viewPrintSheet.getActiveRange().mergeAcross();
  viewPrintSheet
    .getRange("B2:L2")
    .setDataValidation(
      SpreadsheetApp.newDataValidation()
        .setAllowInvalid(false)
        .requireValueInRange(viewPrintSheet.getRange("Data!$A:$A"), true)
        .build()
    );
  viewPrintSheet.getRange("B2:L2").activate();
  viewPrintSheet.getActiveRangeList().setBackground("#d9ead3");

  logSheet.getRange("A2").activate();
  logSheet.getCurrentCell().setValue("Anything Log");
  logSheet
    .getActiveRangeList()
    .setFontSize(11)
    .setFontSize(14)
    .setFontWeight("bold");
  logSheet.getRange("A3").activate();
  logSheet.getCurrentCell().setValue("Date");
  logSheet.getRange("B3").activate();
  logSheet.getCurrentCell().setFormula("=TODAY()");
  logSheet.getRange("A9:O10").activate();
  logSheet.getRange("A9:O10").createFilter();

  updateSheet
    .getRangeList([
      "A:A",
      "E:E",
      "F:F",
      "G:G",
      "H:H",
      "I:I",
      "J:J",
      "K:K",
      "L:L",
      "M:M",
      "N:N",
      "O:O",
      "P:P",
      "Q:Q",
    ])
    .activate()
    .setBackground("#f3f3f3");
  updateSheet.getRange("A1:L1").activate();
  updateSheet.setCurrentCell(updateSheet.getRange("L1"));
  updateSheet
    .getActiveRangeList()
    .setBorder(
      null,
      null,
      true,
      null,
      null,
      null,
      "#000000",
      SpreadsheetApp.BorderStyle.SOLID
    );

  updateSheet.getRange("B1").setFormula("=View_Print!M1");
  updateSheet.getRange("C1").setFormula("=View_Print!N1");
  updateSheet.getRange("D1").setFormula("=View_Print!O1");
  updateSheet.getRange("E1").setFormula("=Input!A1");
  updateSheet.getRange("F1").setFormula("=Input!B1");
  updateSheet.getRange("G1").setFormula("=Input!C1");
  updateSheet.getRange("H1").setFormula("=Input!D1");
  updateSheet.getRange("I1").setFormula("=Input!E1");
  updateSheet.getRange("J1").setFormula("=Input!F1");
  updateSheet.getRange("K1").setFormula("=Input!G1");
  updateSheet.getRange("L1").setFormula("=Input!H1");
  updateSheet.getRange("M1").setFormula("=Input!I1");
  updateSheet.getRange("N1").setFormula("=Input!J1");
  updateSheet.getRange("O1").setFormula("=Input!K1");
  updateSheet.getRange("P1").setFormula("=Input!L1");

  const lastColumn = dataSheet.getMaxColumns();
  const columnsToAdd = 800; // Number of columns from the last column to AFC

  dataSheet.insertColumnsAfter(lastColumn, columnsToAdd);
  const lastRow = dataSheet.getMaxRows();
  const rowsToAdd = 10000;

  dataSheet.insertRowsAfter(lastRow, rowsToAdd);
  dataSheet.getRange("S1").activate();
  dataSheet.getRange("S1").setFormula("=Input!A1");
  dataSheet.getRange("T1").activate();
  dataSheet.getRange("T1").setFormula("=Input!B1");
  dataSheet.getRange("U1").activate();
  dataSheet.getRange("U1").setFormula("=Input!C1");
  dataSheet.getRange("V1").activate();
  dataSheet.getRange("V1").setFormula("=Input!D1");
  dataSheet.getRange("W1").activate();
  dataSheet.getRange("W1").setFormula("=Input!E1");
  dataSheet.getRange("X1").activate();
  dataSheet.getRange("X1").setFormula("=Input!F1");
  dataSheet.getRange("Y1").activate();
  dataSheet.getRange("Y1").setFormula("=Input!G1");
  dataSheet.getRange("Z1").activate();
  dataSheet.getRange("Z1").setFormula("=Input!H1");
  dataSheet.getRange("AA1").activate();
  dataSheet.getRange("AA1").setFormula("=Input!I1");
  dataSheet.getRange("AB1").activate();
  dataSheet.getRange("AB1").setFormula("=Input!J1");
  dataSheet.getRange("AC1").activate();
  dataSheet.getRange("AC1").setFormula("=Input!K1");
  dataSheet.getRange("AD1").activate();
  dataSheet.getRange("AD1").setFormula("=Input!L1");
  dataSheet.getRange("AE1").activate();
  dataSheet.getRange("AE1").setFormula("=Input!M1");
  dataSheet.getRange("AF1").activate();
  dataSheet.getRange("AF1").setFormula("=Input!N1");
  dataSheet.getRange("AG1").activate();
  dataSheet.getRange("AG1").setFormula("=Input!O1");

  copyInput1();
  view();
  sendDataMateEmail();


  inputSheet.getRange("W1").activate();
}

function sendDataMateEmail() {
  MailApp.sendEmail({
    to: "projectprodigyapp@gmail.com",
    subject: "DataMate Dataset Created",
    body: "A custom Dataset has been added to a spreadsheet."
  });
}
  



function view() {
  
  copyInput2(); 
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const wsViewPrint = ss.getSheetByName("View_Print");
  const wsUpdate = ss.getSheetByName("Update");
  const wsData = ss.getSheetByName("Data");

  // Update specific formulas in View_Print
  wsViewPrint.getRange("A1").setFormula("=View_Print!B2");
  wsViewPrint.getRange("M2").setFormula("=VLOOKUP(A1,Update!$A$1:$P$10000,2,FALSE)");
  wsViewPrint.getRange("N2").setFormula("=VLOOKUP(A1,Update!$A$1:$P$10000,3,FALSE)");
  wsViewPrint.getRange("O2").setFormula("=VLOOKUP(A1,Update!$A$1:$P$10000,4,FALSE)");

  // Add dynamic VLOOKUP formulas for other cells
  for (let i = 3; i <= 48; i++) {
    for (let j = 1; j <= 17; j++) {
      const formula = `=VLOOKUP(A1,Data!$A$1:$DZU$10000,${801 + (i - 48) * 17 + (j - 1)},FALSE)`;
      wsViewPrint.getRange(i, j).setFormula(formula);
    }
  }

  Logger.log("View_Print refreshed successfully!");
}

function contacts() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetNames = ["contacts", "Address", "NewContact"];
    const sheets = ss.getSheets();

    // Add missing sheets
    sheetNames.forEach(name => {
      if (!sheets.some(sheet => sheet.getName() === name)) {
        ss.insertSheet(name);
      }
    });

    const contactsSheet = ss.getSheetByName("contacts");
    const addressSheet = ss.getSheetByName("Address");
    const newContactSheet = ss.getSheetByName("NewContact");

    if (!contactsSheet || !addressSheet || !newContactSheet) {
      throw new Error("One or more required sheets are missing.");
    }

    // Setup Contacts sheet
    contactsSheet.activate();
    const lastColumn = contactsSheet.getMaxColumns();
    const columnsToAdd = 80; // Number of columns from the last column to AFC
    contactsSheet.insertColumnsAfter(lastColumn, columnsToAdd);

    const lastRow = contactsSheet.getMaxRows();
    const rowsToAdd = 2000;
    contactsSheet.insertRowsAfter(lastRow, rowsToAdd);

    // Add headers and formatting
    contactsSheet.getRange("A1").setFormula('=B1 & " " & C1 & " " & D1');
    contactsSheet.getRange("B1:E1").setValues([["First Name", "Middle Name", "Last Name", "Title"]]);

    contactsSheet.getRange("P1").setValue("E-mail Address");
    contactsSheet.getRange("T1").setValue("Home Phone");
    contactsSheet.getRange("V1").setValue("Mobile Phone");
    contactsSheet.getRange("Z1").setValue("Home Street");
    contactsSheet.getRange("AD1").setValue("Home City");
    contactsSheet.getRange("AE1").setValue("Home State");
    contactsSheet.getRange("AF1").setValue("Home Postal Code");
    contactsSheet.getRange("AN1").setValue("Business Phone");
    contactsSheet.getRange("AP1").setValue("Business Fax");
    contactsSheet.getRange("AR1").setValue("Company");
    contactsSheet.getRange("AZ1").setValue("Business Street");
    contactsSheet.getRange("BD1").setValue("Business City");
    contactsSheet.getRange("BE1").setValue("Business State");
    contactsSheet.getRange("BF1").setValue("Business Postal Code");
    contactsSheet.getRange("BK1").setValue("Other Street");
    contactsSheet.getRange("BO1").setValue("Other City");
    contactsSheet.getRange("BP1").setValue("Other State");
    contactsSheet.getRange("BQ1").setValue("Other Postal Code");
    contactsSheet.getRange("A1:CL2").createFilter();


    // Setup Address sheet
    addressSheet.activate();
    addressSheet.getRange("B1:D1").merge();
    addressSheet.getRange("B1:D1").setBackground("#D9EAD3");
   // Define the validation range
  const validationRange = addressSheet.getRange("B1:D1");
  const sourceRange = contactsSheet.getRange("A:A"); // Source data for validation

  // Build and apply the data validation rule
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(sourceRange, true) // True for strict validation
    .setAllowInvalid(false) // Prevent invalid values
    .build();
  validationRange.setDataValidation(rule); // Apply the rule

  addressSheet.getRange("B2").setFormula("=VLOOKUP(B1, contacts!A:CJ, 44, FALSE)");
  addressSheet.getRange("B3").setFormula("=VLOOKUP(B1, contacts!A:CJ, 52, FALSE)");
  addressSheet.getRange("B4").setFormula('=VLOOKUP(B1, contacts!A:CJ, 56, FALSE) & ", " & VLOOKUP(B1, contacts!A:CJ, 57, FALSE) & "   " & VLOOKUP(B1, contacts!A:CJ, 58, FALSE)');
  addressSheet.getRange("B5").setFormula('=HYPERLINK(VLOOKUP(B1, contacts!A:CJ, 16, FALSE))');
  addressSheet.getRange("B6").setFormula('=VLOOKUP(B1, contacts!A:CJ, 40, FALSE)');
  addressSheet.getRange("B7").setFormula('=VLOOKUP(B1, contacts!A:CJ, 42, FALSE)');
  addressSheet.getRange("B8").setFormula("=VLOOKUP(B1, contacts!A:CJ, 26, FALSE)");
  addressSheet.getRange("B9").setFormula('=VLOOKUP(B1, contacts!A:CJ, 30, FALSE) & ", " & VLOOKUP(B1, contacts!A:CJ, 31, FALSE) & "   " & VLOOKUP(B1, contacts!A:CJ, 32, FALSE)');
  addressSheet.getRange("B10").setFormula("=VLOOKUP(B1, contacts!A:CJ, 63, FALSE)");
  addressSheet.getRange("B11").setFormula('=VLOOKUP(B1, contacts!A:CJ, 67, FALSE) & ", " & VLOOKUP(B1, contacts!A:CJ, 68, FALSE) & "   " & VLOOKUP(B1, contacts!A:CJ, 69, FALSE)');
  addressSheet.getRange("B12").setFormula("=VLOOKUP(B1, contacts!A:CJ, 5, FALSE)");
  addressSheet.getRange("B13").setFormula("=VLOOKUP(B1, contacts!A:CJ, 20, FALSE)");
  addressSheet.getRange("B14").setFormula("=VLOOKUP(B1, contacts!A:CJ, 22, FALSE)");


    addressSheet.getRange("A1:A14").setFontWeight("bold");
    addressSheet.getRange("E1").setValue("Target Cell on Sheet1").setFontColor("red");
    addressSheet.getRange("F1").setBackground("#D9EAD3");
    addressSheet.getRange("E1:E14").setFontWeight("bold");

    // Add formulas to Address sheet
    const formulasA = [
      "=contacts!A1", "=contacts!AR1", "=contacts!AZ1", "=contacts!BD1",
      "=contacts!P1", "=contacts!AN1", "=contacts!AP1", "=contacts!Z1",
      "=contacts!AD1", "=contacts!BK1", "=contacts!BO1", "=contacts!E1",
      "=contacts!T1", "=contacts!V1"
    ];
    formulasA.forEach((formula, index) => {
      addressSheet.getRange(`A${index + 1}`).setFormula(formula);
    });

    const formulasE = formulasA.slice(1);
    formulasE.forEach((formula, index) => {
      addressSheet.getRange(`E${index + 2}`).setFormula(formula);
    });

    addressSheet.getRange("F15").setValue("Vlookup by Name");
    addressSheet.getRange("G15").setValue("Xlookup by Company");
    addressSheet.setColumnWidth(1, 200);
    addressSheet.setColumnWidth(5, 200);
    addressSheet.setColumnWidth(6, 200);
    addressSheet.setColumnWidth(7, 200);

    // Setup NewContact sheet
    const formulasNewContact = [
      "=contacts!B1", "=contacts!C1", "=contacts!D1", "=contacts!AR1",
      "=contacts!AZ1", "=contacts!BD1", "=contacts!BE1", "=contacts!BF1",
      "=contacts!P1", "=contacts!AN1", "=contacts!AP1", "=contacts!Z1",
      "=contacts!AD1", "=contacts!AE1", "=contacts!AF1", "=contacts!BK1",
      "=contacts!BO1", "=contacts!BP1", "=contacts!BQ1", "=contacts!E1",
      "=contacts!T1", "=contacts!V1"
    ];
    formulasNewContact.forEach((formula, index) => {
      newContactSheet.getRange(`A${index + 1}`).setFormula(formula);
    });
    newContactSheet.getRange("B1:B22").setBackground("#D9EAD3");
    newContactSheet.getRange("A:A").setFontWeight("bold");
    newContactSheet.getRange("B23").setValue("Enter information and select New Contact.");
    
  newContactSheet
    .getActiveRangeList()
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP)
    .setVerticalAlignment("top")
    .setFontSize(14)
    .setFontWeight("bold");

    newContactSheet.setColumnWidth(1, 200);
    newContactSheet.setColumnWidth(2, 200);


    addressSheet.getRange("B1").activate()

    SpreadsheetApp.getUi().alert("Sheets created/updated successfully!");
  } catch (e) {
    SpreadsheetApp.getUi().alert(`Error: ${e.message}`);
  }
}

  
function EditAddressSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const wsAddress = ss.getActiveSheet();
    const wsContacts = ss.getSheetByName("contacts");
    const wsSheet1 = ss.getSheetByName("Sheet1");

    if (!wsAddress) {
      SpreadsheetApp.getUi().alert("No active sheet found.");
      return;
    }

    if (!wsContacts) {
      SpreadsheetApp.getUi().alert("'contacts' sheet not found in the active spreadsheet.");
      return;
    }

    if (!wsSheet1) {
      SpreadsheetApp.getUi().alert("'Sheet1' not found in the active spreadsheet.");
      return;
    }

    // Get the lookup value from F1
    const lookupValue = wsAddress.getRange("F1").getValue();

    if (!lookupValue) {
      SpreadsheetApp.getUi().alert("Cell F1 on the Address sheet is empty. Please enter a valid cell address (e.g., 'B2').");
      return;
    }

    // Set formulas in column F
    const formulas = [
     `=VLOOKUP(${lookupValue}, contacts!A:CJ, 44, FALSE)`,
  `=VLOOKUP(${lookupValue}, contacts!A:CJ, 52, FALSE)`,
  `=VLOOKUP(${lookupValue}, contacts!A:CJ, 56, FALSE) & ", " & VLOOKUP(${lookupValue}, contacts!A:CJ, 57, FALSE) & " " & VLOOKUP(${lookupValue}, contacts!A:CJ, 58, FALSE)`,
  `=HYPERLINK(VLOOKUP(${lookupValue}, contacts!A:CJ, 16, FALSE))`,
  `=VLOOKUP(${lookupValue}, contacts!A:CJ, 40, FALSE)`,
  `=VLOOKUP(${lookupValue}, contacts!A:CJ, 42, FALSE)`, // Fixed this line
  `=VLOOKUP(${lookupValue}, contacts!A:CJ, 26, FALSE)`,
  `=VLOOKUP(${lookupValue}, contacts!A:CJ, 30, FALSE) & ", " & VLOOKUP(${lookupValue}, contacts!A:CJ, 31, FALSE) & " " & VLOOKUP(${lookupValue}, contacts!A:CJ, 32, FALSE)`,
  `=VLOOKUP(${lookupValue}, contacts!A:CJ, 63, FALSE)`,
  `=VLOOKUP(${lookupValue}, contacts!A:CJ, 67, FALSE) & ", " & VLOOKUP(${lookupValue}, contacts!A:CJ, 68, FALSE) & " " & VLOOKUP(${lookupValue}, contacts!A:CJ, 69, FALSE)`,
  `=VLOOKUP(${lookupValue}, contacts!A:CJ, 5, FALSE)`,
  `=VLOOKUP(${lookupValue}, contacts!A:CJ, 20, FALSE)`,
  `=VLOOKUP(${lookupValue}, contacts!A:CJ, 22, FALSE)`
    ];

    wsAddress.getRange("F2:F14").setFormulas(formulas.map(f => [f]));

    // Get the cell address from F1
    const cellAddress = lookupValue;

    if (!cellAddress.match(/^[A-Z]+\d+$/)) {
      SpreadsheetApp.getUi().alert("Invalid cell address in F1. Please enter a valid address (e.g., 'B2').");
      return;
    }

    // Validate the existence of the target cell in Sheet1
    let validationCell;
    try {
      validationCell = wsSheet1.getRange(cellAddress);
    } catch (e) {
      SpreadsheetApp.getUi().alert("Invalid cell address in F1. Please enter a valid address (e.g., 'B2').");
      return;
    }

    // Add data validation to the target cell
    const range = wsContacts.getRange("A:A");
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(range, true)
      .setAllowInvalid(false)
      .build();
    validationCell.setDataValidation(rule);

    // Copy formulas to Sheet1 starting under the validated cell
    const formulasToCopy = wsAddress.getRange("F2:F14").getFormulas();
    const targetRange = wsSheet1.getRange(validationCell.getRow() + 1, validationCell.getColumn(), formulasToCopy.length, 1);
    targetRange.setFormulas(formulasToCopy);

    wsSheet1.getRange("A1").activate()

    SpreadsheetApp.getUi().alert("Validation applied and formulas pasted successfully.");
  } catch (error) {
    SpreadsheetApp.getUi().alert(`Error: ${error.message}`);
  }
}

function EditAddressSheet1() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const wsAddress = ss.getActiveSheet();
    const wsContacts = ss.getSheetByName("contacts");
    const wsSheet1 = ss.getSheetByName("Sheet1");

    if (!wsAddress) {
      SpreadsheetApp.getUi().alert("No active sheet found.");
      return;
    }

    if (!wsContacts) {
      SpreadsheetApp.getUi().alert("'contacts' sheet not found in the active spreadsheet.");
      return;
    }

    if (!wsSheet1) {
      SpreadsheetApp.getUi().alert("'Sheet1' not found in the active spreadsheet.");
      return;
    }

    // Get the lookup value from F1
    const lookupValue = wsAddress.getRange("F1").getValue();

    if (!lookupValue) {
      SpreadsheetApp.getUi().alert("Cell F1 on the Address sheet is empty. Please enter a valid cell address (e.g., 'B2').");
      return;
    }

    // Set XLOOKUP-like formulas in column G
    const formulas = [
      `=IFERROR(INDEX(contacts!A:A, MATCH(${lookupValue}, contacts!AR:AR, 0)), "Not Found")`,
      `=IFERROR(INDEX(contacts!AZ:AZ, MATCH(${lookupValue}, contacts!AR:AR, 0)), "Not Found")`,
      `=IFERROR(INDEX(contacts!BD:BD, MATCH(${lookupValue}, contacts!AR:AR, 0)), "") & ", " & IFERROR(INDEX(contacts!BE:BE, MATCH(${lookupValue}, contacts!AR:AR, 0)), "") & " " & IFERROR(INDEX(contacts!BF:BF, MATCH(${lookupValue}, contacts!AR:AR, 0)), "Not Found")`,
      `=IFERROR(INDEX(contacts!P:P, MATCH(${lookupValue}, contacts!AR:AR, 0)), "Not Found")`,
      `=IFERROR(INDEX(contacts!AN:AN, MATCH(${lookupValue}, contacts!AR:AR, 0)), "Not Found")`,
      `=IFERROR(INDEX(contacts!AP:AP, MATCH(${lookupValue}, contacts!AR:AR, 0)), "Not Found")`,
      `=IFERROR(INDEX(contacts!Z:Z, MATCH(${lookupValue}, contacts!AR:AR, 0)), "Not Found")`,
      `=IFERROR(INDEX(contacts!AD:AD, MATCH(${lookupValue}, contacts!AR:AR, 0)), "") & ", " & IFERROR(INDEX(contacts!AE:AE, MATCH(${lookupValue}, contacts!AR:AR, 0)), "") & " " & IFERROR(INDEX(contacts!AF:AF, MATCH(${lookupValue}, contacts!AR:AR, 0)), "Not Found")`,
      `=IFERROR(INDEX(contacts!BK:BK, MATCH(${lookupValue}, contacts!AR:AR, 0)), "Not Found")`,
      `=IFERROR(INDEX(contacts!BO:BO, MATCH(${lookupValue}, contacts!AR:AR, 0)), "") & ", " & IFERROR(INDEX(contacts!BP:BP, MATCH(${lookupValue}, contacts!AR:AR, 0)), "") & " " & IFERROR(INDEX(contacts!BQ:BQ, MATCH(${lookupValue}, contacts!AR:AR, 0)), "Not Found")`,
      `=IFERROR(INDEX(contacts!E:E, MATCH(${lookupValue}, contacts!AR:AR, 0)), "Not Found")`,
      `=IFERROR(INDEX(contacts!T:T, MATCH(${lookupValue}, contacts!AR:AR, 0)), "Not Found")`,
      `=IFERROR(INDEX(contacts!V:V, MATCH(${lookupValue}, contacts!AR:AR, 0)), "Not Found")`
    ];

    wsAddress.getRange("G2:G14").setFormulas(formulas.map(f => [f]));

    // Get the cell address from F1
    const cellAddress = lookupValue;

    if (!cellAddress.match(/^[A-Z]+\d+$/)) {
      SpreadsheetApp.getUi().alert("Invalid cell address in F1. Please enter a valid address (e.g., 'B2').");
      return;
    }

    // Validate the existence of the target cell in Sheet1
    let validationCell;
    try {
      validationCell = wsSheet1.getRange(cellAddress);
    } catch (e) {
      SpreadsheetApp.getUi().alert("Invalid cell address in F1. Please enter a valid address (e.g., 'B2').");
      return;
    }

    // Add data validation to the target cell
    const range = wsContacts.getRange("AR:AR");
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(range, true)
      .setAllowInvalid(false)
      .build();
    validationCell.setDataValidation(rule);

    // Copy formulas to Sheet1 starting under the validated cell
    const formulasToCopy = wsAddress.getRange("G2:G14").getFormulas();
    const targetRange = wsSheet1.getRange(validationCell.getRow() + 1, validationCell.getColumn(), formulasToCopy.length, 1);
    targetRange.setFormulas(formulasToCopy);

    wsSheet1.getRange("A1").activate()

    SpreadsheetApp.getUi().alert("Validation applied and formulas pasted successfully.");
  } catch (error) {
    SpreadsheetApp.getUi().alert(`Error: ${error.message}`);
  }
}





function setupTS() {
  createTimeSheet();
  newfile();
  cleanupTS();
  
}

function cleanupTS() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName("Input");
  const logSheet = ss.getSheetByName("Log");

  const mappings = [
    ["A1", "=J5"], ["A2", "=J6"], ["B1", "=B3"], ["B2", "=B4"],
    ["C1", "=A5"], ["C2", "=A6"], ["D1", "=A41"], ["D2", "=P43"],
    ["E1", "=P41"], ["E2", "=P42"], ["F1", "=Q41"], ["F2", "=Q42"],
    ["G1", "=A45"], ["G2", "=B45"], ["H1", "=E45"], ["H2", "=F45"],
    ["I1", "=I45"],["I2", "=J45"], ["J1", "=M45"], ["J2", "=O45"],
    ["K1", "Log 11"], ["L1", "Log 12"],
    ["M1", "Update 1"], ["N1", "Update 2"], ["O1", "Update 3"]
  ];

  mappings.forEach(([cell, value]) => {
    inputSheet.getRange(cell).setValue(value);
  });

  logSheet.getRange("A2").setValue("Time/Cost Log");
  logSheet.getRange("C7").setValue("TOTALS");
  logSheet.getRange("D7").setFormula("=SUM(D9:D)");
  logSheet.getRange("E7").setFormula("=SUM(E9:E)");
  logSheet.getRange("F7").setFormula("=SUM(F9:F)");
  logSheet.getRange("G7").setFormula("=SUM(G9:G)");
  logSheet.getRange("H7").setFormula("=SUM(H9:H)");
  logSheet.getRange("I7").setFormula("=SUM(I9:I)");
  logSheet.getRange("J7").setFormula("=SUM(J9:J)");

  // Rename sheet
  inputSheet.setName("Input Template");
}


function copyToCodeTotals() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var inputSheet = ss.getSheetByName("Input");
  var totalsSheet = ss.getSheetByName("Code Totals");

  if (!inputSheet || !totalsSheet) {
    Logger.log("One or both sheets are missing.");
    return;
  }

  // Get data from the Input sheet
  var costCodes = inputSheet.getRange("A9:A40").getValues(); // Cost codes (double rows)
  var otHours = inputSheet.getRange("P9:P40").getValues();   // OT hours
  var dtHours = inputSheet.getRange("Q9:Q40").getValues();   // DT hours
  var regHoursP = inputSheet.getRange("P10:P40").getValues(); // Regular hours from P
  var regHoursQ = inputSheet.getRange("Q10:Q40").getValues(); // Regular hours from Q

  // Get existing data from Code Totals
  var totalsData = totalsSheet.getRange("A2:D" + totalsSheet.getLastRow()).getValues();
  var totalsMap = {}; // Store existing codes and their row index

  // Map existing cost codes to their row index
  totalsData.forEach((row, index) => {
    if (row[0]) totalsMap[row[0]] = index + 2; // Row index (considering header row)
  });

  for (var i = 0; i < costCodes.length; i += 2) { // Process in pairs (double rows)
    var code = costCodes[i][0]; // Cost code in A9, A11, etc.
    if (!code) continue; // Skip empty rows

    var ot = otHours[i][0] || 0;
    var dt = dtHours[i][0] || 0;
    var reg = (regHoursP[i][0] || 0) + (regHoursQ[i][0] || 0); // Sum regular hours correctly

    if (code in totalsMap) {
      // Get the existing row index
      var rowIndex = totalsMap[code];

      // Fetch current values from "Code Totals" before updating
      var currentReg = totalsSheet.getRange(rowIndex, 2).getValue() || 0;
      var currentOT = totalsSheet.getRange(rowIndex, 3).getValue() || 0;
      var currentDT = totalsSheet.getRange(rowIndex, 4).getValue() || 0;

      // Update the values by adding to the existing ones
      totalsSheet.getRange(rowIndex, 2).setValue(currentReg + reg); // Regular Hours
      totalsSheet.getRange(rowIndex, 3).setValue(currentOT + ot);   // OT Hours
      totalsSheet.getRange(rowIndex, 4).setValue(currentDT + dt);   // DT Hours
    } else {
      // Append new row if cost code doesn't exist
      var nextRow = totalsSheet.getLastRow() + 1;
      totalsSheet.getRange(nextRow, 1, 1, 4).setValues([[code, reg, ot, dt]]);
    }
  }
}



function createTimeSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let listsSheet = ss.getSheetByName("Lists");
  
  // Check if Lists sheet exists, create it if it doesn't
  if (!listsSheet) {
    // Create a new Lists sheet
    listsSheet = ss.insertSheet("Lists");

    // Define headers
    const headers = ["Name", "Emp. No", "Rate", "", "Crafts", "Cost Codes"];
    listsSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // Define sample data
    const data = ["Moe", 3000, "$43.68", "", "Carpenter", "13100   Project Superintendent"];
    listsSheet.getRange(2, 1, 1, data.length).setValues([data]);

    listsSheet.getRange("H1")
      .setFormula('=HYPERLINK("https://datamateapp.github.io/About%20Timesheet.html", "About Timesheet")');

    // Auto-size columns for better readability
    listsSheet.autoResizeColumns(1, headers.length);
  }

  // Check if Code Totals sheet exists, create it if it doesn't
  let codeSheet = ss.getSheetByName("Code Totals");
  if (!codeSheet) {
    // Create a new Code Totals sheet
    codeSheet = ss.insertSheet("Code Totals");

    // Define headers
    const headers = ["Cost Code", "Regular Hours", "OT Hours", "DT Hours"];
    codeSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // Auto-size columns for better readability
    codeSheet.autoResizeColumns(1, headers.length);
  }



  
  const existingSheet = ss.getSheetByName("Sheet1");
  if (existingSheet) ss.deleteSheet(existingSheet);

  const sheet = ss.insertSheet("Sheet1");
  sheet.getRange("A3:Q48").clear();

  // Set column widths
  const columnWidths = [300, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30];
  columnWidths.forEach((width, i) => sheet.setColumnWidth(i + 1, width));

  // Merge and set headers
  sheet.getRange("A3").setFontWeight("bold").setValue("EMPLOYEE NO.");
  sheet.getRange("A4").setHorizontalAlignment("center");



  sheet.getRange("B3:G3").merge().setFontWeight("bold").setValue("EMPLOYEE NAME");
  sheet.getRange("B4:G4").merge().setBorder(true, true, true, true, true, true);
  sheet.getRange("B6:G6").merge().setBorder(true, true, true, true, true, true);
  sheet.getRange("H3:I3").merge().setFontWeight("bold").setValue("Note:");
  sheet.getRange("J3:Q3").merge().setBorder(true, true, true, true, true, true);
  sheet.getRange("J4:M4").merge().setFontWeight("bold").setValue("PREPAID CHECK #");
  sheet.getRange("O4:Q4").merge().setBorder(true, true, true, true, true, true);
  sheet.getRange("J5:M5").merge().setFontWeight("bold").setValue("BEGINNING DATE");
  sheet.getRange("J6:M6").merge().setValue("=TODAY()").setBorder(true, true, true, true, true, true);

  sheet.getRange("A5").setFontWeight("bold").setValue("RATE");
  sheet.getRange("A6").setHorizontalAlignment("center");
  sheet.getRange("B6:G6").merge();
  sheet.getRange("B5:G5").merge().setFontWeight("bold").setValue("CRAFT");
  sheet.getRange("A7").setValue("COST CODE");

  // Generate dates and days
  for (let i = 0; i < 7; i++) {
    let col = 2 + (i * 2); // Start from column B (2) and increment by 2
    let dateCell = sheet.getRange(7, col);
    dateCell.setFormula(`=IF(J6="", "", J6+${i})`);
    dateCell.offset(1, 0).setFormula(`=TEXT(${dateCell.getA1Notation()}, "ddd")`);
  }

  sheet.getRange("P7").setFontWeight("bold").setValue("TOTAL");

  // Merge formatting for table header
  const mergeRanges = [
    "A7:A8", "B7:C7", "D7:E7", "F7:G7", "H7:I7", "J7:K7", "L7:M7", "N7:O7",
    "B8:C8", "D8:E8", "F8:G8", "H8:I8", "J8:K8", "L8:M8", "N8:O8", "P7:Q8"
  ];
  mergeRanges.forEach(range => sheet.getRange(range).merge());

  // Set formulas for totals
  sheet.getRange("P9").setFormula('=SUM(B9+D9+F9+H9+J9+L9+N9)');
  sheet.getRange("Q9").setFormula('=SUM(C9+E9+G9+I9+K9+M9+O9)');
  sheet.getRange("P10:Q10").setFormula('=SUM(B10+D10+F10+H10+J10+L10+N10)');


// Get last row with data in column F
let lastRow = listsSheet.getLastRow();
let costCodeRange = listsSheet.getRange("F1:F2000" + lastRow);

// Create data validation rule (Dropdown from range)
const costCodeValidation = SpreadsheetApp.newDataValidation()
  .requireValueInRange(costCodeRange, true) // `true` ensures it's a dynamic range
  .setAllowInvalid(false)
  .build();

// Apply to target range
sheet.getRange("A9:A10").merge().setDataValidation(costCodeValidation);



  const cellMerges = ["B10:C10", "D10:E10", "F10:G10", "H10:I10", "J10:K10", "L10:M10", "N10:O10", "P10:Q10"];
  cellMerges.forEach(range => sheet.getRange(range).merge());

  // Copy and paste rows 9 & 10 to target row pairs
  const targetRows = [11, 13, 15, 17, 19, 21, 23, 25, 27, 29, 31, 33, 35, 37, 39];

  targetRows.forEach(row => {
    let sourceRange = sheet.getRange("A9:Q10");
    let destinationRange = sheet.getRange(row, 1, 2, 17); // Two-row range at target row
    sourceRange.copyTo(destinationRange);
  });

  // Apply borders to table
  sheet.getRange("A7:Q41").setBorder(true, true, true, true, true, true);

  sheet.getRange("A41:A43").merge().setValue("TOTAL HOURS").setHorizontalAlignment("center").setFontWeight("bold").setBorder(true, true, true, true, true, true);
  sheet.getRange("P41").setFontWeight("bold").setValue("OT")
  sheet.getRange("Q41").setFontWeight("bold").setValue("DT")
  sheet.getRange("B42:C43").merge().setBorder(true, true, true, true, true, true);
  sheet.getRange("D42:E43").merge().setBorder(true, true, true, true, true, true);
  sheet.getRange("F42:G43").merge().setBorder(true, true, true, true, true, true);
  sheet.getRange("H42:I43").merge().setBorder(true, true, true, true, true, true);
  sheet.getRange("J42:K43").merge().setBorder(true, true, true, true, true, true);
  sheet.getRange("L42:M43").merge().setBorder(true, true, true, true, true, true);
  sheet.getRange("N42:O43").merge().setBorder(true, true, true, true, true, true);
  sheet.getRange("P43:Q43").merge().setBorder(true, true, true, true, true, true);
  sheet.getRange("P42").setBorder(true, true, true, true, true, true).setFormula('=SUM(P9+P11+P13+P15+P17+P19+P21+P23+P25+P27+P29+P31+P33+P35+P37+P39)');
  sheet.getRange("Q42").setBorder(true, true, true, true, true, true).setFormula('=SUM(Q9+Q11+Q13+Q15+Q17+Q19+Q21+Q23+Q25+Q27+Q29+Q31+Q33+Q35+Q37+Q39)');
  sheet.getRange("P43:Q43").setFormula('=SUM(P10+P12+P14+P16+P18+P20+P22+P24+P26+P28+P30+P32+P34+P36+P38+P40)');
  
    // Get last row with data in column A
let nameRange = listsSheet.getRange("A1:A2000" + lastRow); // Correctly references the range

// Create data validation rule (Dropdown from range)
const nameValidation = SpreadsheetApp.newDataValidation()
  .requireValueInRange(nameRange, true) // `true` ensures it's a dynamic range
  .setAllowInvalid(false)
  .build();

// Apply validation to B4:G4
sheet.getRange("B4:G4").setDataValidation(nameValidation);



   
  // Get last row with data in column E
let craftRange = listsSheet.getRange("E1:E2000" + lastRow); // Correctly references the range

// Create data validation rule (Dropdown from range)
const craftValidation = SpreadsheetApp.newDataValidation()
  .requireValueInRange(craftRange, true) // `true` ensures it's a dynamic range
  .setAllowInvalid(false)
  .build();

// Apply validation to B6:G6
sheet.getRange("B6:G6").setDataValidation(craftValidation);


  sheet.getRange("A4").setFormula('=VLOOKUP(B4, Lists!A:B, 2, FALSE)');
  sheet.getRange("A6").setFormula('=VLOOKUP(B4, Lists!A:C, 3, FALSE)');

  
  sheet.getRange("A45").setHorizontalAlignment("right").setFontWeight("bold").setValue("Regular")
  sheet.getRange("B45:D45").merge().setHorizontalAlignment("center").setFormula('=SUM(P43*A6)')
  sheet.getRange("E45").setFontWeight("bold").setValue("OT")
  sheet.getRange("F45:H45").merge().setHorizontalAlignment("center").setFormula('=SUM(P42*A6*1.5)')
  sheet.getRange("I45").setFontWeight("bold").setValue("DT")
  sheet.getRange("J45:L45").merge().setHorizontalAlignment("center").setFormula('=SUM(Q42*A6*2)')
  sheet.getRange("M45:N45").merge().setFontWeight("bold").setValue("GROSS")
  sheet.getRange("O45:Q45").merge().setHorizontalAlignment("center").setFormula('=SUM(B45+F45+J45)')


}


function setupPO() {
  contacts();
  createPurchaseOrderTemplate();
  newfile();
  cleanupPO();
  
}

function cleanupPO() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName("Input");
  const logSheet = ss.getSheetByName("Log");

  const mappings = [
    ["A1", "=A6"], ["A2", "=C6"], ["B1", "=F6"], ["B2", "=H6"],
    ["C1", "=A8"], ["C2", "=B8"], ["D1", "=A20"], ["D2", "=B20"],
    ["E1", "=F47"], ["E2", "=G47"], ["F1", "=H37"], ["F2", "=J37"],
    ["G1", "=A23"], ["G2", "=B23"], ["H1", "=F39"], ["H2", "=G39"],
    ["I1", "Log 9"], ["J1", "Log 10"], ["K1", "Log 11"], ["L1", "Log 12"],
    ["M1", "=A44"], ["N1", "=A47"], ["O1", "Update 3"]
  ];

  mappings.forEach(([cell, value]) => {
    inputSheet.getRange(cell).setValue(value);
  });

  logSheet.getRange("A2").setValue("Purchase Order Log");
}


function createPurchaseOrderTemplate() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Sheet1") || ss.insertSheet("Sheet1");
    sheet.clear(); // Clear previous content

    // Set column widths to match the layout
    sheet.setColumnWidth(1, 130); // A
    sheet.setColumnWidth(6, 130); // F


    var contactsSheet = ss.getSheetByName("contacts");
    var contactsListRange = contactsSheet.getRange("AR:AR");

    // Apply dropdown validation for "To:" and "Ship To:"
    var validation = SpreadsheetApp.newDataValidation().requireValueInRange(contactsListRange).build();
    sheet.getRange("B6:D6").merge().setFontWeight("bold").setDataValidation(validation); // "To:" dropdown
    sheet.getRange("B6").setValue("Company");
    sheet.getRange("G6:I6").merge().setFontWeight("bold").setDataValidation(validation); // "Ship To:" dropdown
    sheet.getRange("G6").setValue("Company");
    sheet.getRange("G15:I15").merge().setFontWeight("bold").setDataValidation(validation); // "Invoice To:" dropdown
    sheet.getRange("G15").setValue("Company");

    // Set headers and structure
    sheet.getRange("A1:J1").merge()
        .setValue("Your Company Name")
        .setFontWeight("bold")
        .setFontSize(20)
        .setHorizontalAlignment("center");
    sheet.getRange("A2:E2").merge()
        .setValue("Address")
        .setFontWeight("bold")
        .setFontSize(12)
        .setHorizontalAlignment("center");
    sheet.getRange("F2:J2").merge()
        .setValue("City, State Zip")
        .setFontWeight("bold")
        .setFontSize(12)
        .setHorizontalAlignment("center");

    sheet.getRange("A1:J46").setBorder(true, true, true, true, false, false);
    sheet.getRange("A3:J3").merge().setBackground('#cccccc').setBorder(true, true, true, true, false, false);
    sheet.getRange("F4:J17").setBorder(true, true, true, true, false, false);


    sheet.getRange("A4:B4").merge().setFontWeight("bold")
        .setValue("PURCHASE ORDER NUMBER:");
    sheet.getRange("F4:G4").merge().setFontWeight("bold").setValue("JOB/PHASE NUMBER:");

    sheet.getRange("A6").setValue("To:");
    sheet.getRange("B6").setFontWeight("bold"); // Company name (dropdown)

    sheet.getRange("F6").setValue("Ship To:");
    sheet.getRange("H6").setFontWeight("bold"); // Ship To company (dropdown)

    // Company Details
    sheet.getRange("B7:D7").merge().setFontWeight("bold").setFormula("=IFERROR(INDEX(contacts!AZ:AZ, MATCH(B6, contacts!AR:AR, 0)), \"Not Found\")");
    sheet.getRange("B8:D8").merge().setFontWeight("bold").setFormula(
  '=IFERROR(INDEX(contacts!BD:BD, MATCH(B6, contacts!AR:AR, 0)), "") & ", " & ' + 
  'IFERROR(INDEX(contacts!BE:BE, MATCH(B6, contacts!AR:AR, 0)), "") & " " & ' +
  'IFERROR(INDEX(contacts!BF:BF, MATCH(B6, contacts!AR:AR, 0)), "Not Found")'
);

    sheet.getRange("A9").setValue("Attn:");
    sheet.getRange("B9:D9").merge().setFontWeight("bold").setFormula("=IFERROR(INDEX(contacts!A:A, MATCH(B6, contacts!AR:AR, 0)), \"Not Found\")");

    sheet.getRange("G7:I7").merge().setFontWeight("bold").setFormula("=IFERROR(INDEX(contacts!AZ:AZ, MATCH(G6, contacts!AR:AR, 0)), \"Not Found\")");
    sheet.getRange("G8:I8").merge().setFontWeight("bold").setFormula(
  '=IFERROR(INDEX(contacts!BD:BD, MATCH(G6, contacts!AR:AR, 0)), "") & ", " & ' + 
  'IFERROR(INDEX(contacts!BE:BE, MATCH(G6, contacts!AR:AR, 0)), "") & " " & ' +
  'IFERROR(INDEX(contacts!BF:BF, MATCH(G6, contacts!AR:AR, 0)), "Not Found")'
);

    sheet.getRange("F10").setValue("Delivery-Site Phone:");
    sheet.getRange("G10:I10").merge().setFontWeight("bold").setValue("555-5555");

    sheet.getRange("F11").setValue("Site Contact:");
    sheet.getRange("G11:I11").merge().setFontWeight("bold").setValue("Joe Blow");

    sheet.getRange("A12").setValue("Phone:");
    sheet.getRange("B12:D12").merge().setFontWeight("bold").setFormula("=IFERROR(INDEX(contacts!AN:AN, MATCH(B6, contacts!AR:AR, 0)), \"Not Found\")");
    sheet.getRange("A13").setValue("Fax:");
    sheet.getRange("B13:D13").merge().setFontWeight("bold").setFormula("=IFERROR(INDEX(contacts!AP:AP, MATCH(B6, contacts!AR:AR, 0)), \"Not Found\")");
    sheet.getRange("A14").setValue("Email:");
    sheet.getRange("B14:D14").merge().setFontWeight("bold").setFormula("=IFERROR(INDEX(contacts!P:P, MATCH(B6, contacts!AR:AR, 0)), \"Not Found\")");

    sheet.getRange("A15").setValue("Ship VIA:");
    sheet.getRange("B15").setValue("Best Route");

    sheet.getRange("F15").setValue("Send Invoices To:");
    sheet.getRange("G16:I16").merge().setFontWeight("bold").setFormula("=IFERROR(INDEX(contacts!AZ:AZ, MATCH(G15, contacts!AR:AR, 0)), \"Not Found\")");
    sheet.getRange("G17:I17").merge().setFontWeight("bold").setFormula(
  '=IFERROR(INDEX(contacts!BD:BD, MATCH(G15, contacts!AR:AR, 0)), "") & ", " & ' + 
  'IFERROR(INDEX(contacts!BE:BE, MATCH(G15, contacts!AR:AR, 0)), "") & " " & ' +
  'IFERROR(INDEX(contacts!BF:BF, MATCH(G15, contacts!AR:AR, 0)), "Not Found")'
);

    sheet.getRange("A18").setValue("Delivery Required By:");
    sheet.getRange("B18").setValue("=TODAY()");
    sheet.getRange("E18").setValue("F.O.B.:");
    sheet.getRange("F18:G18").merge().setBorder(true, true, true, true, false, false).setValue("Delivery Site");
    sheet.getRange("H18:I18").merge().setBorder(true, true, true, true, false, false).setValue("SALES TAX EXEMPT:");
    sheet.getRange("J18").setBorder(true, true, true, true, false, false).setValue("YES");
    sheet.getRange("A19:J19").merge().setBackground('#cccccc').setBorder(true, true, true, true, false, false);

    // Description of Materials Section
    sheet.getRange("A20:G20").merge().setBorder(true, true, true, true, false, false).setValue("DESCRIPTION OF MATERIALS").setFontWeight("bold");
    sheet.getRange("H20").setBorder(true, true, true, true, false, false).setValue("Unit Price").setFontWeight("bold");
    sheet.getRange("I20").setBorder(true, true, true, true, false, false).setValue("Quantity").setFontWeight("bold");
    sheet.getRange("J20").setBorder(true, true, true, true, false, false).setValue("Amount").setFontWeight("bold");

    sheet.getRange("A21").setBorder(true, true, true, true, false, false).setValue("PROJECT:");
    sheet.getRange("B21:G21").merge().setBorder(true, true, true, true, false, false).setValue("Project Name").setFontStyle("italic");

    sheet.getRange("A22:G22").merge().setBorder(true, true, true, true, false, false).setValue("Fabricate and furnish the following materials per the plans and specifications prepared by");
    sheet.getRange("A23:G23").merge().setBorder(true, true, true, true, false, false).setValue("A/E Name").setFontWeight("bold");
    sheet.getRange("A24:G24").merge().setBorder(true, true, true, true, false, false).setValue("dated --/--/---- and all applicable addenda and correspondence").setFontStyle("italic");
    sheet.getRange("A25:G25").merge().setBorder(true, true, true, true, false, false).setFontWeight("bold");
    sheet.getRange("A26:G26").merge().setBorder(true, true, true, true, false, false).setFontWeight("bold");
    sheet.getRange("A27:G27").merge().setBorder(true, true, true, true, false, false).setFontWeight("bold");
    sheet.getRange("A28:G28").merge().setBorder(true, true, true, true, false, false).setFontWeight("bold");
    sheet.getRange("A29:G29").merge().setBorder(true, true, true, true, false, false).setFontWeight("bold");
    sheet.getRange("A30:G30").merge().setBorder(true, true, true, true, false, false).setFontWeight("bold");
    sheet.getRange("A31:G31").merge().setBorder(true, true, true, true, false, false).setFontWeight("bold");

    // Move "Unit Price, Quantity, Amount" section below
    sheet.getRange("H25").setBorder(true, true, true, true, false, false)
    sheet.getRange("I25").setBorder(true, true, true, true, false, false)
    sheet.getRange("J25").setFormula("=H25*I25");

    sheet.getRange("H26").setBorder(true, true, true, true, false, false)
    sheet.getRange("I26").setBorder(true, true, true, true, false, false)
    sheet.getRange("J26").setFormula("=H26*I26");

    sheet.getRange("H27").setBorder(true, true, true, true, false, false)
    sheet.getRange("I27").setBorder(true, true, true, true, false, false)
    sheet.getRange("J27").setFormula("=H27*I27");

    sheet.getRange("H28").setBorder(true, true, true, true, false, false)
    sheet.getRange("I28").setBorder(true, true, true, true, false, false)
    sheet.getRange("J28").setFormula("=H28*I28");

    sheet.getRange("H29").setBorder(true, true, true, true, false, false)
    sheet.getRange("I29").setBorder(true, true, true, true, false, false)
    sheet.getRange("J29").setFormula("=H29*I29");

    sheet.getRange("H30").setBorder(true, true, true, true, false, false)
    sheet.getRange("I30").setBorder(true, true, true, true, false, false)
    sheet.getRange("J30").setFormula("=H30*I30");

    sheet.getRange("H31").setBorder(true, true, true, true, false, false)
    sheet.getRange("I31").setBorder(true, true, true, true, false, false)
    sheet.getRange("J31").setFormula("=H31*I31");
    sheet.getRange("J21:J34").setBorder(true, true, true, true, false, false)
    sheet.getRange("H32:H34").setBorder(true, true, true, true, false, false)
    sheet.getRange("I32:I34").setBorder(true, true, true, true, false, false)

    sheet.getRange("G32").setValue("Subtotal").setFontWeight("bold");
    sheet.getRange("J32").setFormula("=SUM(J25:J31)").setFontWeight("bold");

    sheet.getRange("A33").setValue("EXCLUSIONS:");
    sheet.getRange("B34").setValue("none");
    sheet.getRange("B34:F35").merge()
		
    sheet.getRange("h35").setValue("GRAND TOTAL:").setFontWeight("bold");
    sheet.getRange("J35").setFormula("=SUM(J32)").setFontWeight("bold");

    sheet.getRange("A36:J36").merge().setBorder(true, true, true, true, false, false).setValue("Purchase Order Number must appear on all invoices, shipments, and correspondence");
    sheet.getRange("F37").setValue("Attachment Link:");
    sheet.getRange("A37:E37").merge().setValue("See attached sheet for additional terms and conditions of this offer")
    sheet.getRange("G37:J37").merge()

    sheet.getRange("A38:E38").merge().setValue("Payment Terms: Reference contract between the owner and");
    sheet.getRange("F38:J38").merge().setValue("Your Company Name");

    sheet.getRange("A39:J39").merge().setBorder(true, true, true, true, false, false)

    sheet.getRange("A40:J40").merge().setValue("Vendor is to supply a minimum of 24 hours advance notice of shipment of material");

    sheet.getRange("A42:B42").merge().setValue("Acknowledged By:");
    sheet.getRange("F42").setValue("Originated By:");

    sheet.getRange("C42:E42").merge()
    sheet.getRange("G42:J42").merge()
  
    sheet.getRange("A43").setValue("Vendor:");
    sheet.getRange("B43").setValue("Company");
    sheet.getRange("B43:E43").merge()

    sheet.getRange("A45").setValue("Date Signed:");
    sheet.getRange("F45").setValue("Date of Order");

    sheet.getRange("A46:J46").merge().setBorder(true, true, true, true, false, false).setValue("IMPORTANT: THIS OFFER DOES NOT BECOME AN ORDER UNTIL ALL COPIES ARE SIGNED AND BOTH COPIES ARE RETURNED TO THIS OFFICE.");

    // Insert two rows at the top
  sheet.insertRowsBefore(1, 2);

    Logger.log("Purchase Order Template Created Successfully");
}

function setup() {
  newfile();
  contacts();
  createInventoryTemplate();
  createInvoiceTemplate();
  createReceiptTemplate();
  createPackingSlipTemplate();
  createAdSheet();
  createFormSetupSheet();
  cleanupIT();
  
}

function cleanupIT() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName("Input");
  const logSheet = ss.getSheetByName("Log");
  const formSetupSheet = ss.getSheetByName("FormSetup");

  const mappings = [
    ["A1", "=A11"], ["A2", "=B11"], ["B1", "=D12"], ["B2", "=D13"],
    ["C1", "=A12"], ["C2", "=A13"], ["D1", "=C34"], ["D2", "=D34"],
    ["E1", "=C32"], ["E2", "=D32"], ["F1", "=H37"], ["F2", "=J37"],
    ["G1", "=A23"], ["G2", "=B23"], ["H1", "Log 8"],
    ["I1", "Log 9"], ["J1", "Log 10"], ["K1", "Log 11"], ["L1", "Log 12"],
    ["M1", "Date Paid"], ["N1", "Amount"], ["O1", "Date Shipped"]
  ];

  mappings.forEach(([cell, value]) => {
    inputSheet.getRange(cell).setValue(value);
  });

  logSheet.getRange("A2").setValue("Order Log");
  formSetupSheet.getRange("A10:J39").clearContent();
  formSetupSheet.getRange("I6").setValue("Store");
  formSetupSheet.getRange("B6").setValue("checkout");
  var sampleFields = [
    ["Form Header", "", "", "", "", "", "", "Header", 
     `<!DOCTYPE html>
<html lang="en" xmlns="http://www.w3.org/1999/xhtml">
<head>
    <meta charset="UTF-8">
    <title>My Custom Form</title>
    <style>
        .header {
            background-color: #2c3e50;
            color: #ecf0f1;
            font-family: Arial, sans-serif;
            padding: 20px;
        }

        .footer {
            background-color: #2c3e50;
            color: #ecf0f1;
            font-family: Arial, sans-serif;
            padding: 20px;
        }

        h2 {
            font-size: 28px;
            margin-bottom: 10px;
        }

        p {
            color: #ffffff;
            font-style: italic;
        }

        a {
            color: yellow; /* Sets link color to yellow */
        }

        .highlight {
            background-color: yellow;
            padding: 5px;
            border-radius: 3px;
            color: black;
        }
    </style>
</head>
<body>
    <h2>Welcome to Our Store</h2>
    <p><a href="https://datamateapp.github.io/" target="_blank">Website</a></p>
</body>
</html>`, "No"],
    ["Inventory", "", "", "", "", "", "", "Table", "Inventory!A1:G", "No"],
    ["Customer Information", "", "", "", "", "", "", "StaticText", "Customer Information", "No"],
    ["First Name", "NewContact", "B1", "", "", "", "", "Text", "", "Yes"],
    ["Last Name", "NewContact", "B3", "", "", "", "", "Text", "", "Yes"],
    ["Company", "NewContact", "B4", "", "", "", "", "Text", "", "No"],
    ["Email", "NewContact", "B9", "", "", "", "", "Email", "", "Yes"],
    ["Address Information", "", "", "", "", "", "", "StaticText", "Address Information", "No"],
    ["Bill to Street", "NewContact", "B12", "", "", "", "", "Text", "", "Yes"],
    ["Bill to City", "NewContact", "B13", "", "", "", "", "Text", "", "Yes"],
    ["Bill to State", "NewContact", "B14", "", "", "", "", "Text", "", "Yes"],
    ["Bill to Zip", "NewContact", "B15", "", "", "", "", "Text", "", "Yes"],
    ["Ship to Street", "NewContact", "B16", "", "", "", "", "Text", "", "Yes"],
    ["Ship to City", "NewContact", "B17", "", "", "", "", "Text", "", "Yes"],
    ["Ship to State", "NewContact", "B18", "", "", "", "", "Text", "", "Yes"],
    ["Ship to Zip", "NewContact", "B19", "", "", "", "", "Text", "", "Yes"],
    ["Container", "", "", "", "", "", "", "Container", "border: 2px dashed #4CAF50;", "No"],
    ["Checkout", "Orders", "A", "", "", "", "", "Checkout", "Inventory!A2:B", "Yes"],
    ["View Payment Instructions", "", "", "", "", "", "", "StaticText", "View Payment Instructions", "No"],
    ["Form Footer", "", "", "", "", "", "", "Footer", 
     `<!DOCTYPE html>
<html lang="en" xmlns="http://www.w3.org/1999/xhtml">
<head>
    <meta charset="UTF-8">
    <title>My Custom Form</title>
    <style>
        .header {
            background-color: #2c3e50;
            color: #ecf0f1;
            font-family: Arial, sans-serif;
            padding: 20px;
        }

        .footer {
            background-color: #2c3e50;
            color: #ecf0f1;
            font-family: Arial, sans-serif;
            padding: 20px;
        }

        h2 {
            font-size: 28px;
            margin-bottom: 10px;
        }

        p {
            color: #ffffff;
            font-style: italic;
        }

        .highlight {
            background-color: yellow;
            padding: 5px;
            border-radius: 3px;
            color: black;
        }
    </style>
</head>
<body>
    <h2>Your contribution helps us grow and improve!</h2>
    <p><a href="https://donate.stripe.com/14kdRf5mweVjg6IaEH?locale=en&__embed_source=buy_btn_1QO2WDG1lPcU42DNgACDVfYi" target="_blank">Payment Instructions</a></p>
</body>
</html>`, "No"]
  ];

  if (sampleFields.length > 0) {
    const lastRow = 10 + sampleFields.length - 1;
    formSetupSheet.getRange(`A10:J${lastRow}`).setValues(sampleFields)
      .setBackground("#ffffff")
      .setBorder(true, true, true, true, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);
  }

  try {
    var sheetsToMove = ['Update', 'Log', 'FormSetup', 'Input', 'View_Print', 'Packing Slip', 'Receipt', 'Inventory']; // Reverse order
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    // Loop through the sheets and move them to the front
    sheetsToMove.forEach(function(sheetName) {
      var sheet = spreadsheet.getSheetByName(sheetName);
      if (sheet) {
        spreadsheet.setActiveSheet(sheet);
        spreadsheet.moveActiveSheet(0); // Move to the front (index 0)
      }
    });

    // Activate 'FormSetup' at the end
    var sheet1 = spreadsheet.getSheetByName('FormSetup');
    if (sheet1) {
      sheet1.activate();
    } else {
      SpreadsheetApp.getUi().alert("Sheet1 not found.");
    }

    SpreadsheetApp.getUi().alert("Inventory Template created successfully. Please support DataMateApps and help us grow!");

  } catch (e) {
    SpreadsheetApp.getUi().alert(`Error: ${e.message}`);
  }
}






function createAdSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const adSheet = ss.getSheetByName("Advertisement") || ss.insertSheet("Advertisement");
  const inventory = ss.getSheetByName("Inventory");

  if (!inventory) {
    SpreadsheetApp.getUi().alert("Inventory sheet not found.");
    return;
  }

  adSheet.clear();
  adSheet.setHiddenGridlines(true);

  // Set column widths (A:H)
  const colWidths = [30, 150, 150, 80, 150, 150, 100, 100]; // A-H
  colWidths.forEach((w, i) => adSheet.setColumnWidth(i + 1, w));
  adSheet.setRowHeights(1, 48, 22);

  adSheet.getRange("A1:H1").merge().setValue(`<!DOCTYPE html>
<html lang="en" xmlns="http://www.w3.org/1999/xhtml">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>My Custom Form</title>
    <style>
        header {
            background: #35424a;
            color: #ffffff;
            padding: 10px 0;
            border-bottom: #e8491d 3px solid;
        }
        header a {
            color: #ffffff;
            text-decoration: none;
            text-transform: uppercase;
            font-size: 14px;
            padding: 5px;
            display: inline-block;
        }
        header h1 {
            margin: 0;
            font-size: 24px;
            text-align: center;
        }
        header nav {
            display: none;
        }
        .highlight, .current a {
            color: #e8491d;
            font-weight: bold;
        }
        header a:hover {
            color: #cccccc;
            font-weight: bold;
        }
        @media (min-width: 768px) {
            .container {
                width: 80%;
            }
            header h1 {
                font-size: 36px;
                text-align: left;
            }
            header nav {
                display: block;
                text-align: right;
            }
            header a {
                font-size: 16px;
                padding: 10px;
            }
            .feature-box {
                float: left;
                width: 30%;
            }
        }
    </style>
</head>
<body>
    <header class="header">
        <h1>DataMateApps</h1>
        <nav class="desktop-nav">
            <a href="https://datamateapp.github.io/">Home</a>
            <a href="https://datamateapp.github.io/features.html">Features</a>
            <a href="https://datamateapp.github.io/help.html" class="highlight">Help</a>
            <a href="https://script.google.com/macros/s/AKfycbx_WcbEWrHHVBd3Bg_7s7aipXoy7LeZwUxS-tyFJHo-h_uozwg4PRZBR3gnhxaFXJvLVg/exec">DataMate Counter</a>
            <a href="https://datamateapp.github.io/DownloadConstruction.html">Excel™ Versions</a>
            <a href="https://script.google.com/macros/s/AKfycbSRF3xIBTC-76ECjCrpb6GFcbqczOZHIZ8ZrGDBd88fDCyVtAEqH2sOiWS3ui1_ordQg/exec">Add-ins and Templates</a>
            <a href="https://datamateapp.github.io/Donate 5 per mo.html">Support DataMate</a>
            <a href="https://billing.stripe.com/p/login/bIY01q6q33nl1Ww4gg">Donor Portal</a>
            <a href="https://script.google.com/macros/s/AKfycbzWZM7qyNnKKw9T3LMODI5MQXMvzTUO9YvP-QvP49KipDphi48pR5iLhad1ZN-3FFqa/exec">Reviews</a>
            <a href="https://www.reddit.com/r/DataMateApps/">Community</a>
            <a href="https://drive.google.com/file/d/1yBphbReTNS95pbjNodU-USWuq-L48Jes/view?usp=drive_link">Open Source</a>
        </nav>
    </header>
</body>
</html>`);





  // Subtitle
  adSheet.getRange("A2:H2").merge().setValue("Limited-time deals pulled directly from our live inventory!")
    .setFontSize(12)
    .setFontStyle("italic")
    .setFontColor("#555555")
    .setHorizontalAlignment("center");

  // Headers (A-H)
  const headers = ["", "Description", "Price", "Qty", "Category", "Supplier", "Picture", "Video"];
  adSheet.getRange("A4:H4").setValues([headers])
    .setFontWeight("bold")
    .setBackground("#f4cccc")
    .setFontSize(11)
    .setHorizontalAlignment("center");

  // Pull top 12 products
  for (let i = 0; i < 12; i++) {
    const row = 5 + i;
    adSheet.getRange(row, 1).setValue("").setFontSize(16).setHorizontalAlignment("center");  // emoji
    adSheet.getRange(row, 2).setFormula(`=Inventory!A${2 + i}`);  // Description
    adSheet.getRange(row, 3).setFormula(`=Inventory!B${2 + i}`);  // Price
    adSheet.getRange(row, 4).setFormula(`=Inventory!C${2 + i}`);  // Qty
    adSheet.getRange(row, 5).setFormula(`=Inventory!D${2 + i}`);  // Category
    adSheet.getRange(row, 6).setFormula(`=Inventory!E${2 + i}`);  // Supplier
    adSheet.getRange(row, 7).setFormula(`=Inventory!F${2 + i}`);  // Picture
    adSheet.getRange(row, 8).setFormula(`=Inventory!G${2 + i}`);  // Video
  }

  // Borders and formatting
  adSheet.getRange("A4:H16").setBorder(true, true, true, true, true, true)
    .setFontSize(10)
    .setWrap(true)
    .setFontFamily("Arial");

  // Call to action
  adSheet.getRange("A18:H18").merge().setValue("Ready to order or learn more?")
    .setFontSize(12)
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  adSheet.getRange("A19:H19").merge().setValue("Watch our 2-minute demo →")
    .setFontSize(11)
    .setFontColor("#1155cc")
    .setHorizontalAlignment("center");

  adSheet.getRange("A20:H20").merge().setFormula('=HYPERLINK("https://youtu.be/_cduOVxVafc", "▶️ Watch the Video: How DataMate Works")')
    .setFontSize(12)
    .setFontColor("#0000FF")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  SpreadsheetApp.getUi().alert("Advertisement sheet created and synced to Inventory.");
}







function createInvoiceTemplate() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // Define ss
  var sheet = ss.getSheetByName('Sheet1') || ss.insertSheet('Sheet1');
  sheet.clear();

  // Header Section
  sheet.getRange('A1').setValue('Your Company Name').setFontSize(16).setFontWeight('bold');
  sheet.getRange('A3').setValue('Business Street').setFontSize(10);
  sheet.getRange('A4').setValue('Business City, Business State Business Postal Code').setFontSize(10);
  sheet.getRange('A5').setValue('E-mail Address').setFontSize(10);
  sheet.getRange('A6').setValue('Business Phone').setFontSize(10);
  sheet.getRange('A8:E8').merge().setValue('INVOICE').setFontSize(14).setFontWeight('bold');
  sheet.getRange('A9').setValue('Number').setFontSize(10).setFontWeight('bold');
  sheet.getRange('A10').setValue('Bill to:').setFontSize(10).setFontWeight('bold');

  // Client Information
  var contactsSheet = ss.getSheetByName("contacts"); 
  if (contactsSheet) { // Ensure the sheet exists before using it
    sheet.getRange("A11").setDataValidation(
      SpreadsheetApp.newDataValidation()
        .setAllowInvalid(false)
        .requireValueInRange(contactsSheet.getRange("A:A"), true) // Removed "$" for Apps Script compatibility
        .build()
    );
  } else {
    Logger.log("Error: 'contacts' sheet not found.");
  } 
  

  sheet.getRange('A12:B12').merge().setFormula("=VLOOKUP(A11, contacts!A:CJ, 44, FALSE)");
  sheet.getRange('A13:B13').merge().setFormula("=VLOOKUP(A11, contacts!A:CJ, 26, FALSE)");
  sheet.getRange('A14:B14').merge().setFormula(
    '=VLOOKUP(A11, contacts!A:CJ, 30, FALSE) & ", " & ' +
    'VLOOKUP(A11, contacts!A:CJ, 31, FALSE) & "   " & ' +
    'VLOOKUP(A11, contacts!A:CJ, 32, FALSE)'
  );
  sheet.getRange('A15:B15').merge().setFormula("=HYPERLINK(VLOOKUP(A11, contacts!A:CJ, 16, FALSE))");


  sheet.getRange('D10:E10').merge().setValue('Date:').setFontWeight('bold');
  sheet.getRange('D11:E11').merge().setFormula("=TODAY()");

  // Invoice Details Table Header
  sheet.getRange('A18').setValue('Description').setFontWeight('bold').setBackground('#cccccc');
  sheet.getRange('B18').setValue('Quantity').setFontWeight('bold').setBackground('#cccccc');
  sheet.getRange('C18').setValue('Unit Price').setFontWeight('bold').setBackground('#cccccc');
  sheet.getRange('D18').setValue('Total').setFontWeight('bold').setBackground('#cccccc');

  var inventorySheet = ss.getSheetByName("Inventory"); 
  if (inventorySheet) { // Ensure the sheet exists before using it
    sheet.getRange("A19:A28").setDataValidation(
      SpreadsheetApp.newDataValidation()
        .setAllowInvalid(false)
        .requireValueInRange(inventorySheet.getRange("A:A"), true) // Removed "$" for Apps Script compatibility
        .build()
    );
  } else {
    
    Logger.log("Error: 'Inventory' sheet not found.");
  }

 // Populate formulas for items in the invoice
for (var row = 19; row <= 28; row++) {
  sheet.getRange('C' + row)
       .setFormula(`=IFERROR(VLOOKUP(A${row},Inventory!$A$2:$CL$9341,2,FALSE), "")`)
       .setFontSize(10);
  sheet.getRange('D' + row)
       .setFormula(`=IF(C${row}="", "", $B${row}*$C${row})`)
       .setFontSize(10);
}


  // Summary Section
  sheet.getRange('C30').setValue('Subtotal:').setFontWeight('bold');
  sheet.getRange('D30').setFormula('=SUM(D19:D28)');
  
  sheet.getRange('B31').setValue('Tax:').setFontWeight('bold');
  sheet.getRange('C31').setValue('.08');
  sheet.getRange('D31').setFormula('=D30*C31');
  
  sheet.getRange('C32').setValue('Total:').setFontWeight('bold');
  sheet.getRange('D32').setFormula('=D30+D31');
  
  // Payment Instructions
  sheet.getRange('A34:E34').merge().setValue('Payment Instructions:').setFontWeight('bold');
  sheet.getRange('A35:E35').merge().setValue('[Your Payment Instructions]');

  // Formatting the sheet
  sheet.getRange('A1:E32').setHorizontalAlignment('center');
  sheet.getRange('A1:E6').setHorizontalAlignment('left');
  sheet.getRange('A9:B17').setHorizontalAlignment('left');
  sheet.getRange('A34:E35').setHorizontalAlignment('left');
  
  sheet.setColumnWidth(1, 350); // Set column A to 350
  sheet.setColumnWidths(2, 3, 100); // Set columns B, C, D to 100
  
  // Setting borders for the table
  sheet.getRange('A18:D28').setBorder(true, true, true, true, true, true);
  
  // Setting number formats
  sheet.getRange('C19:C28').setNumberFormat('$#,##0.00');
  sheet.getRange('D19:D28').setNumberFormat('$#,##0.00'); // Fixed range typo
  sheet.getRange('D30:D32').setNumberFormat('$#,##0.00');

  // Insert two rows at the top
  sheet.insertRowsBefore(1, 2);
}


function createReceiptTemplate() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('Receipt') || spreadsheet.insertSheet('Receipt');
  sheet.clear();

    // Header Section
  sheet.getRange('A1').setFormula("=View_Print!A3").setFontSize(16).setFontWeight('bold');
  sheet.getRange('A3').setFormula("=View_Print!A4").setFontSize(10);
  sheet.getRange('A4').setFormula("=View_Print!A5").setFontSize(10);
  sheet.getRange('A5').setFormula("=View_Print!A6").setFontSize(10);
  sheet.getRange('A6').setFormula("=View_Print!A7").setFontSize(10);
  sheet.getRange('A8:E8').merge().setValue('RECEIPT').setFontSize(14).setFontWeight('bold');

  
  // Client Information
  sheet.getRange('A10:B10').merge().setValue('Bill To:').setFontWeight('bold');
  sheet.getRange('A11:B11').merge().setFormula("=View_Print!A13");
  sheet.getRange('A12:B12').merge().setFormula("=VLOOKUP(A11, contacts!A:CJ, 44, FALSE)");
  sheet.getRange('A13:B13').merge().setFormula("=VLOOKUP(A11, contacts!A:CJ, 26, FALSE)");
  sheet.getRange('A14:B14').merge().setFormula(
  '=VLOOKUP(A11, contacts!A:CJ, 30, FALSE) & ", " & ' +
  'VLOOKUP(A11, contacts!A:CJ, 31, FALSE) & "   " & ' +
  'VLOOKUP(A11, contacts!A:CJ, 32, FALSE)'
);
  sheet.getRange('A15:B15').merge().setFormula("=HYPERLINK(VLOOKUP(A11, contacts!A:CJ, 16, FALSE))");


   
  sheet.getRange('D10:E10').merge().setValue('Date Paid:').setFontWeight('bold');
  sheet.getRange('D11:E11').merge().setFormula("=View_Print!M2");
  
  // Invoice Details Table Header
  sheet.getRange('A18').setValue('Description').setFontWeight('bold').setBackground('#cccccc');
  sheet.getRange('B18').setValue('Quantity').setFontWeight('bold').setBackground('#cccccc');
  sheet.getRange('C18').setValue('Unit Price').setFontWeight('bold').setBackground('#cccccc');
  sheet.getRange('D18').setValue('Total').setFontWeight('bold').setBackground('#cccccc');
  
  // Invoice Details Table (sample rows)
  sheet.getRange('A19').setFormula("=View_Print!A21").setFontSize(10);
  sheet.getRange('A20').setFormula("=View_Print!A22").setFontSize(10);
  sheet.getRange('A21').setFormula("=View_Print!A23").setFontSize(10);
  sheet.getRange('A22').setFormula("=View_Print!A24").setFontSize(10);
  sheet.getRange('A23').setFormula("=View_Print!A25").setFontSize(10);
  sheet.getRange('A24').setFormula("=View_Print!A26").setFontSize(10);
  sheet.getRange('A25').setFormula("=View_Print!A27").setFontSize(10);
  sheet.getRange('A26').setFormula("=View_Print!A28").setFontSize(10);
  sheet.getRange('A27').setFormula("=View_Print!A29").setFontSize(10);
  sheet.getRange('A28').setFormula("=View_Print!A30").setFontSize(10);

  sheet.getRange('B19').setFormula("=View_Print!B21").setFontSize(10);
  sheet.getRange('B20').setFormula("=View_Print!B22").setFontSize(10);
  sheet.getRange('B21').setFormula("=View_Print!B23").setFontSize(10);
  sheet.getRange('B22').setFormula("=View_Print!B24").setFontSize(10);
  sheet.getRange('B23').setFormula("=View_Print!B25").setFontSize(10);
  sheet.getRange('B24').setFormula("=View_Print!B26").setFontSize(10);
  sheet.getRange('B25').setFormula("=View_Print!B27").setFontSize(10);
  sheet.getRange('B26').setFormula("=View_Print!B28").setFontSize(10);
  sheet.getRange('B27').setFormula("=View_Print!B29").setFontSize(10);
  sheet.getRange('B28').setFormula("=View_Print!B30").setFontSize(10);

  sheet.getRange('C19').setFormula("=View_Print!C21").setFontSize(10);
  sheet.getRange('C20').setFormula("=View_Print!C22").setFontSize(10);
  sheet.getRange('C21').setFormula("=View_Print!C23").setFontSize(10);
  sheet.getRange('C22').setFormula("=View_Print!C24").setFontSize(10);
  sheet.getRange('C23').setFormula("=View_Print!C25").setFontSize(10);
  sheet.getRange('C24').setFormula("=View_Print!C26").setFontSize(10);
  sheet.getRange('C25').setFormula("=View_Print!C27").setFontSize(10);
  sheet.getRange('C26').setFormula("=View_Print!C28").setFontSize(10);
  sheet.getRange('C27').setFormula("=View_Print!C29").setFontSize(10);
  sheet.getRange('C28').setFormula("=View_Print!C30").setFontSize(10);

  sheet.getRange('D19').setFormula("=View_Print!D21").setFontSize(10);
  sheet.getRange('D20').setFormula("=View_Print!D22").setFontSize(10);
  sheet.getRange('D21').setFormula("=View_Print!D23").setFontSize(10);
  sheet.getRange('D22').setFormula("=View_Print!D24").setFontSize(10);
  sheet.getRange('D23').setFormula("=View_Print!D25").setFontSize(10);
  sheet.getRange('D24').setFormula("=View_Print!D26").setFontSize(10);
  sheet.getRange('D25').setFormula("=View_Print!D27").setFontSize(10);
  sheet.getRange('D26').setFormula("=View_Print!D28").setFontSize(10);
  sheet.getRange('D27').setFormula("=View_Print!D29").setFontSize(10);
  sheet.getRange('D28').setFormula("=View_Print!D30").setFontSize(10);
  
  
  // Summary Section
  sheet.getRange('C30').setValue('Subtotal:').setFontWeight('bold');
  sheet.getRange('D30').setValue('=View_Print!D32');
  
  sheet.getRange('B31').setValue('Tax:').setFontWeight('bold');
  sheet.getRange('C31').setValue('=View_Print!C33');
  sheet.getRange('D31').setValue('=View_Print!D33');
  
  sheet.getRange('C32').setValue('Total:').setFontWeight('bold');
  sheet.getRange('D32').setValue('=View_Print!D34');

  // Receipt Note
  sheet.getRange('A29:E29').merge().setValue('Thank you for your business!').setFontWeight('bold');
  
   // Formatting the sheet
  sheet.getRange('A1:E32').setHorizontalAlignment('center');
  sheet.getRange('A1:E6').setHorizontalAlignment('left');
  sheet.getRange('A9:B17').setHorizontalAlignment('left');
  sheet.getRange('A34:E35').setHorizontalAlignment('left');
  sheet.setColumnWidth(1, 350); // Set column A to 350
  sheet.setColumnWidths(2, 3, 100); // Set columns B, C, D to 100
  
  // Setting borders for the table
  sheet.getRange('A18:D28').setBorder(true, true, true, true, true, true);
  
  // Setting number formats
  sheet.getRange('C19:C28').setNumberFormat('$#,##0.00');
  sheet.getRange('D193:D28').setNumberFormat('$#,##0.00');
  sheet.getRange('D30:D32').setNumberFormat('$#,##0.00');
}

function createInventoryTemplate() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('Inventory');
  
  // If the sheet already exists, delete it
  if (sheet) {
    spreadsheet.deleteSheet(sheet);
  }
  
  // Create a new Inventory sheet
  sheet = spreadsheet.insertSheet('Inventory');

  // Set up the header row
  sheet.getRange('A1').setValue('Item Description').setFontWeight('bold');
  sheet.getRange('B1').setValue('Unit Price').setFontWeight('bold');;
  sheet.getRange('C1').setValue('Quantity in Stock').setFontWeight('bold');
  sheet.getRange('D1').setValue('Category').setFontWeight('bold');
  sheet.getRange('E1').setValue('Supplier').setFontWeight('bold');
  sheet.getRange('F1').setValue('Image').setFontWeight('bold');
  sheet.getRange('G1').setValue('Video').setFontWeight('bold');

  // Set column widths for better readability
  sheet.setColumnWidths(1, 5, 150);

  // Sample data
  var sampleData = [
    ['Item 1', 10.00, 100, 'Category 1', 'Supplier A', 'https://drive.google.com/uc?export=view&id=165kqv1atBk1WBbSkIbj6pnoikR9JOpLj', 'https://www.youtube.com/watch?v=dQw4w9WgXcQ'],
    ['Item 2', 15.00, 200, 'Category 2', 'Supplier B', 'https://drive.google.com/uc?export=view&id=165kqv1atBk1WBbSkIbj6pnoikR9JOpLj', 'https://www.youtube.com/watch?v=dQw4w9WgXcQ'],
    ['Item 3', 20.00, 150, 'Category 3', 'Supplier C', 'https://drive.google.com/uc?export=view&id=165kqv1atBk1WBbSkIbj6pnoikR9JOpLj', 'https://www.youtube.com/watch?v=dQw4w9WgXcQ']
  ];

  // Populate the sheet with sample data
  sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
}



function createPackingSlipTemplate() {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName('Packing Slip') || spreadsheet.insertSheet('Packing Slip');
    sheet.clear();

    // Header Section
    sheet.getRange('A1').setFormula("=View_Print!A3").setFontSize(16).setFontWeight('bold');
    sheet.getRange('A3').setFormula("=View_Print!A4").setFontSize(10);
    sheet.getRange('A4').setFormula("=View_Print!A5").setFontSize(10);
    sheet.getRange('A5').setFormula("=View_Print!A6").setFontSize(10);
    sheet.getRange('A6').setFormula("=View_Print!A7").setFontSize(10);

    sheet.getRange('A8:E8').merge().setValue('PACKING SLIP').setFontSize(14).setFontWeight('bold');

    // Client Information
    sheet.getRange('A10:B10').merge().setValue('Ship To:').setFontWeight('bold');
    sheet.getRange('A11:B11').merge().setFormula("=View_Print!A13");
    sheet.getRange('A12:B12').merge().setFormula("=VLOOKUP(A11, contacts!A:CJ, 44, FALSE)");
    sheet.getRange('A13:B13').merge().setFormula("=VLOOKUP(A11, contacts!A:CJ, 63, FALSE)");
    sheet.getRange('A14:B14').merge().setFormula(
      '=VLOOKUP(A11, contacts!A:CJ, 67, FALSE) & ", " & ' +
      'VLOOKUP(A11, contacts!A:CJ, 68, FALSE) & "   " & ' +
      'VLOOKUP(A11, contacts!A:CJ, 69, FALSE)'

    );
    sheet.getRange('A15:B15').merge().setFormula("=HYPERLINK(VLOOKUP(A11, contacts!A:CJ, 16, FALSE))");

    sheet.getRange('D10:E10').merge().setValue('Date Shipped:').setFontWeight('bold');
    sheet.getRange('D11:E11').merge().setFormula("=View_Print!O2");

    // Invoice Details Table Header
    sheet.getRange('A18').setValue('Description').setFontWeight('bold').setBackground('#cccccc');
    sheet.getRange('B18').setValue('Quantity').setFontWeight('bold').setBackground('#cccccc');
    sheet.getRange('C18').setValue('Unit Price').setFontWeight('bold').setBackground('#cccccc');
    sheet.getRange('D18').setValue('Total').setFontWeight('bold').setBackground('#cccccc');

    // Invoice Details Table (sample rows)
    for (var i = 19; i <= 28; i++) {
      sheet.getRange('A' + i).setFormula(`=View_Print!A${i + 2}`).setFontSize(10);
      sheet.getRange('B' + i).setFormula(`=View_Print!B${i + 2}`).setFontSize(10);
      sheet.getRange('C' + i).setFormula(`=View_Print!C${i + 2}`).setFontSize(10);
      sheet.getRange('D' + i).setFormula(`=View_Print!D${i + 2}`).setFontSize(10);
    }

    // Summary Section
    sheet.getRange('C30').setValue('Subtotal:').setFontWeight('bold');
    sheet.getRange('D30').setFormula('=View_Print!D32');

    sheet.getRange('B31').setValue('Tax:').setFontWeight('bold');
    sheet.getRange('C31').setFormula('=View_Print!C33');
    sheet.getRange('D31').setFormula('=View_Print!D33');

    sheet.getRange('C32').setValue('Total:').setFontWeight('bold');
    sheet.getRange('D32').setFormula('=View_Print!D34');

    // Note Section
    sheet.getRange('A29:E29').merge().setValue('Thank you for your business!').setFontWeight('bold');

    // Formatting the sheet
    sheet.getRange('A1:E32').setHorizontalAlignment('center');
    sheet.getRange('A1:E6').setHorizontalAlignment('left');
    sheet.getRange('A9:B17').setHorizontalAlignment('left');
    sheet.getRange('A34:E35').setHorizontalAlignment('left');
    sheet.setColumnWidth(1, 350); // Set column A to 350
    sheet.setColumnWidths(2, 3, 100); // Set columns B, C, D to 100

    // Setting borders for the table
    sheet.getRange('A18:D28').setBorder(true, true, true, true, true, true);

    // Setting number formats
    sheet.getRange('C19:C28').setNumberFormat('$#,##0.00');
    sheet.getRange('D19:D28').setNumberFormat('$#,##0.00');
    sheet.getRange('D30:D32').setNumberFormat('$#,##0.00');
    




  copyInput1()
  view()

  SpreadsheetApp.getUi().alert("Inventory Template created successfully. Please support DataMateApps and help us grow!");
  } catch (e) {
    SpreadsheetApp.getUi().alert(`Error: ${e.message}`);
  }
}


























function createFormSetupSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var formSetupSheet = ss.getSheetByName("FormSetup");
  if (!formSetupSheet) {
    formSetupSheet = ss.insertSheet("FormSetup");
    formSetupSheet.getRange("A1:Z100").setBackground("#e0e0e0");

    // Header
    formSetupSheet.getRange("A1:J1").merge()
      .setValue("DataMate FormBuilder Form Setup Dashboard")
      .setFontSize(18)
      .setFontWeight("bold")
      .setFontColor("#ffffff")
      .setBackground("#2c3e50")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");

    // Instructions
    const instructions = [
      {
        range: "A2:J2",
        value: "Configure your form below. Delete sample rows and add fields and targets in A10:J directly or use Form Builer.",
        background: "#e0e0e0"
      },
      {
        range: "A3:J3",
        value: "Enter comma-separated function names (e.g., checkout, newcontact, save, copyInput1) to run on form submission.",
        background: "#e8ecef"
      },
      {
        range: "A4:J4",
        value: "Enter the tax rate as a decimal (e.g., 0.08 for 8%) for checkout calculations.",
        background: "#e8ecef"
      },
      {
        range: "A5:J5",
        value: "Enter an email address to receive notifications for each form submission.",
        background: "#e8ecef"
      },
      {
        range: "D7:G7",
        value: "Tax rate used for Checkout field and checkout function.",
        background: "#e8ecef"
      },
      {
        range: "K1:O1",
        value: "All 29 Field Types:",
        background: "#e8ecef"
      }
    ];

    instructions.forEach(inst => {
      formSetupSheet.getRange(inst.range).merge()
        .setValue(inst.value)
        .setFontSize(12)
        .setFontColor("#666666")
        .setBackground(inst.background)
        .setHorizontalAlignment("center")
        .setWrap(true);
    });

    // On Submit Functions
    formSetupSheet.getRange("A6").setValue("On Submit Functions:");
    formSetupSheet.getRange("B6:F6").merge()
      .setFontSize(12)
      .setFontColor("#333333")
      .setBackground("#ffffff")
      .setHorizontalAlignment("left")
      .setNote("Enter comma-separated function names (e.g., checkout, newcontact, save, copyInput1) to run on form submission.");

    formSetupSheet.getRange("H6").setValue("Form Name:");
    formSetupSheet.getRange("I6:J6").merge()
      .setValue("Sample")
      .setFontSize(12)
      .setFontColor("#333333")
      .setBackground("#ffffff")
      .setHorizontalAlignment("left")
      .setNote("Enter your form name.");  

    // Tax Rate
    formSetupSheet.getRange("A7").setValue("Tax Rate:");
    formSetupSheet.getRange("B7:C7").merge()
      .setValue("0.08")
      .setFontSize(12)
      .setFontColor("#333333")
      .setBackground("#ffffff")
      .setNote("Tax rate used for Checkout field and checkout function.");

    // Email Notification
    formSetupSheet.getRange("A8").setValue("Email Notification:");
    formSetupSheet.getRange("B8:F8").merge()
      .setValue("")
      .setFontSize(12)
      .setFontColor("#333333")
      .setBackground("#ffffff")
      .setHorizontalAlignment("left")
      .setNote("Enter an email address to receive form submission notifications.");

    formSetupSheet.getRange("K2").setValue("Header");
    formSetupSheet.getRange("K3").setValue("Text");
    formSetupSheet.getRange("K4").setValue("Email");
    formSetupSheet.getRange("K5").setValue("Date");
    formSetupSheet.getRange("K6").setValue("Time");
    formSetupSheet.getRange("K7").setValue("Number");
    formSetupSheet.getRange("K8").setValue("Checkbox");
    formSetupSheet.getRange("L2").setValue("Radio");
    formSetupSheet.getRange("L3").setValue("Textarea");
    formSetupSheet.getRange("L4").setValue("Dropdown");
    formSetupSheet.getRange("L5").setValue("MultiSelect");
    formSetupSheet.getRange("L6").setValue("StarRating");
    formSetupSheet.getRange("L7").setValue("RangeSlider");
    formSetupSheet.getRange("L8").setValue("FileUpload*");
    formSetupSheet.getRange("M2").setValue("Conditional");
    formSetupSheet.getRange("M3").setValue("Calculated");
    formSetupSheet.getRange("M4").setValue("Signature*");
    formSetupSheet.getRange("M5").setValue("Geolocation");
    formSetupSheet.getRange("M6").setValue("ProgressBar");
    formSetupSheet.getRange("M7").setValue("Captcha");
    formSetupSheet.getRange("M8").setValue("Image");
    formSetupSheet.getRange("N2").setValue("Video");
    formSetupSheet.getRange("N3").setValue("ImageLink");
    formSetupSheet.getRange("N4").setValue("VideoLink");
    formSetupSheet.getRange("N5").setValue("StaticText");
    formSetupSheet.getRange("N6").setValue("Table");
    formSetupSheet.getRange("N7").setValue("Container");
    formSetupSheet.getRange("N8").setValue("Checkout");
    formSetupSheet.getRange("O2").setValue("Hyperlink");
    formSetupSheet.getRange("O2").setValue("Footer");
    formSetupSheet.getRange("O3").setValue("*FileUpload and Signature fields are available only when deployed as a web app with the open source code.");
    formSetupSheet.getRange("O4").setValue("Deploy as a Web App via Google Apps Script.");
    formSetupSheet.getRange("O5").setValue("Optional: Use Open Source Code From Our Website");
    formSetupSheet.getRange("O6").setValue("Copy the provided open-source code and paste it into the Apps Script Editor.");
    formSetupSheet.getRange("O7").setValue("In the Apps Script Editor, click Deploy > New Deployment, Select Web App,Configure the deployment, Click Deploy.");
    formSetupSheet.getRange("O8").setValue("Share the Web App URL with users to access the form directly in their browsers.");

    // Field Headers
    const headers = [
      "Form Fields", "Target Sheet 1", "Target Cell/Column 1",
      "Target Sheet 2", "Target Cell/Column 2", "Target Sheet 3",
      "Target Cell/Column 3", "Field Type", "Options", "Required"
    ];
    formSetupSheet.getRange("A9:J9").setValues([headers])
      .setFontWeight("bold")
      .setFontColor("#ffffff")
      .setBackground("#4CAF50")
      .setBorder(true, true, true, true, false, false);

    // Sample Fields (unchanged, just confirming it's correct)
    var sampleFields = [
      ["Form Header", "", "", "", "", "", "", "Header", `<!DOCTYPE html>
<html lang="en" xmlns="http://www.w3.org/1999/xhtml">
<head>
    <meta charset="UTF-8">
    <title>My Custom Form</title>
    <style>
        .header {
            background-color: #2c3e50;
            color: #ecf0f1;
            font-family: Arial, sans-serif;
            padding: 20px;
        }

        .footer {
            background-color: #2c3e50;
            color: #ecf0f1;
            font-family: Arial, sans-serif;
            padding: 20px;
        }

        h2 {
            font-size: 28px;
            margin-bottom: 10px;
        }

        p {
            color: #ffffff;
            font-style: italic;
        }

        a {
            color: yellow; /* Sets link color to yellow */
        }

        .highlight {
            background-color: yellow;
            padding: 5px;
            border-radius: 3px;
            color: black;
        }
    </style>
</head>
<body>
    <h2>Welcome</h2>
    <p>to <b>My Custom Form</b></p>
    <h2>Master Your Data with Ease</h2>
    <p>Transform your spreadsheets into powerful data management tools with DataMateApps.</p>
    <p>This demonstrates how custom HTML can be used in a form header.</p>
    
    <span class="highlight">Below are 27 of all 29 field types.</span>
</body>
</html>`, "No"],
      ["Name", "Responses", "A", "Sheet2", "B2", "", "", "Text", "", "Yes"],
      ["Email", "Responses", "B", "", "", "", "", "Email", "", "Yes"],
      ["Date", "Responses", "C", "Records", "B1", "", "", "Date", "", "No"],
      ["Time", "Responses", "D", "", "", "", "", "Time", "", "No"],
      ["Number", "Responses", "E", "", "", "", "", "Number", "", "No"],
      ["Checkbox", "Responses", "F", "", "", "", "", "Checkbox", "", "No"],
      ["Radio", "Responses", "G", "", "", "", "", "Radio", "Yes,No,Maybe", "No"],
      ["Textarea", "Responses", "H", "", "", "", "", "Textarea", "", "No"],
      ["Dropdown", "Responses", "I", "", "", "", "", "Dropdown", "=Sheet1!A:A", "No"],
      ["MultiSelect", "Responses", "J", "", "", "", "", "MultiSelect", "Red,Green,Blue", "No"],
      ["StarRating", "Responses", "K", "", "", "", "", "StarRating", "", "No"],
      ["RangeSlider", "Responses", "L", "", "", "", "", "RangeSlider", "0,100,5", "No"],
      ["Conditional", "Responses", "N", "", "", "", "", "Conditional", "Checkbox=true", "No"],
      ["Calculated", "Responses", "O", "", "", "", "", "Calculated", "=Number*2", "No"],
      ["Geolocation", "Responses", "Q", "", "", "", "", "Geolocation", "", "No"],
      ["ProgressBar", "", "", "", "", "", "", "ProgressBar", "75", "No"],
      ["Captcha", "Responses", "R", "", "", "", "", "Captcha", "", "No"],
      ["Image", "", "", "", "", "", "", "Image", "https://drive.google.com/uc?export=view&id=165kqv1atBk1WBbSkIbj6pnoikR9JOpLj", "No"],
      ["Video", "", "", "", "", "", "", "Video", "https://www.youtube.com/watch?v=dQw4w9WgXcQ", "No"],
      ["ImageLink", "Responses", "S", "", "", "", "", "ImageLink", "", "No"],
      ["VideoLink", "Responses", "T", "", "", "", "", "VideoLink", "", "No"],
      ["StaticText", "", "", "", "", "", "", "StaticText", "This is static text", "No"],
      ["Table", "", "", "", "", "", "", "Table", "Sheet1!A1:F10", "No"],
      ["Container", "", "", "", "", "", "", "Container", "border: 2px dashed #4CAF50;", "No"],
      ["Checkout", "Orders", "A", "", "", "", "", "Checkout", "Sheet1!A2:B10", "No"],
      ["Hyperlink", "", "", "", "", "", "", "Hyperlink", "https://datamateapp.github.io/Donate%205%20per%20mo.html", "No"],
      ["Form Footer", "", "", "", "", "", "", "Footer", "UploadFile and Signature fields require Drive permissions authorized and will only submit if deployed as a web app. Sheet1 must contain data in columns A 'Descripton' and B 'Price' for the Checkout and Table feilds.", "No"]
    ];

    if (sampleFields.length > 0) {
      const lastRow = 10 + sampleFields.length - 1;
      formSetupSheet.getRange(`A10:J${lastRow}`).setValues(sampleFields)
        .setBackground("#ffffff")
        .setBorder(true, true, true, true, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);
    }

    // Formatting
    formSetupSheet.setFrozenRows(9);
    const columnWidths = [150, 100, 100, 100, 100, 100, 100, 100, 150, 80];
    columnWidths.forEach((width, index) => formSetupSheet.setColumnWidth(index + 1, width));
  }
  return formSetupSheet;
}



function previewForm() {
  var html = generateFormHTML();
  html.setWidth(1200).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, html.getTitle());
}



function showFormBuilder() {
  var html = HtmlService.createHtmlOutputFromFile('FormBuilder')
    .setWidth(1200)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Form Builder');
}








// Store configuration in ScriptProperties
function saveSnapshotConfig(data) {
  const properties = PropertiesService.getScriptProperties();
  properties.setProperty('snapshotConfig', JSON.stringify(data));
}

// Retrieve configuration
function getSnapshotConfig() {
  const properties = PropertiesService.getScriptProperties();
  const config = properties.getProperty('snapshotConfig');
  return config ? JSON.parse(config) : null;
}

// Modified copyCameraSnapshot to save config
function copyCameraSnapshot(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName(data.sourceSheet);
  const targetSheet = ss.getSheetByName(data.targetSheet);

  if (!sourceSheet || !targetSheet) {
    throw new Error("Sheet not found.");
  }

  const sourceRange = sourceSheet.getRange(data.sourceRange);
  const values = sourceRange.getValues();
  const bgs = sourceRange.getBackgrounds();
  const fontColors = sourceRange.getFontColors();
  const fontWeights = sourceRange.getFontWeights();
  const fontStyles = sourceRange.getFontStyles();
  const hAligns = sourceRange.getHorizontalAlignments();
  const vAligns = sourceRange.getVerticalAlignments();

  const numRows = values.length;
  const numCols = values[0].length;

  const targetStart = targetSheet.getRange(data.targetCell);
  const targetRange = targetSheet.getRange(
    targetStart.getRow(),
    targetStart.getColumn(),
    numRows,
    numCols
  );

  targetRange.clearContent().clearFormat();

  targetRange.setValues(values);
  targetRange.setBackgrounds(bgs);
  targetRange.setFontColors(fontColors);
  targetRange.setFontWeights(fontWeights);
  targetRange.setFontStyles(fontStyles);
  targetRange.setHorizontalAlignments(hAligns);
  targetRange.setVerticalAlignments(vAligns);

  // Save the configuration for real-time syncing
  saveSnapshotConfig(data);
}

// onEdit trigger to sync changes
function onEdit(e) {
  const config = getSnapshotConfig();
  if (!config) return; // No snapshot configured

  const range = e.range;
  const sheet = range.getSheet();
  const sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(config.sourceSheet);

  if (sheet.getName() !== config.sourceSheet) return; // Edit is not in source sheet

  // Check if the edited range overlaps with the source range
  const sourceRange = sourceSheet.getRange(config.sourceRange);
  const editedRowStart = range.getRow();
  const editedRowEnd = range.getLastRow();
  const editedColStart = range.getColumn();
  const editedColEnd = range.getLastColumn();
  const sourceRowStart = sourceRange.getRow();
  const sourceRowEnd = sourceRange.getLastRow();
  const sourceColStart = sourceRange.getColumn();
  const sourceColEnd = sourceRange.getLastColumn();

  if (
    editedRowStart <= sourceRowEnd &&
    editedRowEnd >= sourceRowStart &&
    editedColStart <= sourceColEnd &&
    editedColEnd >= sourceColStart
  ) {
    // Edited range overlaps with source range, so update the target
    copyCameraSnapshot(config);
  }
}



function showCameraSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('CameraSidebar')
    .setTitle('Camera Tool');
  SpreadsheetApp.getUi().showSidebar(html);
}

function getSheetNames() {
  return SpreadsheetApp.getActiveSpreadsheet()
    .getSheets()
    .map(sheet => sheet.getName());
}




















//The Apps Script below is open-source and editable. It is required for Web App Deployment. Custom functions can be added and called on the DataMate FormBuilder Form Setup Dashboard.//


//This checkout function is a custom function made for the Inventory Template. It is a good example of how custom functions can automate the form to spreadsheet processes upon form submission.//

function checkout() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName("Input");
  const ordersSheet = ss.getSheetByName("Orders");
  const viewSheet = ss.getSheetByName("View_Print");

  // Set formulas for A21:A30
  inputSheet.getRange("A21").setFormula("=Orders!A2");
  inputSheet.getRange("A22").setFormula("=Orders!A3");
  inputSheet.getRange("A23").setFormula("=Orders!A4");
  inputSheet.getRange("A24").setFormula("=Orders!A5");
  inputSheet.getRange("A25").setFormula("=Orders!A6");
  inputSheet.getRange("A26").setFormula("=Orders!A7");
  inputSheet.getRange("A27").setFormula("=Orders!A8");
  inputSheet.getRange("A28").setFormula("=Orders!A9");
  inputSheet.getRange("A29").setFormula("=Orders!A10");
  inputSheet.getRange("A30").setFormula("=Orders!A11"); 

  // Set formulas for B21:B30
  inputSheet.getRange("B21").setFormula("=Orders!B2");
  inputSheet.getRange("B22").setFormula("=Orders!B3");
  inputSheet.getRange("B23").setFormula("=Orders!B4");
  inputSheet.getRange("B24").setFormula("=Orders!B5");
  inputSheet.getRange("B25").setFormula("=Orders!B6");
  inputSheet.getRange("B26").setFormula("=Orders!B7");
  inputSheet.getRange("B27").setFormula("=Orders!B8");
  inputSheet.getRange("B28").setFormula("=Orders!B9");
  inputSheet.getRange("B29").setFormula("=Orders!B10");
  inputSheet.getRange("B30").setFormula("=Orders!B11");

  // Set other formulas
  inputSheet.getRange("B11").setFormula("=Log!A10+1");
  

  // Call other functions
  newcontact();
  inputSheet.getRange("A13").clearDataValidations().setFormula("=contacts!A2");
  save();
  updateInventory();
  copyInput1();


  // Clear the Orders sheet range A1:C12
  ordersSheet.getRange("A1:C").clear();

  // Reapply formulas (if needed)

  inputSheet.getRange("B11").setFormula("=Log!A10+1");

}

//This updateInventory function is a custom function made for the Invoice Template. It is a good example of how custom functions can automate the form to spreadsheet processes upon form submission.//

function updateInventory() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var inventorySheet = spreadsheet.getSheetByName('Inventory');
  var invoiceSheet = spreadsheet.getSheetByName('Input');

  if (!inventorySheet || !invoiceSheet) {
    Logger.log('Missing required sheets: Inventory or Input');
    return;
  }

  // Get invoice data
  var invoiceData = invoiceSheet.getRange('A21:D30').getValues();

  // Loop through invoice data to process each item
  for (var i = 0; i < invoiceData.length; i++) {
    var itemDescription = invoiceData[i][0];
    var quantitySold = invoiceData[i][1];

    if (itemDescription && quantitySold) {
      // Get inventory data (updated to include column C)
      var inventoryData = inventorySheet.getRange('A2:C' + inventorySheet.getLastRow()).getValues();

      for (var j = 0; j < inventoryData.length; j++) {
        if (inventoryData[j][0] == itemDescription) {
          var currentStock = inventoryData[j][2]; // Changed from [1] to [2] for column C

          if (typeof currentStock === 'number' && currentStock >= quantitySold) {
            // Update inventory stock in column C
            inventorySheet.getRange('C' + (j + 2)).setValue(currentStock - quantitySold);
          } else {
            Logger.log('Insufficient stock for item: ' + itemDescription);
          }
          break; // Exit inner loop once match is found
        }
      }
    }
  }
}

function saveAndUpdate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName("Input");

  // Set other formulas
  inputSheet.getRange("B11").setFormula("=Log!A10+1");

  // Call other functions
  save();
  updateInventory();
  copyInput1();

  // Reapply formulas (if needed)

  inputSheet.getRange("B11").setFormula("=Log!A10+1");

}





function generateFormHTML() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var setupSheet = ss.getSheetByName("FormSetup");
  if (!setupSheet) {
    createFormSetupSheet();
    setupSheet = ss.getSheetByName("FormSetup");
  }

  var fieldsRange = setupSheet.getRange("A10:J" + setupSheet.getLastRow());
  var fieldsData = fieldsRange.getValues().filter(row => row[0] !== "");
  var taxRate = parseFloat(setupSheet.getRange("B7").getValue()) || 0.08;
  var formName = setupSheet.getRange("I6").getValue() || "Custom";

  var processedFieldsData = fieldsData.map((row, index) => {
    var fieldName = row[0];
    var fieldType = row[7] || "Text";
    var cell = setupSheet.getRange("I" + (index + 10));
    var options = cell.getFormula() || row[8] || "";
    var required = row[9].toString().toLowerCase() === "yes";
    options = String(options);
    var targets = [
      { sheet: row[1], cell: row[2] },
      { sheet: row[3], cell: row[4] },
      { sheet: row[5], cell: row[6] }
    ].filter(t => t.sheet && t.cell);

    var fieldOptions = [];
    if (fieldType.toUpperCase() === "CHECKOUT" && options) {
      try {
        var range = ss.getRange(options);
        fieldOptions = range.getValues().map(r => ({
          description: String(r[0] || ""),
          unitPrice: Number(r[1]) || 0
        })).filter(item => item.description);
      } catch (e) {
        fieldOptions = [{ description: "Error: Invalid range " + options, unitPrice: 0 }];
      }
    } else if (["DROPDOWN", "RADIO", "MULTISELECT"].includes(fieldType.toUpperCase())) {
      if (options.startsWith("=")) {
        try {
          var range = ss.getRange(options.substring(1));
          fieldOptions = range.getValues().flat().filter(String);
        } catch (e) {
          fieldOptions = ["Error: Invalid range " + options];
        }
      } else if (options) {
        fieldOptions = options.split(",");
      }
    } else if (["FILEUPLOAD", "CONDITIONAL", "CALCULATED", "STATICTEXT", "PROGRESSBAR", "CONTAINER", "HEADER", "FOOTER"].includes(fieldType.toUpperCase())) {
      fieldOptions = [options];
    } else if (fieldType.toUpperCase() === "HYPERLINK" && options) {
      var hrefMatch = options.match(/href=["'](.*?)["']/i);
      fieldOptions = [options, hrefMatch ? hrefMatch[1] : options];
    } else if (fieldType.toUpperCase() === "RANGESLIDER" && options) {
      var parts = options.split(",");
      fieldOptions = parts.length === 3 ? parts.map(Number) : [0, 100, 1];
    } else if (["IMAGE", "VIDEO"].includes(fieldType.toUpperCase()) && options) {
      if (options.includes("drive.google.com")) {
        var fileIdMatch = options.match(/([a-zA-Z0-9_-]{20,})/);
        if (fileIdMatch) options = "https://drive.google.com/thumbnail?id=" + fileIdMatch[1];
      }
      fieldOptions = [options];
    } else if (fieldType.toUpperCase() === "TABLE" && options) {
      try {
        var range = ss.getRange(options);
        fieldOptions = range.getValues().filter(row => 
          row.some(cell => String(cell || '').trim() !== '')
        ).map(row => row.map(cell => String(cell || '')));
      } catch (e) {
        fieldOptions = [["Error: Invalid range " + options]];
      }
    }

    return [fieldName, fieldType, fieldOptions, targets, required];
  });

  var additionalStyles = [];
  var hasCustomHeader = false;
  var hasCustomFooter = false;
  processedFieldsData.forEach(field => {
    if (["HEADER", "FOOTER"].includes(field[1].toUpperCase()) && field[2][0]) {
      var options = field[2][0];
      if (options.match(/<!DOCTYPE|<html/i)) {
        var styleMatch = options.match(/<style[^>]*>([\s\S]*?)<\/style>/i);
        if (styleMatch) additionalStyles.push(styleMatch[1]);
        if (field[1].toUpperCase() === "HEADER") hasCustomHeader = true;
        if (field[1].toUpperCase() === "FOOTER") hasCustomFooter = true;
      }
    }
  });

  var template = HtmlService.createTemplate(`
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body {
            font-family: 'Roboto', Arial, sans-serif;
            margin: 0;
            padding: 20px;
            background: #f5f5f5;
            color: #333;
          }
          .container {
            max-width: 1200px;
            margin: 0 auto;
            background: white;
            padding: 30px;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
          }
          .custom-container {
            padding: 20px;
            margin-bottom: 20px;
            border: 1px solid #ddd;
            border-radius: 4px;
          }
          .header {
            ${hasCustomHeader ? '' : 'background: #ffffff;'}
            color: ${hasCustomHeader ? 'white' : 'black'};
            padding: 15px;
            text-align: center;
            border-radius: 4px;
            margin-bottom: 20px;
            font-size: 24px;
          }
          .footer {
            ${hasCustomFooter ? '' : 'background: #ffffff;'}
            color: ${hasCustomHeader ? 'white' : 'black'};
            padding: 15px;
            text-align: center;
            border-radius: 4px;
            margin-top: 20px;
            font-size: 14px;
          }

          .form-group {
            margin-bottom: 20px;
            display: flex;
            align-items: center;
            flex-wrap: wrap;
          }
          label {
            width: 150px;
            font-weight: 500;
            margin-right: 15px;
            color: #555;
          }
          input[type="text"], input[type="date"], input[type="number"], 
          input[type="email"], input[type="time"], input[type="range"], 
          select, textarea, input[type="file"] {
            width: 250px;
            padding: 8px 12px;
            border: 1px solid #ddd;
            border-radius: 4px;
            font-size: 14px;
            transition: border-color 0.3s;
          }
          input:focus, select:focus, textarea:focus {
            border-color: #4CAF50;
            outline: none;
          }
          textarea {
            resize: vertical;
            min-height: 100px;
          }
          input[type="checkbox"] {
            margin-left: 150px;
          }
          .radio-group {
            display: flex;
            flex-direction: column;
            gap: 8px;
          }
          .radio-group label {
            width: auto;
            margin: 0 0 0 5px;
            display: inline;
          }
          .star-rating {
            display: inline-flex;
            font-size: 28px;
            direction: rtl;
          }
          .star-rating input[type="radio"] {
            display: none;
          }
          .star-rating label {
            color: #ddd;
            cursor: pointer;
            margin: 0 3px;
            width: auto;
            transition: color 0.2s;
          }
          .star-rating label:hover,
          .star-rating label:hover ~ label,
          .star-rating input[type="radio"]:checked ~ label {
            color: #f5b301;
          }
          img, video, iframe {
            max-width: 250px;
            max-height: 250px;
            margin-top: 10px;
            border-radius: 4px;
          }
          .static-text {
            width: 100%;
            padding: 15px;
            background: #f9f9f9;
            border-left: 4px solid #4CAF50;
            border-radius: 4px;
            margin: 0 0 20px 150px;
            font-size: 16px;
            color: #444;
          }
          .table-display {
            width: 100%;
            margin: 0 0 20px 0;
            border-collapse: collapse;
            background: #fff;
            border: 1px solid #ddd;
          }
          .table-display th, .table-display td {
            padding: 10px;
            border: 1px solid #ddd;
            text-align: left;
          }
          .table-display th {
            background: #f1f1f1;
            font-weight: bold;
          }
          .table-display img {
            width: 100px;
            height: auto;
            display: block;
          }
          .table-display iframe {
            width: 200px;
            height: 150px;
            border: none;
          }
          .hyperlink {
            margin-left: 150px;
            color: #4CAF50;
            text-decoration: none;
            font-size: 16px;
          }
          .hyperlink:hover {
            text-decoration: underline;
          }
          .range-output {
            margin-left: 10px;
            font-size: 14px;
            color: #666;
          }
          .conditional-field {
            display: none;
          }
          .calculated-field {
            background: #f9f9f9;
            pointer-events: none;
          }
          canvas {
            border: 1px solid #ddd;
            border-radius: 4px;
            width: 250px;
            height: 100px;
          }
          progress {
            width: 250px;
            height: 20px;
          }
          button {
            padding: 12px 25px;
            background: #4CAF50;
            color: white;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            font-size: 16px;
          }
          button:hover:not(:disabled) {
            background: #45a049;
          }
          button:disabled {
            background: #cccccc;
            cursor: not-allowed;
          }
          .spinner {
            display: none;
            border: 4px solid #f3f3f3;
            border-top: 4px solid #3498db;
            border-radius: 50%;
            width: 20px;
            height: 20px;
            animation: spin 1s linear infinite;
            position: absolute;
            top: 50%;
            left: 50%;
            transform: translate(-50%, -50%);
          }
          @keyframes spin {
            0% { transform: translate(-50%, -50%) rotate(0deg); }
            100% { transform: translate(-50%, -50%) rotate(360deg); }
          }
          #message {
            color: #4CAF50;
            text-align: center;
            margin-top: 20px;
            font-size: 16px;
            display: none;
          }
          .error {
            color: #d32f2f;
            font-size: 12px;
            margin-left: 165px;
            margin-top: 5px;
          }
          .no-fields {
            text-align: center;
            color: #666;
            font-size: 16px;
            padding: 20px;
          }
          .required::after {
            content: " *";
            color: #d32f2f;
          }
          .checkout-table {
            width: 100%;
            border-collapse: collapse;
            margin: 10px 0 10px 0;
            border: 1px solid #ddd;
          }
          .checkout-table th, .checkout-table td {
            padding: 12px;
            border: 1px solid #ddd;
            text-align: left;
          }
          .checkout-table th {
            background: #e8491d;
            color: white;
          }
          .checkout-table tr:nth-child(even) {
            background: #f2f2f2;
          }
          .checkout-table select {
            width: 100%;
            padding: 8px;
            border: 1px solid #ddd;
            border-radius: 4px;
          }
          .checkout-table input[type="number"] {
            width: 60px;
            padding: 8px;
            border: 1px solid #ddd;
            border-radius: 4px;
          }
          .calculation-field {
            margin-left: 150px;
            margin-bottom: 10px;
          }
          .calculation-field span {
            display: inline-block;
            width: 100px;
          }
          .add-item-btn {
            background: #2196F3;
            margin-left: 150px;
            margin-top: 10px;
            padding: 8px 16px;
            border-radius: 4px;
          }
          .remove-item-btn {
            background: #e74c3c;
            padding: 6px 12px;
            border-radius: 4px;
          }
          .hyperlink::after {
  content: ' 🔗';
  font-size: 0.9em;
  color: #888;
}
          ${additionalStyles.join('\n')}
        </style>
      </head>
      <body>
        <div class="container">
          <? if (processedFieldsData.length > 0) { ?>
            <form id="myForm" onsubmit="handleSubmit(event)" enctype="multipart/form-data">
              <? var inContainer = false; ?>
              <? for (var i = 0; i < processedFieldsData.length; i++) { ?>
                <? if (processedFieldsData[i][1].toUpperCase() === "HEADER" && processedFieldsData[i][2][0]) { ?>
                  <? var options = processedFieldsData[i][2][0]; ?>
                  <? if (options.match(/<!DOCTYPE|<html/i)) { ?>
                    <? var bodyMatch = options.match(/<body[^>]*>([\\s\\S]*?)<\\/body>/i); ?>
                    <? if (bodyMatch) { ?>
                      <div class="header"><?!= bodyMatch[1] ?></div>
                    <? } else { ?>
                      <div class="header"><?= processedFieldsData[i][0] ?></div>
                    <? } ?>
                  <? } else if (options.includes(':')) { ?>
                    <div class="header" style="<?= options ?>"><?= processedFieldsData[i][0] ?></div>
                  <? } else { ?>
                    <div class="header"><?!= options ?></div>
                  <? } ?>
                <? } else if (processedFieldsData[i][1].toUpperCase() === "FOOTER" && processedFieldsData[i][2][0]) { ?>
                  <? if (inContainer) { ?></div><? inContainer = false; } ?>
                  <? var options = processedFieldsData[i][2][0]; ?>
                  <? if (options.match(/<!DOCTYPE|<html/i)) { ?>
                    <? var bodyMatch = options.match(/<body[^>]*>([\\s\\S]*?)<\\/body>/i); ?>
                    <? if (bodyMatch) { ?>
                      <div class="footer"><?!= bodyMatch[1] ?></div>
                    <? } else { ?>
                      <div class="footer"><?= processedFieldsData[i][0] ?></div>
                    <? } ?>
                  <? } else if (options.includes(':')) { ?>
                    <div class="footer" style="<?= options ?>"><?= processedFieldsData[i][0] ?></div>
                  <? } else { ?>
                    <div class="footer"><?!= options ?></div>
                  <? } ?>
                <? } else if (processedFieldsData[i][1].toUpperCase() === "CONTAINER" && processedFieldsData[i][2][0]) { ?>
                  <? if (inContainer) { ?></div><? } ?>
                  <div class="custom-container" style="<?= processedFieldsData[i][2][0] ?>">
                  <? inContainer = true; ?>
                <? } else if (processedFieldsData[i][1].toUpperCase() === "CHECKOUT" && processedFieldsData[i][2].length > 0) { ?>
                  <div class="form-group" id="group-<?= processedFieldsData[i][0] ?>">
                    <div>
                      <table class="checkout-table" id="<?= processedFieldsData[i][0] ?>-table">
                        <thead>
                          <tr>
                            <th>Description</th>
                            <th>Quantity</th>
                            <th>Unit Price</th>
                            <th>Total</th>
                            <th>Action</th>
                          </tr>
                        </thead>
                        <tbody id="<?= processedFieldsData[i][0] ?>-tbody">
                          <tr id="<?= processedFieldsData[i][0] ?>-row-0">
                            <td>
                              <select name="description" onchange="updateCheckoutTotals('<?= processedFieldsData[i][0] ?>')">
                                <option value="">Select an item</option>
                                <? for (var j = 0; j < processedFieldsData[i][2].length; j++) { ?>
                                  <option value='<?= JSON.stringify({ description: processedFieldsData[i][2][j].description, unitPrice: processedFieldsData[i][2][j].unitPrice }) ?>'><?= processedFieldsData[i][2][j].description ?></option>
                                <? } ?>
                              </select>
                            </td>
                            <td><input type="number" name="quantity" min="0" value="0" oninput="updateCheckoutTotals('<?= processedFieldsData[i][0] ?>')"></td>
                            <td class="unitPrice">$0.00</td>
                            <td class="itemTotal">$0.00</td>
                            <td><button type="button" class="remove-item-btn" onclick="removeCheckoutItem('<?= processedFieldsData[i][0] ?>', '<?= processedFieldsData[i][0] ?>-row-0')">Remove</button></td>
                          </tr>
                        </tbody>
                      </table>
                      <div class="calculation-field"><span>Subtotal:</span><span id="<?= processedFieldsData[i][0] ?>-subtotal">$0.00</span></div>
                      <div class="calculation-field"><span>Tax (${(taxRate * 100).toFixed(2)}%):</span><span id="<?= processedFieldsData[i][0] ?>-tax">$0.00</span></div>
                      <div class="calculation-field"><span>Total:</span><span id="<?= processedFieldsData[i][0] ?>-total">$0.00</span></div>
                      <button type="button" class="add-item-btn" onclick="addCheckoutItem('<?= processedFieldsData[i][0] ?>')">Add Item</button>
                      <input type="hidden" id="<?= processedFieldsData[i][0] ?>" name="<?= processedFieldsData[i][0] ?>" <?= processedFieldsData[i][4] ? 'required' : '' ?>>
                      <span class="error" id="<?= processedFieldsData[i][0] ?>-error"></span>
                    </div>
                  </div>
                <? } else { ?>
                  <div class="form-group <?= processedFieldsData[i][1].toUpperCase() === 'CONDITIONAL' ? 'conditional-field' : '' ?>" id="group-<?= processedFieldsData[i][0] ?>">
                    <? if (processedFieldsData[i][1].toUpperCase() === "STATICTEXT" && processedFieldsData[i][2][0]) { ?>
                      <div class="static-text"><?!= processedFieldsData[i][2][0] ?></div>
                    <? } else if (processedFieldsData[i][1].toUpperCase() === "TABLE" && processedFieldsData[i][2].length > 0) { ?>
  <label class="<?= processedFieldsData[i][4] ? 'required' : '' ?>"><?= processedFieldsData[i][0] ?>:</label>
  <table class="table-display">
    <? var tableData = processedFieldsData[i][2]; ?>
    <? for (var row = 0; row < tableData.length; row++) { ?>
      <tr>
        <? var isHeader = row === 0; ?>
        <? for (var col = 0; col < tableData[row].length; col++) { ?>
          <? var cellValue = String(tableData[row][col] || '').trim(); ?>
          <? if (isHeader) { ?>
            <th><?!= cellValue ? cellValue.replace(/</g, '<').replace(/>/g, '>') : '' ?></th>
          <? } else { ?>
            <td>
  <? if (cellValue.match(/\.(jpg|jpeg|png|gif)$/i) || cellValue.includes("drive.google.com")) { ?>
    <? 
      var imgSrc = cellValue;
      var fallbackSrc = cellValue;
      if (cellValue.includes("drive.google.com")) {
        var idMatch = cellValue.match(/([a-zA-Z0-9_-]{20,})/);
        if (idMatch) {
          imgSrc = "https://drive.google.com/thumbnail?id=" + idMatch[1];
          fallbackSrc = "https://drive.google.com/uc?id=" + idMatch[1];
        }
      }
    ?>
    <img src="<?= imgSrc ?>" style="width: 100px; height: auto;" alt="Table Image" 
         onerror="this.src='<?= fallbackSrc ?>'; if(this.complete && this.naturalHeight === 0) { this.style.display='none'; document.getElementById('<?= processedFieldsData[i][0] ?>-error').textContent='Image failed to load: <?= cellValue.replace(/'/g, "\\'") ?>'; }">
  <? } else if (cellValue.match(/(youtube\.com|youtu\.be)/i)) { ?>
    <? 
      var videoId;
      if (cellValue.includes("youtu.be")) {
        videoId = cellValue.split('/').pop().split('?')[0];
      } else {
        var match = cellValue.match(/[?&]v=([^&]+)/);
        videoId = match ? match[1] : cellValue.split('/').pop().split('?')[0];
      }
    ?>
    <iframe src="https://www.youtube.com/embed/<?= videoId ?>" frameborder="0" allowfullscreen></iframe>
  <? } else if (cellValue.match(/^(https?:\\/\\/|mailto:|tel:).+$/i)) { ?>
    <? 
      var displayText = cellValue;
      if (cellValue.length > 40) {
        displayText = cellValue.substring(0, 37) + '...';
      }
    ?>
    <a href="<?= cellValue ?>" target="_blank" class="hyperlink">
      <?= displayText ?>
    </a>
  <? } else { ?>
    <?!= cellValue ? cellValue.replace(/</g, '<').replace(/>/g, '>') : '' ?>
    <!-- Debug: cellValue='<?= cellValue.replace(/'/g, "\\'") ?>', isURL=<?= cellValue.match(/^(https?:\\/\\/|mailto:|tel:).+$/i) ? 'true' : 'false' ?> -->
  <? } ?>
</td>
          <? } ?>
        <? } ?>
      </tr>
    <? } ?>
  </table>
  <span class="error" id="<?= processedFieldsData[i][0] ?>-error"></span>
                    <? } else if (processedFieldsData[i][1].toUpperCase() === "HYPERLINK" && processedFieldsData[i][2][0]) { ?>
                      <label class="<?= processedFieldsData[i][4] ? 'required' : '' ?>"><?= processedFieldsData[i][0] ?>:</label>
                      <? if (processedFieldsData[i][2][0].match(/<a\s[^>]*href=["'][^"']*["'][^>]*>/i)) { ?>
                        <?!= processedFieldsData[i][2][0] ?>
                        <input type="hidden" name="<?= processedFieldsData[i][0] ?>" value="<?= processedFieldsData[i][2][1] ?>">
                      <? } else { ?>
                        <a href="<?= processedFieldsData[i][2][0] ?>" class="hyperlink" target="_blank"><?= processedFieldsData[i][2][0] ?></a>
                        <input type="hidden" name="<?= processedFieldsData[i][0] ?>" value="<?= processedFieldsData[i][2][0] ?>">
                      <? } ?>
                      <span class="error" id="<?= processedFieldsData[i][0] ?>-error"></span>
                    <? } else { ?>
                      <label for="<?= processedFieldsData[i][0] ?>" class="<?= processedFieldsData[i][4] ? 'required' : '' ?>"><?= processedFieldsData[i][0] ?>:</label>
                      <? if (processedFieldsData[i][1].toUpperCase() === "DROPDOWN" && processedFieldsData[i][2].length > 0) { ?>
                        <select id="<?= processedFieldsData[i][0] ?>" name="<?= processedFieldsData[i][0] ?>" <?= processedFieldsData[i][4] ? 'required' : '' ?>>
                          <? var options = processedFieldsData[i][2]; ?>
                          <? for (var j = 0; j < options.length; j++) { ?>
                            <option value="<?= options[j] ?>"><?= options[j] ?></option>
                          <? } ?>
                        </select>
                      <? } else if (processedFieldsData[i][1].toUpperCase() === "MULTISELECT" && processedFieldsData[i][2].length > 0) { ?>
                        <select id="<?= processedFieldsData[i][0] ?>" name="<?= processedFieldsData[i][0] ?>" multiple <?= processedFieldsData[i][4] ? 'required' : '' ?>>
                          <? var options = processedFieldsData[i][2]; ?>
                          <? for (var j = 0; j < options.length; j++) { ?>
                            <option value="<?= options[j] ?>"><?= options[j] ?></option>
                          <? } ?>
                        </select>
                      <? } else if (processedFieldsData[i][1].toUpperCase() === "DATE") { ?>
                        <input type="date" id="<?= processedFieldsData[i][0] ?>" name="<?= processedFieldsData[i][0] ?>" <?= processedFieldsData[i][4] ? 'required' : '' ?>>
                      <? } else if (processedFieldsData[i][1].toUpperCase() === "TIME") { ?>
                        <input type="time" id="<?= processedFieldsData[i][0] ?>" name="<?= processedFieldsData[i][0] ?>" <?= processedFieldsData[i][4] ? 'required' : '' ?>>
                      <? } else if (processedFieldsData[i][1].toUpperCase() === "NUMBER") { ?>
                        <input type="number" id="<?= processedFieldsData[i][0] ?>" name="<?= processedFieldsData[i][0] ?>" <?= processedFieldsData[i][4] ? 'required' : '' ?>>
                      <? } else if (processedFieldsData[i][1].toUpperCase() === "CHECKBOX") { ?>
                        <input type="checkbox" id="<?= processedFieldsData[i][0] ?>" name="<?= processedFieldsData[i][0] ?>" <?= processedFieldsData[i][4] ? 'required' : '' ?>>
                      <? } else if (processedFieldsData[i][1].toUpperCase() === "RADIO" && processedFieldsData[i][2].length > 0) { ?>
                        <div class="radio-group" id="<?= processedFieldsData[i][0] ?>">
                          <? var options = processedFieldsData[i][2]; ?>
                          <? for (var j = 0; j < options.length; j++) { ?>
                            <div>
                              <input type="radio" id="<?= processedFieldsData[i][0] + '-' + j ?>" name="<?= processedFieldsData[i][0] ?>" value="<?= options[j] ?>" <?= processedFieldsData[i][4] && j === 0 ? 'required' : '' ?>>
                              <label for="<?= processedFieldsData[i][0] + '-' + j ?>"><?= options[j] ?></label>
                            </div>
                          <? } ?>
                        </div>
                      <? } else if (processedFieldsData[i][1].toUpperCase() === "TEXTAREA") { ?>
                        <textarea id="<?= processedFieldsData[i][0] ?>" name="<?= processedFieldsData[i][0] ?>" <?= processedFieldsData[i][4] ? 'required' : '' ?>></textarea>
                      <? } else if (processedFieldsData[i][1].toUpperCase() === "EMAIL") { ?>
                        <input type="email" id="<?= processedFieldsData[i][0] ?>" name="<?= processedFieldsData[i][0] ?>" <?= processedFieldsData[i][4] ? 'required' : '' ?>>
                      <? } else if (processedFieldsData[i][1].toUpperCase() === "STARRATING") { ?>
                        <div class="star-rating" id="<?= processedFieldsData[i][0] ?>">
                          <input type="radio" id="<?= processedFieldsData[i][0] ?>-5" name="<?= processedFieldsData[i][0] ?>" value="5" <?= processedFieldsData[i][4] ? 'required' : '' ?>>
                          <label for="<?= processedFieldsData[i][0] ?>-5">★</label>
                          <input type="radio" id="<?= processedFieldsData[i][0] ?>-4" name="<?= processedFieldsData[i][0] ?>" value="4">
                          <label for="<?= processedFieldsData[i][0] ?>-4">★</label>
                          <input type="radio" id="<?= processedFieldsData[i][0] ?>-3" name="<?= processedFieldsData[i][0] ?>" value="3">
                          <label for="<?= processedFieldsData[i][0] ?>-3">★</label>
                          <input type="radio" id="<?= processedFieldsData[i][0] ?>-2" name="<?= processedFieldsData[i][0] ?>" value="2">
                          <label for="<?= processedFieldsData[i][0] ?>-2">★</label>
                          <input type="radio" id="<?= processedFieldsData[i][0] ?>-1" name="<?= processedFieldsData[i][0] ?>" value="1">
                          <label for="<?= processedFieldsData[i][0] ?>-1">★</label>
                        </div>
                      <? } else if (processedFieldsData[i][1].toUpperCase() === "RANGESLIDER" && processedFieldsData[i][2].length === 3) { ?>
                        <input type="range" id="<?= processedFieldsData[i][0] ?>" name="<?= processedFieldsData[i][0] ?>" 
                          min="<?= processedFieldsData[i][2][0] ?>" max="<?= processedFieldsData[i][2][1] ?>" step="<?= processedFieldsData[i][2][2] ?>" <?= processedFieldsData[i][4] ? 'required' : '' ?>>
                        <span class="range-output" id="<?= processedFieldsData[i][0] ?>-output"><?= processedFieldsData[i][2][0] ?></span>
                      <? } else if (processedFieldsData[i][1].toUpperCase() === "FILEUPLOAD") { ?>
                        <input type="file" id="<?= processedFieldsData[i][0] ?>" name="<?= processedFieldsData[i][0] ?>" accept="image/*,.pdf" <?= processedFieldsData[i][4] ? 'required' : '' ?>>
                      <? } else if (processedFieldsData[i][1].toUpperCase() === "CONDITIONAL" && processedFieldsData[i][2][0]) { ?>
                        <input type="text" id="<?= processedFieldsData[i][0] ?>" name="<?= processedFieldsData[i][0] ?>" data-condition="<?= processedFieldsData[i][2][0] ?>" <?= processedFieldsData[i][4] ? 'required' : '' ?>>
                      <? } else if (processedFieldsData[i][1].toUpperCase() === "CALCULATED" && processedFieldsData[i][2][0]) { ?>
                        <input type="text" id="<?= processedFieldsData[i][0] ?>" name="<?= processedFieldsData[i][0] ?>" class="calculated-field" readonly>
                      <? } else if (processedFieldsData[i][1].toUpperCase() === "SIGNATURE") { ?>
                        <canvas id="<?= processedFieldsData[i][0] ?>" name="<?= processedFieldsData[i][0] ?>"></canvas>
                        <input type="hidden" id="<?= processedFieldsData[i][0] ?>-hidden" name="<?= processedFieldsData[i][0] ?>" <?= processedFieldsData[i][4] ? 'required' : '' ?>>
                      <? } else if (processedFieldsData[i][1].toUpperCase() === "GEOLOCATION") { ?>
                        <input type="text" id="<?= processedFieldsData[i][0] ?>" name="<?= processedFieldsData[i][0] ?>" readonly <?= processedFieldsData[i][4] ? 'required' : '' ?>>
                        <button type="button" onclick="getLocation('<?= processedFieldsData[i][0] ?>')">Get Location</button>
                      <? } else if (processedFieldsData[i][1].toUpperCase() === "PROGRESSBAR" && processedFieldsData[i][2].length > 0) { ?>
                        <progress id="<?= processedFieldsData[i][0] ?>" value="<?= String(processedFieldsData[i][2][0] || '0').startsWith('=') ? 0 : processedFieldsData[i][2][0] || 0 ?>" max="100"></progress>
                      <? } else if (processedFieldsData[i][1].toUpperCase() === "CAPTCHA") { ?>
                        <input type="text" id="<?= processedFieldsData[i][0] ?>" name="<?= processedFieldsData[i][0] ?>" placeholder="Enter sum (e.g., 3 + 5)" <?= processedFieldsData[i][4] ? 'required' : '' ?>>
                        <span id="captcha-question">What is 3 + 5?</span>
                      <? } else if (processedFieldsData[i][1].toUpperCase() === "IMAGE" && processedFieldsData[i][2][0]) { ?>
                        <img src="<?= processedFieldsData[i][2][0] ?>" alt="<?= processedFieldsData[i][0] ?>" id="<?= processedFieldsData[i][0] ?>" 
                             onerror="this.style.display='none'; document.getElementById('<?= processedFieldsData[i][0] ?>-error').textContent='Image failed to load: <?= processedFieldsData[i][2][0].replace(/'/g, "\\'") ?>';">
                        <input type="hidden" name="<?= processedFieldsData[i][0] ?>" value="<?= processedFieldsData[i][2][0] ?>">
                      <? } else if (processedFieldsData[i][1].toUpperCase() === "VIDEO" && processedFieldsData[i][2][0]) { ?>
  <? 
    var videoUrl = processedFieldsData[i][2][0];
    var videoId = '';
    if (videoUrl.includes("youtu.be") || videoUrl.includes("youtube.com")) {
      if (videoUrl.includes("youtu.be")) {
        videoId = videoUrl.split('/').pop().split('?')[0];
      } else if (videoUrl.includes("youtube.com")) {
        var match = videoUrl.match(/[?&]v=([^&]+)/);
        if (match) {
          videoId = match[1];
        } else if (videoUrl.includes("/embed/")) {
          videoId = videoUrl.split('/embed/')[1].split('?')[0];
        }
      }
    }
  ?>
  <? if (videoId) { ?>
    <iframe width="250" height="150" src="https://www.youtube.com/embed/<?= videoId ?>" frameborder="0" allowfullscreen></iframe>
  <? } else if (!videoUrl.includes("youtu")) { ?>
    <video controls id="<?= processedFieldsData[i][0] ?>">
      <source src="<?= videoUrl ?>" type="video/mp4">
      Your browser does not support the video tag.
    </video>
  <? } else { ?>
    <span class="error" id="<?= processedFieldsData[i][0] ?>-error">Invalid YouTube URL</span>
  <? } ?>
  <input type="hidden" name="<?= processedFieldsData[i][0] ?>" value="<?= videoUrl ?>">
                      <? } else if (processedFieldsData[i][1].toUpperCase() === "IMAGELINK") { ?\>
                        <input type="text" id="<?= processedFieldsData[i][0] ?>" name="<?= processedFieldsData[i][0] ?>" placeholder="Enter Image URL" <?= processedFieldsData[i][4] ? 'required' : '' ?> oninput="previewImage(this)">
                        <img id="<?= processedFieldsData[i][0] ?>-preview" style="display: none;" alt="Preview">
                        <span class="error" id="<?= processedFieldsData[i][0] ?>-error"></span>
                      <? } else if (processedFieldsData[i][1].toUpperCase() === "VIDEOLINK") { ?>
                        <input type="text" id="<?= processedFieldsData[i][0] ?>" name="<?= processedFieldsData[i][0] ?>" placeholder="Enter Video URL" <?= processedFieldsData[i][4] ? 'required' : '' ?>>
                      <? } else { ?>
                        <input type="text" id="<?= processedFieldsData[i][0] ?>" name="<?= processedFieldsData[i][0] ?>" <?= processedFieldsData[i][4] ? 'required' : '' ?>>
                      <? } ?>
                      <span class="error" id="<?= processedFieldsData[i][0] ?>-error"></span>
                    <? } ?>
                  </div>
                <? } ?>
              <? } ?>
              <? if (inContainer) { ?></div><? } ?>
              <button type="submit" id="submitButton">Submit <span class="spinner" id="spinner"></span></button>
            </form>
            <div id="message">Data submitted successfully!</div>
          <? } else { ?>
            <div class="no-fields">No fields defined. Please add fields in FormSetup A10:J.</div>
          <? } ?>
        </div>
        <script>
          <? if (processedFieldsData.length > 0) { ?>
            const processedFieldsData = <?!= JSON.stringify(processedFieldsData) ?>;
            const taxRate = <?!= taxRate ?>;
            let signatureCanvases = {};

            function escapeHtml(text) {
              const div = document.createElement('div');
              div.textContent = text;
              return div.innerHTML;
            }

            function addCheckoutItem(fieldId) {
              const tbody = document.getElementById(fieldId + "-tbody");
              const rowCount = tbody.getElementsByTagName("tr").length;
              const rowId = fieldId + "-row-" + rowCount;
              const field = processedFieldsData.find(f => f[0] === fieldId);

              const row = document.createElement("tr");
              row.id = rowId;

              const selectTd = document.createElement("td");
              const select = document.createElement("select");
              select.name = "description";
              select.onchange = () => updateCheckoutTotals(fieldId);
              const defaultOption = document.createElement("option");
              defaultOption.value = "";
              defaultOption.text = "Select an item";
              select.appendChild(defaultOption);

              field[2].forEach(item => {
                const option = document.createElement("option");
                option.value = JSON.stringify({ description: item.description, unitPrice: item.unitPrice });
                option.text = escapeHtml(item.description);
                select.appendChild(option);
              });
              selectTd.appendChild(select);

              const quantityTd = document.createElement("td");
              const quantityInput = document.createElement("input");
              quantityInput.type = "number";
              quantityInput.name = "quantity";
              quantityInput.min = "0";
              quantityInput.value = "0";
              quantityInput.oninput = () => updateCheckoutTotals(fieldId);
              quantityTd.appendChild(quantityInput);

              const unitPriceTd = document.createElement("td");
              unitPriceTd.className = "unitPrice";
              unitPriceTd.textContent = "$0.00";

              const itemTotalTd = document.createElement("td");
              itemTotalTd.className = "itemTotal";
              itemTotalTd.textContent = "$0.00";

              const actionTd = document.createElement("td");
              const removeButton = document.createElement("button");
              removeButton.type = "button";
              removeButton.className = "remove-item-btn";
              removeButton.textContent = "Remove";
              removeButton.onclick = () => removeCheckoutItem(fieldId, rowId);
              actionTd.appendChild(removeButton);

              row.appendChild(selectTd);
              row.appendChild(quantityTd);
              row.appendChild(unitPriceTd);
              row.appendChild(itemTotalTd);
              row.appendChild(actionTd);

              tbody.appendChild(row);
              updateCheckoutTotals(fieldId);
            }

            function removeCheckoutItem(fieldId, rowId) {
              const row = document.getElementById(rowId);
              if (row) row.parentNode.removeChild(row);
              updateCheckoutTotals(fieldId);
            }

            function updateCheckoutTotals(fieldId) {
              const tbody = document.getElementById(fieldId + "-tbody");
              if (!tbody) return;
              const rows = tbody.querySelectorAll("tr");
              let subtotal = 0;

              rows.forEach(row => {
                const select = row.querySelector("select[name='description']");
                const quantityInput = row.querySelector("input[name='quantity']");
                const unitPriceCell = row.querySelector(".unitPrice");
                const itemTotalCell = row.querySelector(".itemTotal");

                let unitPrice = 0;
                let description = "";
                const quantity = parseFloat(quantityInput.value) || 0;

                if (select.value) {
                  try {
                    const item = JSON.parse(select.value);
                    description = item.description;
                    unitPrice = parseFloat(item.unitPrice) || 0;
                    unitPriceCell.textContent = "$" + unitPrice.toFixed(2); // Fixed: Update unit price immediately
                  } catch (e) {
                    console.error("Error parsing item:", e);
                    unitPriceCell.textContent = "$0.00";
                  }
                } else {
                  unitPriceCell.textContent = "$0.00";
                }

                const total = quantity * unitPrice;
                subtotal += total;
                itemTotalCell.textContent = "$" + total.toFixed(2);
              });

              const items = Array.from(rows).map(row => {
                const select = row.querySelector("select[name='description']");
                const quantityInput = row.querySelector("input[name='quantity']");
                const value = select.value;
                const quantity = parseFloat(quantityInput.value) || 0;
                let description = "";
                let unitPrice = 0;

                if (value) {
                  try {
                    const item = JSON.parse(value);
                    description = item.description;
                    unitPrice = parseFloat(item.unitPrice) || 0;
                  } catch (e) {
                    console.error("Error parsing item:", e);
                  }
                }

                return { description, quantity, unitPrice };
              }).filter(item => item.quantity > 0 && item.description);

              const tax = subtotal * taxRate;
              const total = subtotal + tax;

              document.getElementById(fieldId + "-subtotal").textContent = "$" + subtotal.toFixed(2);
              document.getElementById(fieldId + "-tax").textContent = "$" + tax.toFixed(2);
              document.getElementById(fieldId + "-total").textContent = "$" + total.toFixed(2);
              document.getElementById(fieldId).value = JSON.stringify(items);
            }

            function handleSubmit(event) {
              event.preventDefault();
              const form = document.getElementById('myForm');
              const submitButton = document.getElementById('submitButton');
              const spinner = document.getElementById('spinner');
              const dataToSend = {};
              let isValid = true;
              let pendingFiles = 0;

              const inputs = form.querySelectorAll('input, select, textarea');
              inputs.forEach(input => {
                const name = input.name;
                if (!name || input.type === 'button') return;

                let value;
                const errorSpan = document.getElementById(name + '-error');
                const fieldData = processedFieldsData.find(f => f[0] === name);
                const isRequired = fieldData && fieldData[4];

                if (input.type === 'file' && input.files.length > 0) {
                  const file = input.files[0];
                  if (file.size > 6 * 1024 * 1024) {
                    errorSpan.textContent = 'File too large (max 6 MB)';
                    isValid = false;
                  } else {
                    pendingFiles++;
                    const reader = new FileReader();
                    reader.onload = function(e) {
                      dataToSend[name] = {
                        name: file.name,
                        data: e.target.result.split(',')[1],
                        type: file.type || 'application/octet-stream'
                      };
                      pendingFiles--;
                      if (pendingFiles === 0 && isValid) submitForm();
                    };
                    reader.readAsDataURL(file);
                  }
                } else if (input.type === 'checkbox') {
                  value = input.checked;
                  dataToSend[name] = value;
                } else if (input.type === 'radio') {
                  if (input.checked) dataToSend[name] = input.value;
                  return;
                } else if (input.tagName === 'SELECT' && input.multiple) {
                  value = Array.from(input.selectedOptions).map(option => option.value).join(',');
                  dataToSend[name] = value;
                } else if (input.id.endsWith('-hidden') && signatureCanvases[name]) {
                  value = signatureCanvases[name].toDataURL().split(',')[1];
                  dataToSend[name] = { name: name + '.png', data: value, type: 'image/png' };
                } else if (fieldData && fieldData[1].toUpperCase() === 'CHECKOUT') {
                  value = input.value;
                  if (isRequired && (!value || value === '[]')) {
                    errorSpan.textContent = 'Please add at least one item with a quantity greater than 0';
                    isValid = false;
                  } else {
                    errorSpan.textContent = '';
                    dataToSend[name] = value;
                  }
                } else {
                  value = input.value;
                  dataToSend[name] = value;
                }

                if (fieldData && fieldData[1].toUpperCase() !== 'CHECKOUT' && isRequired && (!value || value === '')) {
                  errorSpan.textContent = 'This field is required';
                  isValid = false;
                } else if (input.type === 'number' && value && isNaN(value)) {
                  errorSpan.textContent = 'Please enter a valid number';
                  isValid = false;
                } else if (input.type === 'email' && value && !/^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/.test(value)) {

                  errorSpan.textContent = 'Please enter a valid email';
                  isValid = false;
                } else if (input.id.startsWith('CAPTCHA') && value !== '8') {
                  errorSpan.textContent = 'Incorrect answer. Please enter 8.';
                  isValid = false;
                } else if (fieldData && fieldData[1].toUpperCase() !== 'CHECKOUT') {
                  errorSpan.textContent = '';
                }
              });

              processedFieldsData.forEach(field => {
                if (field[1].toUpperCase() === "CALCULATED" && field[2][0]) {
                  const calcField = document.getElementById(field[0]);
                  const formula = field[2][0].split('=')[1];
                  const parts = formula.match(/(\w+|\d+|[*+/-])/g);
                  let result = 0;
                  if (parts) {
                    result = evaluateFormula(parts, dataToSend);
                    calcField.value = result;
                    dataToSend[field[0]] = result;
                  }
                }
              });

              if (pendingFiles === 0 && isValid) submitForm();

              function submitForm() {
                submitButton.disabled = true;
                submitButton.textContent = 'Submitting...';
                spinner.style.display = 'inline-block';

                google.script.run
                  .withSuccessHandler(() => {
                    form.reset();
                    resetSignatures();
                    processedFieldsData.forEach(field => {
                      if (field[1].toUpperCase() === "CHECKOUT") {
                        const tbody = document.getElementById(field[0] + "-tbody");
                        tbody.innerHTML = "";
                        addCheckoutItem(field[0]);
                        updateCheckoutTotals(field[0]);
                      }
                    });
                    showMessage();
                    submitButton.disabled = false;
                    submitButton.textContent = 'Submit';
                    spinner.style.display = 'none';
                  })
                  .withFailureHandler(error => {
                    alert('Error submitting form: ' + error.message);
                    submitButton.disabled = false;
                    submitButton.textContent = 'Submit';
                    spinner.style.display = 'none';
                  })
                  .processForm(dataToSend);
              }
            }

            function evaluateFormula(parts, data) {
              let result = 0;
              let operator = '+';
              parts.forEach(part => {
                if (['+', '-', '*', '/'].includes(part)) {
                  operator = part;
                } else {
                  const num = isNaN(part) ? (data[part] || 0) : Number(part);
                  if (operator === '+') result += num;
                  else if (operator === '-') result -= num;
                  else if (operator === '*') result *= num;
                  else if (operator === '/' && num !== 0) result /= num;
                }
              });
              return result;
            }

            function showMessage() {
              const message = document.getElementById('message');
              message.style.display = 'block';
              setTimeout(() => message.style.display = 'none', 3000);
            }

            function getLocation(fieldId) {
              if (navigator.geolocation) {
                navigator.geolocation.getCurrentPosition(
                  position => {
                    document.getElementById(fieldId).value = position.coords.latitude + ',' + position.coords.longitude;
                  },
                  error => document.getElementById(fieldId + '-error').textContent = 'Unable to get location'
                );
              } else {
                document.getElementById(fieldId + '-error').textContent = 'Geolocation not supported';
              }
            }

            function previewImage(input) {
              const preview = document.getElementById(input.id + '-preview');
              if (input.value) {
                preview.src = input.value;
                preview.style.display = 'block';
                preview.onerror = () => {
                  preview.style.display = 'none';
                  document.getElementById(input.id + '-error').textContent = 'Invalid image URL';
                };
              } else {
                preview.style.display = 'none';
                document.getElementById(input.id + '-error').textContent = '';
              }
            }

            processedFieldsData.forEach(field => {
              if (field[1].toUpperCase() === "RANGESLIDER") {
                const slider = document.getElementById(field[0]);
                const output = document.getElementById(field[0] + '-output');
                if (slider && output) {
                  slider.oninput = () => output.textContent = slider.value;
                }
              } else if (field[1].toUpperCase() === "SIGNATURE") {
                const canvas = document.getElementById(field[0]);
                if (canvas) {
                  const ctx = canvas.getContext('2d');
                  let drawing = false;
                  signatureCanvases[field[0]] = canvas;

                  canvas.onmousedown = e => {
                    drawing = true;
                    ctx.beginPath();
                    ctx.moveTo(e.offsetX, e.offsetY);
                  };
                  canvas.onmousemove = e => {
                    if (drawing) {
                      ctx.lineTo(e.offsetX, e.offsetY);
                      ctx.stroke();
                    }
                  };
                  canvas.onmouseup = () => drawing = false;
                  canvas.onmouseleave = () => drawing = false;
                }
              } else if (field[1].toUpperCase() === "CONDITIONAL" && field[2][0]) {
                const [triggerField, triggerValue] = field[2][0].split('=');
                const triggerInput = document.getElementById(triggerField);
                const conditionalGroup = document.getElementById('group-' + field[0]);
                if (triggerInput && conditionalGroup) {
                  triggerInput.onchange = () => {
                    const show = (triggerInput.type === 'checkbox' ? triggerInput.checked : triggerInput.value) === triggerValue;
                    conditionalGroup.style.display = show ? 'flex' : 'none';
                  };
                }
              } else if (field[1].toUpperCase() === "CHECKOUT") {
                updateCheckoutTotals(field[0]);
              }
            });

            function resetSignatures() {
              Object.values(signatureCanvases).forEach(canvas => {
                const ctx = canvas.getContext('2d');
                ctx.clearRect(0, 0, canvas.width, canvas.height);
              });
            }
          <? } ?>
        </script>
      </body>
    </html>
  `);

  template.formName = formName;
  template.processedFieldsData = processedFieldsData;
  template.taxRate = taxRate;
  template.additionalStyles = additionalStyles;
  return template.evaluate().setTitle(formName || "Form Preview");
}

function doGet(e) {
  return generateFormHTML();
}

function loadFormRows() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var setupSheet = ss.getSheetByName("FormSetup");
  if (!setupSheet) {
    createFormSetupSheet();
    setupSheet = ss.getSheetByName("FormSetup");
  }

  var lastRow = setupSheet.getLastRow();
  if (lastRow < 9) return [];

  var range = setupSheet.getRange("A9:J" + lastRow);
  var values = range.getValues();

  return values.map(row => ({
    fieldName: row[0],
    sheet1: row[1],
    cell1: row[2],
    sheet2: row[3],
    cell2: row[4],
    sheet3: row[5],
    cell3: row[6],
    type: row[7],
    options: row[8],
    required: row[9]
  }));
}

function saveFormRows(rows) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var setupSheet = ss.getSheetByName("FormSetup");
  if (!setupSheet) {
    createFormSetupSheet();
    setupSheet = ss.getSheetByName("FormSetup");
  }

  // Get the last row with data in column A starting from row 10
  var lastRow = setupSheet.getLastRow();
  var startRow = 10; // Default start row
  if (lastRow >= 10) {
    var fieldNames = setupSheet.getRange("A10:A" + lastRow).getValues();
    // Find the last row with non-empty data in column A
    for (var i = fieldNames.length - 1; i >= 0; i--) {
      if (fieldNames[i][0] !== "") {
        startRow = 10 + i + 1; // Next row after the last non-empty row
        break;
      }
    }
  }

  // Write new data starting at startRow
  if (rows.length > 0) {
    var data = rows.map(row => [
      row.fieldName || "",
      row.sheet1 || "",
      row.cell1 || "",
      row.sheet2 || "",
      row.cell2 || "",
      row.sheet3 || "",
      row.cell3 || "",
      row.type || "",
      row.options || "",
      row.required || ""
    ]);
    setupSheet.getRange("A" + startRow + ":J" + (startRow + data.length - 1)).setValues(data);
  }
}

function saveToResponses(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var setupSheet = ss.getSheetByName("FormSetup");
  if (!setupSheet) throw new Error("FormSetup sheet not found.");

  Logger.log("Saving to Responses: " + JSON.stringify(data));

  var fieldsRange = setupSheet.getRange("A10:J" + setupSheet.getLastRow());
  var fieldsData = fieldsRange.getValues().filter(row => row[0] !== "");

  fieldsData.forEach(row => {
    var fieldName = row[0];
    var fieldType = row[7] || "Text";
    var targetSheetName = row[1];
    var targetColumn = row[2];
    var fieldValue = data[fieldName];

    if (fieldValue && fieldType.toUpperCase() === "CHECKOUT" && targetSheetName && targetColumn) {
      try {
        var items = JSON.parse(fieldValue);
        if (items.length === 0) return;

        var checkoutData = items.map(item => [item.description, item.quantity]);
        var targetSheet = getOrCreateSheet(ss, targetSheetName);
        var startColumn = columnToNumber(targetColumn);
        var lastRow = targetSheet.getLastRow();
        var nextRow = lastRow >= 1 ? lastRow + 1 : 1;
        targetSheet.getRange(nextRow, startColumn, checkoutData.length, 2).setValues(checkoutData);
        
        Logger.log(`Data appended to ${targetSheetName} at ${targetColumn}${nextRow}:${String.fromCharCode(64 + startColumn + 1)}${nextRow + checkoutData.length - 1}: ` + JSON.stringify(checkoutData));
      } catch (e) {
        Logger.log(`Error processing Checkout field ${fieldName}: ${e.message}`);
      }
    }
  });
}

function columnToNumber(column) {
  return column.toUpperCase().charCodeAt(0) - 64;
}

// Helper function to get or create a folder in Google Drive
function getOrCreateFolder(folderName) {
  var folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next(); // Return the first matching folder
  }
  return DriveApp.createFolder(folderName); // Create new folder if none exists
}

function processForm(formData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var setupSheet = ss.getSheetByName("FormSetup");
  if (!setupSheet) throw new Error("FormSetup sheet not found.");

  var fieldsRange = setupSheet.getRange("A9:J" + setupSheet.getLastRow());
  var fieldsData = fieldsRange.getValues().filter(row => row[0] !== "");

  // Validate required fields
  fieldsData.forEach(row => {
    var fieldName = row[0];
    var isRequired = row[9].toString().toLowerCase() === "yes";
    var fieldValue = formData[fieldName];

    if (isRequired && (fieldValue === undefined || fieldValue === "" || fieldValue === null)) {
      throw new Error(`Field "${fieldName}" is required but was not provided.`);
    }
  });

  var sheetsData = {};
  fieldsData.forEach(row => {
    var fieldName = row[0];
    var targetSheets = [row[1], row[3], row[5]].filter(Boolean);
    var targetCells = [row[2], row[4], row[6]].filter(Boolean);
    var fieldValue = formData[fieldName];

    if (fieldValue === undefined) return;

    // Handle file uploads and checkout fields
    if (typeof fieldValue === 'object' && fieldValue.data) {
      fieldValue = uploadFile(fieldValue);
    } else if (row[7].toUpperCase() === "CHECKOUT" && typeof fieldValue === 'string') {
      try {
        var checkoutItems = JSON.parse(fieldValue);
        if (checkoutItems.length === 0) {
          fieldValue = "";
        } else {
          fieldValue = checkoutItems
            .map(item => `${item.description}:${item.quantity}:${item.unitPrice}`)
            .join(",");
        }
      } catch (e) {
        fieldValue = "";
      }
    }

    targetSheets.forEach((sheetName, index) => {
      if (!sheetName || !targetCells[index]) return;
      if (!sheetsData[sheetName]) sheetsData[sheetName] = { singleCell: [], tableRow: [] };

      var targetCell = targetCells[index];
      if (/^[A-Z]+[0-9]+$/.test(targetCell)) {
        sheetsData[sheetName].singleCell.push({ fieldName, targetCell, value: fieldValue });
      } else if (/^[A-Z]+$/.test(targetCell) && row[7].toUpperCase() === "CHECKOUT" && fieldValue) {
        var checkoutItems = fieldValue.split(",").map(item => {
          var [description, quantity, price] = item.split(":");
          return [description, parseInt(quantity), parseFloat(price)];
        });
        sheetsData[sheetName].tableRow.push({ fieldName, column: targetCell, value: checkoutItems });
      } else if (/^[A-Z]+$/.test(targetCell)) {
        sheetsData[sheetName].tableRow.push({ fieldName, column: targetCell, value: fieldValue });
      }
    });
  });

  // Write data to target sheets
  Object.keys(sheetsData).forEach(sheetName => {
    var sheet = getOrCreateSheet(ss, sheetName);
    var singleCellData = sheetsData[sheetName].singleCell;
    singleCellData.forEach(data => {
      sheet.getRange(data.targetCell).setValue(data.value);
    });

    var tableRowData = sheetsData[sheetName].tableRow;
    if (tableRowData.length > 0) {
      // Find the last row with data across all relevant columns
      var lastRow = Math.max(1, sheet.getLastRow());
      var nextRow = lastRow >= 1 ? lastRow + 1 : 2;

      tableRowData.forEach(data => {
        if (Array.isArray(data.value) && data.fieldName && fieldsData.find(row => row[0] === data.fieldName && row[7].toUpperCase() === "CHECKOUT")) {
          data.value.forEach(item => {
            var colIndex = data.column.charCodeAt(0) - 65;
            var rowData = new Array(Math.max(colIndex + 3, 1)).fill('');
            rowData[colIndex] = item[0]; // description
            rowData[colIndex + 1] = item[1]; // quantity
            rowData[colIndex + 2] = item[2]; // price
            sheet.getRange(nextRow, 1, 1, rowData.length).setValues([rowData]);
            nextRow++;
          });
        } else {
          var colIndex = columnToNumber(data.column);
          var targetRange = sheet.getRange(nextRow, colIndex, 1, 1);
          targetRange.setValue(data.value);
        }
      });
    }
  });

  // Email notification
  var emailRecipient = setupSheet.getRange("B8").getValue();
  if (emailRecipient) {
    var subject = "New Form Submission";
    var body = "A new form submission has been received:\n\n";
    fieldsData.forEach(row => {
      var fieldName = row[0];
      var fieldValue = formData[fieldName] || "(No value)";
      if (row[7].toUpperCase() === "CHECKOUT" && fieldValue) {
        try {
          fieldValue = JSON.parse(fieldValue)
            .filter(item => item.quantity > 0)
            .map(item => `${item.description} (Qty: ${item.quantity}, Price: ${item.unitPrice})`)
            .join("\n") || "(No items selected)";
        } catch (e) {
          fieldValue = "(Invalid checkout data)";
        }
      }
      body += `${fieldName}: ${fieldValue}\n`;
    });
    try {
      MailApp.sendEmail(emailRecipient, subject, body);
    } catch (e) {
      Logger.log(`Error sending email: ${e.message}`);
    }
  }

  // Execute on-submit functions
  var onSubmitFunctions = setupSheet.getRange("B6").getValue();
  if (onSubmitFunctions) {
    var functionNames = onSubmitFunctions.split(',').map(name => name.trim());
    var functionMap = {
      "save": save,
      "copyInput1": copyInput1,
      "newcontact": newcontact,
      "updateInventory": updateInventory
    };

    functionNames.forEach(funcName => {
      if (functionMap[funcName]) {
        try {
          functionMap[funcName]();
        } catch (e) {
          Logger.log(`Error executing function ${funcName}: ${e.message}`);
        }
      } else {
        try {
          var func = new Function(`return ${funcName}`)();
          if (typeof func === "function") func();
          else Logger.log(`Function ${funcName} is not callable`);
        } catch (e) {
          Logger.log(`Function ${funcName} not found or invalid: ${e.message}`);
        }
      }
    });
  }

  return "Success";
}

function uploadFile(fileData) {
  // Initialize folders
  var fileUploadsFolder = getOrCreateFolder("File Uploads_Signatures");
 

  // Extract file details from fileData
  var fileName = fileData.fileName || "Untitled";
  var mimeType = fileData.mimeType || "application/octet-stream";
  var data = fileData.data;

  // Convert base64 data to Blob (assuming data is base64-encoded)
  var bytes = Utilities.base64Decode(data);
  var blob = Utilities.newBlob(bytes, mimeType, fileName);

  // Determine if the file is a signature (e.g., based on MIME type)
  var isSignature = mimeType.startsWith("image/"); // Adjust logic as needed

  // Choose the target folder
  var targetFolder = isSignature ? signaturesFolder : fileUploadsFolder;

  // Save the file to the target folder
  var file = targetFolder.createFile(blob);

  // Return the file URL or ID for storage in the spreadsheet
  return file.getUrl();
}

function getOrCreateSheet(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}


function save() {
  Logger.log("save() function triggered.");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName("Data");
  const viewPrintSheet = ss.getSheetByName("View_Print");
  const updateSheet = ss.getSheetByName("Update");
  const logSheet = ss.getSheetByName("Log");

  if (!dataSheet || !viewPrintSheet || !updateSheet || !logSheet) {
    Logger.log('Missing one or more required sheets.');
    return;
  }

  let inputSheet = ss.getSheetByName("Input");
  let originalName = "";
  let renamed = false;

  try {
    if (!inputSheet) {
      const activeSheet = ss.getActiveSheet();
      originalName = activeSheet.getName();

      if (!["Data", "Log", "View_Print", "Update"].includes(originalName)) {
        activeSheet.setName("Input");
        renamed = true;
        inputSheet = activeSheet;
      } else {
        Logger.log('No sheet named "Input", and active sheet is not suitable for renaming.');
        return;
      }
    }
 dataSheet.insertRowAfter(1);

  inputSheet
    .getRange("A1:Q1")
    .copyTo(dataSheet.getRange("B2"), { contentsOnly: true });
  inputSheet
    .getRange("A2:Q2")
    .copyTo(dataSheet.getRange("S2"), { contentsOnly: true });
  inputSheet
    .getRange("A3:Q3")
    .copyTo(dataSheet.getRange("AJ2"), { contentsOnly: true });
  inputSheet
    .getRange("A4:Q4")
    .copyTo(dataSheet.getRange("BA2"), { contentsOnly: true });
  inputSheet
    .getRange("A5:Q5")
    .copyTo(dataSheet.getRange("BR2"), { contentsOnly: true });
  inputSheet
    .getRange("A6:Q6")
    .copyTo(dataSheet.getRange("CI2"), { contentsOnly: true });
  inputSheet
    .getRange("A7:Q7")
    .copyTo(dataSheet.getRange("CZ2"), { contentsOnly: true });
  inputSheet
    .getRange("A8:Q8")
    .copyTo(dataSheet.getRange("DQ2"), { contentsOnly: true });
  inputSheet
    .getRange("A9:Q9")
    .copyTo(dataSheet.getRange("EH2"), { contentsOnly: true });
  inputSheet
    .getRange("A10:Q10")
    .copyTo(dataSheet.getRange("EY2"), { contentsOnly: true });
  inputSheet
    .getRange("A11:Q11")
    .copyTo(dataSheet.getRange("FP2"), { contentsOnly: true });
  inputSheet
    .getRange("A12:Q12")
    .copyTo(dataSheet.getRange("GG2"), { contentsOnly: true });
  inputSheet
    .getRange("A13:Q13")
    .copyTo(dataSheet.getRange("GX2"), { contentsOnly: true });
  inputSheet
    .getRange("A14:Q14")
    .copyTo(dataSheet.getRange("HO2"), { contentsOnly: true });
  inputSheet
    .getRange("A15:Q15")
    .copyTo(dataSheet.getRange("IF2"), { contentsOnly: true });
  inputSheet
    .getRange("A16:Q16")
    .copyTo(dataSheet.getRange("IW2"), { contentsOnly: true });
  inputSheet
    .getRange("A17:Q17")
    .copyTo(dataSheet.getRange("JN2"), { contentsOnly: true });
  inputSheet
    .getRange("A18:Q18")
    .copyTo(dataSheet.getRange("KE2"), { contentsOnly: true });
  inputSheet
    .getRange("A19:Q19")
    .copyTo(dataSheet.getRange("KV2"), { contentsOnly: true });
  inputSheet
    .getRange("A20:Q20")
    .copyTo(dataSheet.getRange("LM2"), { contentsOnly: true });
  inputSheet
    .getRange("A21:Q21")
    .copyTo(dataSheet.getRange("MD2"), { contentsOnly: true });
  inputSheet
    .getRange("A22:Q22")
    .copyTo(dataSheet.getRange("MU2"), { contentsOnly: true });
  inputSheet
    .getRange("A23:Q23")
    .copyTo(dataSheet.getRange("NL2"), { contentsOnly: true });
  inputSheet
    .getRange("A24:Q24")
    .copyTo(dataSheet.getRange("OC2"), { contentsOnly: true });
  inputSheet
    .getRange("A25:Q25")
    .copyTo(dataSheet.getRange("OT2"), { contentsOnly: true });
  inputSheet
    .getRange("A26:Q26")
    .copyTo(dataSheet.getRange("PK2"), { contentsOnly: true });
  inputSheet
    .getRange("A27:Q27")
    .copyTo(dataSheet.getRange("QB2"), { contentsOnly: true });
  inputSheet
    .getRange("A28:Q28")
    .copyTo(dataSheet.getRange("QS2"), { contentsOnly: true });
  inputSheet
    .getRange("A29:Q29")
    .copyTo(dataSheet.getRange("RJ2"), { contentsOnly: true });
  inputSheet
    .getRange("A30:Q30")
    .copyTo(dataSheet.getRange("SA2"), { contentsOnly: true });
  inputSheet
    .getRange("A31:Q31")
    .copyTo(dataSheet.getRange("SR2"), { contentsOnly: true });
  inputSheet
    .getRange("A32:Q32")
    .copyTo(dataSheet.getRange("TI2"), { contentsOnly: true });
  inputSheet
    .getRange("A33:Q33")
    .copyTo(dataSheet.getRange("TZ2"), { contentsOnly: true });
  inputSheet
    .getRange("A34:Q34")
    .copyTo(dataSheet.getRange("UQ2"), { contentsOnly: true });
  inputSheet
    .getRange("A35:Q35")
    .copyTo(dataSheet.getRange("VH2"), { contentsOnly: true });
  inputSheet
    .getRange("A36:Q36")
    .copyTo(dataSheet.getRange("VY2"), { contentsOnly: true });
  inputSheet
    .getRange("A37:Q37")
    .copyTo(dataSheet.getRange("WP2"), { contentsOnly: true });
  inputSheet
    .getRange("A38:Q38")
    .copyTo(dataSheet.getRange("XG2"), { contentsOnly: true });
  inputSheet
    .getRange("A39:Q39")
    .copyTo(dataSheet.getRange("XX2"), { contentsOnly: true });
  inputSheet
    .getRange("A40:Q40")
    .copyTo(dataSheet.getRange("YO2"), { contentsOnly: true });
  inputSheet
    .getRange("A41:Q41")
    .copyTo(dataSheet.getRange("ZF2"), { contentsOnly: true });
  inputSheet
    .getRange("A42:Q42")
    .copyTo(dataSheet.getRange("ZW2"), { contentsOnly: true });
  inputSheet
    .getRange("A43:Q43")
    .copyTo(dataSheet.getRange("AAN2"), { contentsOnly: true });
  inputSheet
    .getRange("A44:Q44")
    .copyTo(dataSheet.getRange("ABE2"), { contentsOnly: true });
  inputSheet
    .getRange("A45:Q45")
    .copyTo(dataSheet.getRange("ABV2"), { contentsOnly: true });
  inputSheet
    .getRange("A46:Q46")
    .copyTo(dataSheet.getRange("ACM2"), { contentsOnly: true });
  inputSheet
    .getRange("A47:Q47")
    .copyTo(dataSheet.getRange("ADD2"), { contentsOnly: true });
  inputSheet
    .getRange("A48:Q48")
    .copyTo(dataSheet.getRange("ADU2"), { contentsOnly: true });

  dataSheet
    .getRange("A2")
    .setFormula(
      '=Data!S2&"- "&T2&"- "&U2&"- "&V2&"- "&W2&"- "&X2&"- "&Y2&"- "&Z2&"- "&AA2&"- "&AB2&"- "&AC2&"- "&AD2'
    );
  dataSheet
    .getRange("A1")
    .setFormula(
      '=Data!S1&"- "&T1&"- "&U1&"- "&V1&"- "&W1&"- "&X1&"- "&Y1&"- "&Z1&"- "&AA1&"- "&AB1&"- "&AC1&"- "&AD1'
    );

  dataSheet
    .getRange("AE2")
    .setFormula("=VLOOKUP(A2,Update!$A$1:$CF$1000000,2,FALSE)");
  dataSheet
    .getRange("AF2")
    .setFormula("=VLOOKUP(A2,Update!$A$1:$CF$1000000,3,FALSE)");
  dataSheet
    .getRange("AG2")
    .setFormula("=VLOOKUP(A2,Update!$A$1:$CF$1000000,4,FALSE)");

      const images = inputSheet.getImages();
  images.forEach(image => {
    const sourceRange = image.getAnchorCell();
    // Check if the image's anchor cell is within A3:Q48
    if (sourceRange.getRow() >= 3 && sourceRange.getRow() <= 48 && sourceRange.getColumn() >= 1 && sourceRange.getColumn() <= 17) {
      const blob = image.getBlob();
      const targetRow = sourceRange.getRow() - 2; // Adjust row to fit new data structure
      const targetColumn = sourceRange.getColumn();
      
      // Insert image into Data sheet at the corresponding position
      dataSheet.insertImage(blob, targetColumn, targetRow + 1); // +1 for row offset due to inserted row
    }
  });

  updateSheet.insertRowAfter(1);
  updateSheet.getRange("A2").setFormula("=Data!A2");
  updateSheet.getRange("E2").setFormula("=Data!S2");
  updateSheet.getRange("F2").setFormula("=Data!T2");
  updateSheet.getRange("G2").setFormula("=Data!U2");
  updateSheet.getRange("H2").setFormula("=Data!V2");
  updateSheet.getRange("I2").setFormula("=Data!W2");
  updateSheet.getRange("J2").setFormula("=Data!X2");
  updateSheet.getRange("K2").setFormula("=Data!Y2");
  updateSheet.getRange("L2").setFormula("=Data!Z2");
  updateSheet.getRange("M2").setFormula("=Data!AA2");
  updateSheet.getRange("N2").setFormula("=Data!AB2");
  updateSheet.getRange("O2").setFormula("=Data!AC2");
  updateSheet.getRange("P2").setFormula("=Data!AD2");

  var rangeWithFilter = logSheet.getRange("A10:O10");
  var filterCriteria = rangeWithFilter.getFilter().getRange().getA1Notation();

  logSheet.insertRowBefore(10);

  rangeWithFilter.getFilter().remove();

  var fullRange = logSheet.getRange("A9:O9" + logSheet.getLastRow());
  fullRange.createFilter();

  logSheet.getRange("A10").setFormula("=Data!S2");
  logSheet.getRange("B10").setFormula("=Data!T2");
  logSheet.getRange("C10").setFormula("=Data!U2");
  logSheet.getRange("D10").setFormula("=Data!V2");
  logSheet.getRange("E10").setFormula("=Data!W2");
  logSheet.getRange("F10").setFormula("=Data!X2");
  logSheet.getRange("G10").setFormula("=Data!Y2");
  logSheet.getRange("H10").setFormula("=Data!Z2");
  logSheet.getRange("I10").setFormula("=Data!AA2");
  logSheet.getRange("J10").setFormula("=Data!AB2");
  logSheet.getRange("K10").setFormula("=Data!AC2");
  logSheet.getRange("L10").setFormula("=Data!AD2");
  logSheet.getRange("M10").setFormula("=Data!AE2");
  logSheet.getRange("N10").setFormula("=Data!AF2");
  logSheet.getRange("O10").setFormula("=Data!AG2");

  logSheet.getRange("A9").setFormula("=Input!A1");
  logSheet.getRange("B9").setFormula("=Input!B1");
  logSheet.getRange("C9").setFormula("=Input!C1");
  logSheet.getRange("D9").setFormula("=Input!D1");
  logSheet.getRange("E9").setFormula("=Input!E1");
  logSheet.getRange("F9").setFormula("=Input!F1");
  logSheet.getRange("G9").setFormula("=Input!G1");
  logSheet.getRange("H9").setFormula("=Input!H1");
  logSheet.getRange("I9").setFormula("=Input!I1");
  logSheet.getRange("J9").setFormula("=Input!J1");
  logSheet.getRange("K9").setFormula("=Input!K1");
  logSheet.getRange("L9").setFormula("=Input!L1");
  logSheet.getRange("M9").setFormula("=Input!M1");
  logSheet.getRange("N9").setFormula("=Input!N1");
  logSheet.getRange("O9").setFormula("=Input!O1");

  viewPrintSheet.getRange("B2").activate();
      // Clear any existing content or hyperlink in cells B1:L1
  viewPrintSheet.getRange("B1:L1").clearContent();

  // Merge cells B1:L1
  viewPrintSheet.getRange("B1:L1").merge();

  // Add the hyperlink with the display text
  viewPrintSheet.getRange("B1").setFormula('=HYPERLINK("https://datamateapp.github.io/Donate%205%20per%20mo.html", "Support DataMate")');

  // Set the font style and color for the hyperlink
  const cell = viewPrintSheet.getRange("B1");
  cell.setFontWeight("bold");
  cell.setFontSize(12);
  cell.setFontColor("#0066CC"); // Set font color to a noticeable blue

  // Set fill to No Fill for merged cells B1:L1
  cell.setBackground(null); // Removes any background color

  // Center-align the text in the merged cells
  cell.setHorizontalAlignment("center");
  cell.setVerticalAlignment("middle");


    
    sendDataMateEmail1();

   } catch (e) {
    Logger.log("Error in save(): " + e.message);
  } finally {
    if (renamed && originalName) {
      ss.getSheetByName("Input").setName(originalName);
    }
  }
}

function sendDataMateEmail1() {
  MailApp.sendEmail({
    to: "projectprodigyapp@gmail.com",
    subject: "DataMate Record",
    body: "Another record saved."
  });
}





function newcontact() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const newContact = ss.getSheetByName("NewContact");
  const contactsSheet = ss.getSheetByName("contacts");

   contactsSheet.insertRowAfter(1);

newContact.getRange('B1').copyTo(contactsSheet.getRange('contacts!B2'), { contentsOnly: true });
newContact.getRange('B2').copyTo(contactsSheet.getRange('contacts!C2'), { contentsOnly: true });
newContact.getRange('B3').copyTo(contactsSheet.getRange('contacts!D2'), { contentsOnly: true });
newContact.getRange('B4').copyTo(contactsSheet.getRange('contacts!AR2'), { contentsOnly: true });
newContact.getRange('B5').copyTo(contactsSheet.getRange('contacts!AZ2'), { contentsOnly: true });
newContact.getRange('B6').copyTo(contactsSheet.getRange('contacts!BD2'), { contentsOnly: true });
newContact.getRange('B7').copyTo(contactsSheet.getRange('contacts!BE2'), { contentsOnly: true });
newContact.getRange('B8').copyTo(contactsSheet.getRange('contacts!BF2'), { contentsOnly: true });
newContact.getRange('B9').copyTo(contactsSheet.getRange('contacts!P2'), { contentsOnly: true });
newContact.getRange('B10').copyTo(contactsSheet.getRange('contacts!AN2'), { contentsOnly: true });
newContact.getRange('B11').copyTo(contactsSheet.getRange('contacts!AP2'), { contentsOnly: true });
newContact.getRange('B12').copyTo(contactsSheet.getRange('contacts!Z2'), { contentsOnly: true });
newContact.getRange('B13').copyTo(contactsSheet.getRange('contacts!AD2'), { contentsOnly: true });
newContact.getRange('B14').copyTo(contactsSheet.getRange('contacts!AE2'), { contentsOnly: true });
newContact.getRange('B15').copyTo(contactsSheet.getRange('contacts!AF2'), { contentsOnly: true });
newContact.getRange('B16').copyTo(contactsSheet.getRange('contacts!BK2'), { contentsOnly: true });
newContact.getRange('B17').copyTo(contactsSheet.getRange('contacts!BO2'), { contentsOnly: true });
newContact.getRange('B18').copyTo(contactsSheet.getRange('contacts!BP2'), { contentsOnly: true });
newContact.getRange('B19').copyTo(contactsSheet.getRange('contacts!BQ2'), { contentsOnly: true });
newContact.getRange('B20').copyTo(contactsSheet.getRange('contacts!E2'), { contentsOnly: true });
newContact.getRange('B21').copyTo(contactsSheet.getRange('contacts!T2'), { contentsOnly: true });
newContact.getRange('B22').copyTo(contactsSheet.getRange('contacts!V2'), { contentsOnly: true });

contactsSheet.getRange('A2').activate();
contactsSheet.getCurrentCell().setFormula('=CONCATENATE(B2," ",C2," ",D2)');
contactsSheet.getRange('A:A').activate();
contactsSheet.getRange('A1').copyTo(contactsSheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);


newContact.getRange('B1:B22').activate();
newContact.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  
}

function copyInput1() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName("Sheet1");
  var targetSheet = ss.getSheetByName("Input");
  
  // Define the range to copy
  var copyRange = sourceSheet.getRange("A3:Q48");
  var targetRange = targetSheet.getRange("A3:Q48");
  
  // Copy everything from source to target
  copyRange.copyTo(targetRange); // This will copy values, formats, and formulas
  
  // Copy column widths
  var sourceColWidths = [];
  var lastColumnSource = sourceSheet.getLastColumn();
  var lastColumnTarget = targetSheet.getLastColumn();
  
  // Ensure we only consider columns up to the last column in both sheets
  var columnsToCopy = Math.min(lastColumnSource, lastColumnTarget, 17); // A to Q = 17 columns
  
  for (var i = 1; i <= columnsToCopy; i++) {
    sourceColWidths.push(sourceSheet.getColumnWidth(i));
  }
  
  // Set column widths in target sheet, but only for existing columns
  for (var j = 1; j <= columnsToCopy; j++) {
    targetSheet.setColumnWidth(j, sourceColWidths[j - 1]);
  }
  
  // Select cell C4 in the target sheet
  targetSheet.activate();
  targetSheet.getRange("C4").activate();

}

