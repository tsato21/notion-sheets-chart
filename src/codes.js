//Defines the variable, "spreadsheet" that will be used in the following functions
const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

/**
 * Lists name and url of each sheet in the target spreadsheet and set borders.
 * Data will be input from the A2 cell.
 */
function display_sheet_info() {
  const spreadsheet_url = spreadsheet.getUrl();
  const main_sheet = spreadsheet.getSheetByName(main_sheet_name);
  const sheets = spreadsheet.getSheets();

  const sheet_info = [];
  for(sheet of sheets){
    let sheet_url = spreadsheet_url + "#gid=" + sheet.getSheetId();
    let sheet_name = sheet.getSheetName();
    sheet_info.push([sheet_name,sheet_url]);
  }

  const range = main_sheet.getRange(2,1,sheet_info.length,2);
  range.setValues(sheet_info);
  range.setBorder(true,true,true,true,true,true);
}

/**
 * Displays info of files stored in a designated folders and set borders.
 * Data will be input from the A2 cell.
 */

function display_file_info(){
  const sheet = spreadsheet.getSheetByName(display_sheet);
  //Checks if there is previous data on the sheet. If so, that data is cleared.
  let last_row_pre = sheet.getRange('A:A').getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  console.log(last_row_pre);
  if(last_row_pre >= 2 && last_row_pre !== 1000){
    sheet.getRange(2,1,last_row_pre-1,4).clearContent();
    console.log('Previous data was cleared.');
  } else {
    console.log('There is no previous data on the sheet.');
  }
  
  
  let folder_id = Browser.inputBox("Enter folder ID",Browser.Buttons.OK_CANCEL);
  if(folder_id == "cancel"){
    Browser.msgBox("Cancelled");
    return;
  } else if (folder_id == ""){
    Browser.msgBox("No folder ID entered");
    return;
  }
  let folder = DriveApp.getFolderById(folder_id);
  let files = folder.getFiles();
  let file_info = [];
  while(files.hasNext()){
    let file = files.next();
    let file_name = file.getName();
    let file_id = file.getId();
    let file_url = file.getUrl();
    let editors = file.getEditors().map((editor) => editor.getEmail()).join(",");
    if (editors === ''){
      editors = 'no editors';
    }
    file_info.push([file_name,file_id,file_url,editors]);
  }
  console.log(file_info);
  const range = sheet.getRange(2,1,file_info.length,4);
  range.setValues(file_info);
  range.setBorder(true,true,true,true,true,true);
  
}

/**
 * Adds designated accounts (more than one) as editor to designated files
 * addEditor() function can be used when it comes to adding only one account as editor
 * This function is for adding more than one account as editor by using Drive API
 * By sendNotificationEmails: false, the designated accounts will not receive notification emails
 * Must enable Drive API in the target project in advance.
 * Must display necessary data for each target file: file name, id, an email for the target student and faculty member in advance.
 */
function add_editors_each_file() {
  let data_check = Browser.msgBox("Make sure that necessary data is input in the A to D columns in 'add-editors' sheet.",Browser.Buttons.YES_NO_CANCEL);
  if (data_check == "cancel"){
    Browser.msgBox("Cancelled");
    return;
  } else if (data_check == "no"){
    Browser.msgBox("Input necessary data and try again.");
    return;
  }
  
  // Gets values of target ranges on the sheet
  const sheet = spreadsheet.getSheetByName(add_editor_sheet);
  // Gets preparatory data from the sheet
  const pre_data = sheet.getDataRange().getValues();
  const pre_header = pre_data.shift();
  // Gets target data from the sheet
  const data = pre_data.map((row) => pre_header.reduce((o, k, i) => {
    o[k] = row[i];
    return o;
  }, {}));

  console.log(`data:\n${JSON.stringify(data)}`); // Checks if the data is correctly retrieved

  //Checks how many cells in E column on the sheet has "Cancelled" as a value. If there is any, data in E2 or after is cleared.
  let last_row_pre = sheet.getRange('E:E').getNextDataCell(SpreadsheetApp.Direction.DOWN).getRow();
  console.log(last_row_pre);
  if(last_row_pre >= 2 && last_row_pre !== 1000){
    sheet.getRange(2,5,999,1).clearContent();
    console.log('Previous data was cleared.');
  } else {
    console.log('There is no previous data on the sheet.');
  }
  
  //Checks whether Drive API is enabled in this project; otherwise, the following commands will not be executed.
  let drive_api_check;
  try {
    Drive.Files.list();
    drive_api_check = true;
  } catch (error) {
    drive_api_check = false;
  }
  if (drive_api_check) {
    Logger.log('Drive API is enabled in the project.');
  } else {
    Logger.log('Drive API is not enabled in the project.');
    Browser.msgBox('Drive API is not enabled in the project. Enable it in the Apps Script and try again.');
    return;
  }

  // Adds designated accounts (student/ faculty member) as an editor to a Spreadsheet that the student and faculty member works on together
  for (let i = 0; i < data.length; i++) {
    const each_data = data[i];
    const each_file = DriveApp.getFileById(each_data['ID']);
    //Checks whether the type of the stored file is shortcut version. If so, the following execution for the file is cancelled and the execution for the next file starts.
    if (each_file.getMimeType() === 'application/vnd.google-apps.shortcut') {
      console.log(`Skipping item with ID ${each_data['ID']} due to mime-type 'application/vnd.google-apps.shortcut'`);
      sheet.getRange(`E${i+2}`).setValue('Cancelled');
      continue;
    }
    const add_student_editor = Drive.Permissions.insert(
      {
        value: each_data['SHARE-TO (STUDENT)'],
        type: 'user',
        role: 'writer'
      },
      each_data['ID'],
      {
        sendNotificationEmails: false
      }
    );
    console.log(`add_student_editor: ${JSON.stringify(add_student_editor)}`); // Return values from API execution. Details of the permissions actually set via the API are returned.
    const add_faculty_editor = Drive.Permissions.insert(
      {
        value: each_data['SHARE-TO (FACULTY)'],
        type: 'user',
        role: 'writer'
      },
      each_data['ID'],
      {
        sendNotificationEmails: false
      }
    );
    console.log(`add_faculty_editor: ${JSON.stringify(add_faculty_editor)}`); // Return values from API execution. Details of the permissions actually set via the API are returned.
  }
  Browser.msgBox('Designated emails were added as editor to the designated file, excluding shortcut version.');
  
}

/**
 * Resets all authorization of files of a designated folder.
 */
function reset_all() {
  let folder_id = Browser.inputBox("Enter folder ID",Browser.Buttons.OK_CANCEL);
  if(folder_id == "cancel"){
    Browser.msgBox("Cancelled");
    return;
  } else if (folder_id == ""){
    Browser.msgBox("No folder ID entered");
    return;
  }
  const files = DriveApp.getFolderById(folder_id).getFiles();
  while (files.hasNext()) {
    // Gets editor info of each file in that folder and remove those editors
    const file = files.next();
    const file_name = file.getName();
    const editors = file.getEditors();
    editors.forEach((editor) => {
      const editor_email = editor.getEmail();
      file.removeEditor(editor_email);
      console.info(`Editor ${editor_email} removed for ${file_name}`);
    })
  }
  Browser.msgBox("Resetting completed");
}

/**
 * Adds designated emails as editor to files stored in a designated folder sending without notification email.
 */
function add_editors_files() {
  const folder_id = spreadsheet.getSheetByName(add_editor_sheet2).getRange('B1').getValue();
  const files = DriveApp.getFolderById(folder_id).getFiles();
  const editors = spreadsheet.getSheetByName(add_editor_sheet2).getRange('B2:B').getValues();
  const editors_list = editors.map((editor) => editor[0]);
  
  while (files.hasNext()) {
    const file = files.next();
    for (let i = 0; i < editors_list.length; i++){
      let editor = editors_list[i];
      if(editor){
        Drive.Permissions.insert(
          {
            value: editor,
            type: 'user',
            role: 'writer'
          },
          file.getId(),
          {
            'sendNotificationEmails': false
          }
        );
      }
    }
  }
  // Browser.msgBox('Added designated emails as editors to stored files.');
  SpreadsheetApp.getUi().alert('Added designated emails as editors to stored files');
}
