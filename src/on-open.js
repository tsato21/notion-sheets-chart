function onOpen() {
    let ui = SpreadsheetApp.getUi();
    ui.createMenu('Function Execution')
    .addItem('Display sheet info of this spreadsheet', 'display_sheet_info')
    .addSeparator()
    .addItem('Display file info in the designated folder', 'display_file_info')
    .addSeparator()
    .addItem('Add designated emails as editors to a designated file without notification', 'add_editors_each_file')
    .addSeparator()
    .addItem('Add designated emails as editors to files stored in a designated folder without notification','add_editors_files')
    .addSeparator()
    .addItem('Reset authorization of files in a designated folder','reset_all')
    .addToUi();
}
