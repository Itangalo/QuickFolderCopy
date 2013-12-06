/**
 * Adds the relevant menu to the spreadsheet.
 */
function onOpen() {
  var menuEntries = [];
  menuEntries.push({name : 'Skapa grundmappar', functionName : 'createBaseFolders'});
  menuEntries.push({name : 'Skapa elevmappar', functionName : 'createStudentFolders'});
  menuEntries.push({name : 'Dela ut en fil', functionName : 'shareFile'});
  SpreadsheetApp.getActiveSpreadsheet().addMenu('QuickFolderCopy', menuEntries);
};

/**
 * Displays a notification to the user.
 */
function message(message, hold) {
  if (hold == true) {
    Browser.msgBox(message);
  }
  else {
    SpreadsheetApp.getActiveSpreadsheet().toast(message, '');
  }
}

/**
 * Creates all the basic folders for QuickFolderCopy.
 */
function createBaseFolders() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Elevmappar');
  
  var className = sheet.getRange('B1').getValue();
  if (className == '') {
    message('Du måste ange ett klassnamn först. Använd de beiga rutorna för att fylla i information.', true);
    return;
  }
  
  // Create the base folder for the class.
  var classFolderID = sheet.getRange('B2').getValue();
  try {
    var classFolder = DocsList.getFolderById(classFolderID);
  }
  catch(e) {
    var classFolder = createFolder(className);
    sheet.getRange('B2').setFormula('=hyperlink("' + classFolder.getUrl() + '";"' + classFolder.getId() + '")');
  }
  
  // Create the teacher-only folder for the class.
  var teacherFolderID = sheet.getRange('B3').getValue();
  try {
    var teacherFolder = DocsList.getFolderById(teacherFolderID);
  }
  catch(e) {
    var teacherFolder = createFolder(className + ' (lärarmapp)', classFolder);
    sheet.getRange('B3').setFormula('=hyperlink("' + teacherFolder.getUrl() + '";"' + teacherFolder.getId() + '")');
  }
  
  // Create the folder viewable by the whole class.
  var allViewFolderID = sheet.getRange('B4').getValue();
  try {
    var allViewFolder = DocsList.getFolderById(allViewFolderID);
  }
  catch(e) {
    var allViewFolder = createFolder(className + ' (alla kan se)', classFolder);
    sheet.getRange('B4').setFormula('=hyperlink("' + allViewFolder.getUrl() + '";"' + allViewFolder.getId() + '")');
    shareWithStudents(allViewFolder, 'view');
  }

  // Create the folder editable by the whole class.
  var allEditFolderID = sheet.getRange('B5').getValue();
  try {
    var allEditFolder = DocsList.getFolderById(allEditFolderID);
  }
  catch(e) {
    var allEditFolder = createFolder(className + ' (alla kan redigera)', classFolder);
    sheet.getRange('B5').setFormula('=hyperlink("' + allEditFolder.getUrl() + '";"' + allEditFolder.getId() + '")');
    shareWithStudents(allEditFolder, 'edit');
  }

  // Create the base folder for all student folders.
  var studentBaseFolderID = sheet.getRange('B6').getValue();
  try {
    var studentBaseFolder = DocsList.getFolderById(studentBaseFolderID);
  }
  catch(e) {
    var studentBaseFolder = createFolder(className + ' (elevmappar)', classFolder);
    sheet.getRange('B6').setFormula('=hyperlink("' + studentBaseFolder.getUrl() + '";"' + studentBaseFolder.getId() + '")');
  }
}

/**
 * Creates folders for all students.
 */
function createStudentFolders() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Elevmappar');
  
  var className = sheet.getRange('B1').getValue();
  if (className == '') {
    message('Du måste ange ett klassnamn först. Använd de beiga rutorna för att fylla i information.', true);
    return;
  }
  var studentBaseFolderID = sheet.getRange('B6').getValue();
  try {
    var studentBaseFolder = DocsList.getFolderById(studentBaseFolderID);
  }
  catch(e) {
    message('Det finns ingen giltig basmapp att lägga elevmappar i. Kör skriptet för att skapa grundmappar först.', true);
    return;
  }

  if (getStudentMails() == []) {
    message('Det finns inga e-postadresser för eleverna. Använd de beiga rutorna för att fylla i information.', true);
    return;
  }

  // Process all student folders.
  var studentData = sheet.getRange(9, 1, sheet.getLastRow() - 8, 5).getValues();
  for (var row in studentData) {
    var actualRow = parseInt(row) + 9;
    if (studentData[row][0] == '' || studentData[row][1] == '') {
      message('Elev på rad ' + actualRow + ' saknar namn eller e-postadress och hoppas över. Kör skriptet igen när det är fixat.');
    }
    else {
      // Create the student container folder.
      var containerFolderID = studentData[row][2];
      try {
        var containerFolder = DocsList.getFolderById(containerFolderID);
      }
      catch(e) {
        var containerFolder = createFolder(studentData[row][0], studentBaseFolder);
        sheet.getRange(actualRow, 3).setFormula('=hyperlink("' + containerFolder.getUrl() + '";"' + containerFolder.getId() + '")');
      }

      // Create the student view folder.
      var viewFolderID = studentData[row][3];
      try {
        var viewFolder = DocsList.getFolderById(viewFolderID);
      }
      catch(e) {
        var viewFolder = createFolder(studentData[row][0] + ' (' + className + ', endast se)', containerFolder);
        sheet.getRange(actualRow, 4).setFormula('=hyperlink("' + viewFolder.getUrl() + '";"' + viewFolder.getId() + '")');
        shareWithStudent(viewFolder, 'view', studentData[row][1]);
      }

      // Create the student edit folder.
      var editFolderID = studentData[row][4];
      try {
        var editFolder = DocsList.getFolderById(editFolderID);
      }
      catch(e) {
        var editFolder = createFolder(studentData[row][0] + ' (' + className + ', redigering)', containerFolder);
        sheet.getRange(actualRow, 5).setFormula('=hyperlink("' + editFolder.getUrl() + '";"' + editFolder.getId() + '")');
        shareWithStudent(editFolder, 'edit', studentData[row][1]);
      }
    }
  }

}

/**
 * Creates a folder, shares with the teachers, and moves it into any parentFolder.
 */
function createFolder(name, parentFolder) {
  var folder = DocsList.createFolder(name);
  if (typeof parentFolder != 'undefined') {
    folder.addToFolder(parentFolder);
    folder.removeFromFolder(DocsList.getRootFolder());
  }

  var teacherEmails = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Elevmappar').getRange('B7').getValue();
  if (teacherEmails != '') {
    folder.addEditors(teacherEmails.split(','));
  }

  return folder;
}

/**
 * Shares a file or folder with all students, either for 'view' or 'edit' mode.
 */
function shareWithStudents(item, mode) {
  if (mode == 'edit') {
    try {
      item.addEditors(getStudentMails());
      message('Delade ' + item.getName() + ' med eleverna, så att de kan redigera.');
    }
    catch(e) {
      message('Kunde inte dela ' + item.getName() + ' med eleverna. Förmodligen finns det e-postadresser som inte stämmer.', true);
    }
  }
  else {
    try {
      item.addViewers(getStudentMails());
      message('Delade ' + item.getName() + ' med eleverna, så att de kan se.');
    }
    catch(e) {
      message('Kunde inte dela ' + item.getName + ' med eleverna. Förmodligen finns det e-postadresser som inte stämmer.', true);
    }
  }
}

/**
 * Shares a file or folder with a single student, either for 'view' or 'edit' mode.
 */
function shareWithStudent(item, mode, mail) {
  if (mode == 'edit') {
    try {
      item.addEditor(mail);
    }
    catch(e) {
      message('Kunde inte dela ' + item.getName() + ' med ' + mail + ' för redigering. Är e-postadressen korrekt, och kopplad till ett Google-konto?', true);
    }
  }
  else {
    try {
      item.addViewer(mail);
    }
    catch(e) {
      message('Kunde inte dela ' + item.getName() + ' med ' + mail + ' för visning. Är e-postadressen korrekt, och kopplad till ett Google-konto?', true);
    }
  }
}

/**
 * Builds an array with all student emails.
 */
function getStudentMails() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Elevmappar');
  var studentMailsRaw = sheet.getRange(9, 2, sheet.getLastRow() - 8).getValues();
  var studentMails = '';
  for (var row in studentMailsRaw) {
    if (studentMailsRaw[row][0] != '') {
      if (studentMails != '') {
        studentMails += ',' + studentMailsRaw[row][0];
      }
      else {
        studentMails += studentMailsRaw[row][0];
      }
    }
  }
  if (studentMails == '') {
    return [];
  }
  else {
    return studentMails.split(',');
  }
}

/**
 * Allows copying a single file to all students.
 */
function shareFile() {
  var app = UiApp.createApplication();
  var handler = app.createServerHandler('shareFileModeSelector');
  app.createDocsListDialog().setDialogTitle('Välj fil att kopiera till eleverna').showDocsPicker().addSelectionHandler(handler);

  SpreadsheetApp.getActiveSpreadsheet().show(app);
}

/**
 * Popup for selecting how a file should be distributed to the students.
 */
function shareFileModeSelector(eventInfo) {
  var app = UiApp.getActiveApplication();
  
  var handler = app.createServerHandler('shareFileDistribute');
  
  var file = app.createHidden('file', eventInfo.parameter.items[0].id);
  app.add(file);
  handler.addCallbackElement(file);

  app.add(app.createLabel('Välj hur filen ska kopieras'));

  var mode = app.createListBox().setName('mode');
  mode.addItem('Endast visa', 'view');
  mode.addItem('Eleven får redigera', 'edit');
  app.add(mode);
  handler.addCallbackElement(mode);
  
  app.add(app.createLabel(''));
  
  var ok = app.createButton('Kopiera fil till elever', handler);
  app.add(ok);
  
  return app;
}

/**
 * Handler for actually copying and distributing a file.
 */
function shareFileDistribute(eventInfo) {
  message('Påbörjar kopiering...');
  var app = UiApp.getActiveApplication();
  app.close();
  
  if (eventInfo.parameter.mode == 'edit') {
    var column = 4;
  }
  else {
    var column = 3;
  }
  var sourceFile = DocsList.getFileById(eventInfo.parameter.file);

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Elevmappar');
  var studentData = sheet.getRange(9, 1, sheet.getLastRow() - 8, 5).getValues();
  for (var row in studentData) {
    var copy = sourceFile.makeCopy(sourceFile.getName() + ' (' + studentData[row][0] + ')');
    var targetFolder = DocsList.getFolderById(studentData[row][column]);
    copy.addToFolder(targetFolder);
    copy.removeFromFolder(DocsList.getRootFolder());
  }
  
  message('Kopiering klar. Filerna finns nu i respektive elevmapp.');

  return app;
}

function debug(variable, mode) {
  if (mode == 'index') {
    var output = '';
    for (var i in variable) {
      output += i + ': ' + variable[i];
    }
    Browser.msgBox(output);
  }
  else {
    message(variable, typeof variable);
  }
}
