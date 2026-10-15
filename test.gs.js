function testFolderAccess() {
  var folderId = "1ISUyp1CDBKgfIAzV8BbcAXrA5F8skZZq"; // 対象のフォルダID
  var folder = DriveApp.getFolderById(folderId);
  Logger.log("フォルダ名: " + folder.getName());
}