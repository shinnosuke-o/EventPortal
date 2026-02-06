function testSetDestination() {
  const DEST_SS_ID = '1BL08rl8USs2cbIgjyD_9EDP9xL_ilRKwF1pStHv7El4';
  const form = FormApp.create('destination-test');
  form.setDestination(FormApp.DestinationType.SPREADSHEET, DEST_SS_ID);
  Logger.log(form.getEditUrl());
}

function authDrive() {
  DriveApp.getFolderById('1w7QeWmi0rWRSkVf9fsNTe6_dTlNm9VYa').getName();
  DriveApp.getFileById(form.getId()).moveTo(folder);

}

function authAll() {
  // Drive API（高度なGoogleサービス）側の呼び出しでスコープ承認を発火
  Drive.Files.get('root');
  // Form / Spreadsheet も触っておく（念のため）
  FormApp.getActiveForm(); // なくてもOK（失敗しても承認画面は出ることが多い）
}
// テスト