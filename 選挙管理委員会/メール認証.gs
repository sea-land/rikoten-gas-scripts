function main() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('投票結果');
  if (!sheet) {
    Logger.log('シートが見つかりません: 投票結果');
    return;
  }

  const firstRow = 2; // データの開始行(=今回は2行目)
  const mailCol = 3;  // メールアドレスの列(=今回はC列)
  const lastRow = sheet.getLastRow();
  
  if (lastRow < firstRow) {
    Logger.log('データがありません。');
    return;
  }

  const range = sheet.getRange(firstRow, mailCol, lastRow - firstRow + 1);
  const values = range.getValues();

  // メールアドレスのリストを取得
  const allAddresses = values.flat().filter(address => address !== '');

  // 重複を削除し、ユニークなメールアドレスリストを作成
  const uniqueAddresses = [...new Set(allAddresses)];

  // 重複数を計算
  const duplicateCount = allAddresses.length - uniqueAddresses.length;

  Logger.log('ユニークなメールアドレスリストを取得しました: ' + uniqueAddresses);
  Logger.log('重複したメールアドレスの数: ' + duplicateCount);

  let sentCount = 0;

  uniqueAddresses.forEach(address => {
    if (address) {
      try {
        sendMailToAll(address);
        Logger.log('メール送信に成功しました: ' + address);
        sentCount++;
      } catch (e) {
        Logger.log('メール送信に失敗しました: ' + address + ' エラー: ' + e.message);
      }
    }
  });

  Logger.log('送信されたメールアドレスの総数: ' + sentCount);
}

function sendMailToAll(address) {
  const subject = '【理工展/選挙投票】投票アドレス認証'; // メールの件名
  const body = `
メールアドレスを認証しました。

この度は理工展連絡会代表・副代表選挙にご参加ありがとうございました。

（このメールは有効投票数計算のために理工展連絡会選挙管理委員会より自動送信されています。）
`;

  GmailApp.sendEmail(address, subject, body);
}
