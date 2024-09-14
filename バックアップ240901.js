// // メール件名とGoogleドライブのフォルダIDを定義します。
// // これにより、コードの再利用性が高まり、特定のメールやフォルダを簡単に変更できます。
// const mailSubject = 'DMM：DMMプレミアム 会員登録完了のご連絡'; //例：アロマの蒸気と熱気でリフレッシュできるサウナや東京湾を眺められる露天風呂など、夏の疲れにぴったりの日帰り温泉をご紹介
// const driveFolderId = '1Dl1DRxWj7v0qri4Bqwr5cH3B5YAA9Egj'; //例：16wDTdRVca720l74Xj4SDa4TeGIO56zEM

// function emailToPDF() {
//   // 指定された件名のメールスレッドを検索し、最初に見つかったスレッドを取得します。
//   const thread = GmailApp.search('subject:"' + mailSubject + '"')[0];
  
//   // スレッド内の最初のメールメッセージを取得します。
//   const message = thread.getMessages()[0];
  
//   // 送信者の情報を取得し、正規表現で名前とメールアドレスを抽出します。
//   // 名前とメールアドレスが一緒に表示される場合、全体を使用し、
//   // そうでない場合は「Unknown <メールアドレス>」として表示します。
//   const from = message.getFrom();
//   const fromFormatted = from.match(/(.*?)(<.*>)/) ? from : `Unknown <${from}>`;

//   // メールの日付、件名、受信者情報を取得します。
//   const date = message.getDate();
//   const subject = message.getSubject();
//   const to = message.getTo();

//   // PDFの冒頭に表示するヘッダー部分をHTML形式で作成します。
//   const header = `
//     <div style="font-family: Arial, sans-serif; font-size: 12px;">
//       <strong>From:</strong> ${fromFormatted}<br>
//       <strong>Date:</strong> ${Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy年M月d日(E) HH:mm')}<br>
//       <strong>Subject:</strong> ${subject}<br>
//       <strong>To:</strong> ${to}<br><br>
//     </div>
//   `;
  
//   // メール本文のHTMLを取得し、ヘッダー部分を追加します。
//   let htmlBody = header + message.getBody();
  
//   // HTML内の改行コードを削除し、<p>タグと<div>タグのスタイルを調整します。
//   // これにより、PDF出力時の余分な改行やスペースを防ぎます。
//   htmlBody = htmlBody.replace(/\n/g, '').replace(/<p/g, '<p style="margin: 0; padding: 0;"').replace(/<div/g, '<div style="margin: 0; padding: 0;"');

//   // CID参照の画像をBase64エンコードされたデータURIに変換します。
//   htmlBody = inlineImages(htmlBody, message);
  
//   // 外部URL参照の画像をBase64エンコードされたデータURIに変換します。
//   htmlBody = inlineExternalImages(htmlBody);
  
//   // HTMLをPDF形式に変換し、ファイル名を「日時＋メール件名」に設定します。
//   const blob = Utilities.newBlob(htmlBody, 'text/html', 'email.html');
//   const pdfName = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyyMMdd_HHmm') + ' ' + subject + '.pdf';
//   const pdf = blob.getAs('application/pdf').setName(pdfName);
  
//   // 生成されたPDFを指定されたGoogleドライブのフォルダに保存します。
//   const folder = DriveApp.getFolderById(driveFolderId);
//   folder.createFile(pdf);
// }

// function inlineImages(html, message) {
//   // メール内のCID参照画像をBase64エンコードされたデータURIに変換します。
//   // CID (Content-ID) はメール内で埋め込み画像を指定するために使用されます。
//   return html.replace(/<img[^>]+src="cid:([^">]+)"/g, function(match, cid) {
//     try {
//       // メールに添付されているすべてのファイルを取得します。
//       const attachments = message.getAttachments();
      
//       // 各添付ファイルを順番に処理します。
//       for (let i = 0; i < attachments.length; i++) {
//         const attachment = attachments[i];
        
//         // ファイルのコンテンツタイプ（MIMEタイプ）を取得します。
//         const contentType = attachment.getContentType();
        
//         // ファイルのバイトデータをBase64エンコードします。
//         const base64 = Utilities.base64Encode(attachment.getBytes());
        
//         // Base64エンコードされたデータをデータURIスキームに変換します。
//         const dataUri = 'data:' + contentType + ';base64,' + base64;
        
//         // CIDが一致する場合、その画像をHTML内で置換します。
//         if (attachment.getContentId() && attachment.getContentId() === cid) {
//           return match.replace('cid:' + cid, dataUri);
//         }
//       }
//     } catch (e) {
//       // 画像の取得に失敗した場合、エラーログにメッセージを記録します。
//       Logger.log('Image fetch failed for CID: ' + cid + ' Error: ' + e.toString());
//     }
//     return match; // 画像が見つからない場合、元のHTMLコードをそのまま返します。
//   });
// }

// function inlineExternalImages(html) {
//   // HTML内の外部URL参照画像をBase64エンコードされたデータURIに変換します。
//   return html.replace(/<img[^>]+src="([^">]+)"/g, function(match, url) {
//     // CIDで始まるURL（すでに処理済みのもの）はスキップします。
//     if (url.startsWith('cid:')) {
//       return match;
//     }
//     try {
//       // URLが#を含む場合、#以降の部分を実際のURLとして使用します。
//       const actualUrl = url.split('#')[1] || url;
      
//       // 画像を取得するためのリクエストオプションを設定します。
//       const options = {
//         'followRedirects': true, // リダイレクトを自動で追跡
//         'muteHttpExceptions': true, // HTTPエラーを無視して処理を続行
//         'headers': {
//           'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36' // ブラウザと同じユーザーエージェントを使用
//         }
//       };
      
//       // 外部URLから画像データを取得し、Base64エンコードしてデータURIに変換します。
//       const response = UrlFetchApp.fetch(actualUrl, options);
//       if (response.getResponseCode() === 200) {
//         const contentType = response.getHeaders()['Content-Type']; // コンテンツタイプを取得
//         const base64 = Utilities.base64Encode(response.getContent()); // 画像データをBase64エンコード
//         const dataUri = 'data:' + contentType + ';base64,' + base64; // データURI形式に変換
//         return match.replace(url, dataUri); // HTML内の画像URLをデータURIで置換
//       } else {
//         // 画像の取得に失敗した場合、エラーログにメッセージを記録します。
//         Logger.log('Failed to fetch image from URL: ' + actualUrl + ' Error: HTTP ' + response.getResponseCode());
//       }
//     } catch (e) {
//       // 画像の取得に失敗した場合、エラーログにメッセージを記録します。
//       Logger.log('Failed to fetch image from URL: ' + url + ' Error: ' + e.toString());
//     }
//     return match; // 画像が取得できない場合、元のHTMLコードをそのまま返します。
//   });
// }
