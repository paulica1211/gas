/**
 * e-Gov法令APIから特許法（昭和三十四年法律第百二十一号）の条文を取得し、
 * Googleスプレッドシートに書き出すためのGoogle Apps Scriptです。
 */
function fetchAndWritePatentLaw() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = '特許法';
  let sheet = ss.getSheetByName(sheetName);

  // シートが存在すればクリアし、存在しなければ新規作成
  if (sheet) {
    sheet.clear();
    Logger.log(`シート「${sheetName}」をクリアしました。`);
  } else {
    sheet = ss.insertSheet(sheetName);
    Logger.log(`シート「${sheetName}」を新規作成しました。`);
  }

  // UIに処理中であることを表示
  ss.toast('法令APIからデータの取得を開始します...', '処理中', -1);

  try {
    // 1. 法令APIから特許法のデータを取得
    const lawNum = '334AC0000000121'; // 特許法の法令番号（昭和三十四年法律第百二十一号）
    const url = `https://laws.e-gov.go.jp/api/1/lawdata/${lawNum}`;
    
    // ブラウザからのアクセスを装うためのヘッダーを追加
    const options = {
      'muteHttpExceptions': true,
      'headers': {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
      }
    };
    
    Logger.log(`リクエストURL: ${url}`);
    const response = UrlFetchApp.fetch(url, options);
    
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();

    if (responseCode !== 200) {
      const responseHeaders = response.getHeaders();
      Logger.log(`エラー時のレスポンスヘッダー: ${JSON.stringify(responseHeaders)}`);
      throw new Error(`APIからのデータ取得に失敗しました。ステータスコード: ${responseCode}, URL: ${url}`);
    }
    Logger.log('APIからのデータ取得に成功しました。');
    ss.toast('データの解析と書き込みを開始します...', '処理中', -1);


    // 2. XMLデータを解析
    const document = XmlService.parse(responseBody);
    const root = document.getRootElement();

    // 名前空間を取得
    const namespace = root.getNamespace();

    Logger.log(`ルート要素名: ${root.getName()}`);
    Logger.log(`名前空間URI: ${namespace ? namespace.getURI() : 'なし'}`);

    // Law要素配下のLawBodyを取得（DataRoot > ApplData > Law > LawBodyの構造）
    let lawBody;

    // DataRoot構造の場合
    const applData = root.getChild('ApplData', namespace);
    if (applData) {
      const lawFullText = applData.getChild('LawFullText', namespace);
      if (lawFullText) {
        const law = lawFullText.getChild('Law', namespace);
        if (law) {
          lawBody = law.getChild('LawBody', namespace);
        }
      }
    } else {
      // 直接Law要素の場合
      lawBody = root.getChild('LawBody', namespace);
    }

    if (!lawBody) {
      Logger.log(`XML構造（最初の1000文字）: ${responseBody.substring(0, 1000)}`);
      throw new Error('LawBody要素が見つかりませんでした。');
    }

    // MainProvision要素を取得
    const mainProvision = lawBody.getChild('MainProvision', namespace);
    if (!mainProvision) {
      throw new Error('MainProvision要素が見つかりませんでした。');
    }

    // Article要素を全て取得（再帰的に検索）
    const articles = [];

    function extractArticles(element) {
      // 現在の要素がArticleの場合は追加
      const articleElements = element.getChildren('Article', namespace);
      articles.push(...articleElements);

      // Part, Chapter, Section配下も再帰的に探索
      ['Part', 'Chapter', 'Section'].forEach(tagName => {
        const children = element.getChildren(tagName, namespace);
        children.forEach(child => extractArticles(child));
      });
    }

    extractArticles(mainProvision);

    if (articles.length === 0) {
      throw new Error('Article要素が見つかりませんでした。');
    }

    Logger.log(`取得した条文数: ${articles.length}`);

    // 3. ヘッダーと書き込むデータを準備
    const header = [['条番号', '条文本文']];
    const data = articles.map(article => {
      // 条番号の取得（Num属性から）
      const articleNum = article.getAttribute('Num');
      const title = articleNum ? articleNum.getValue() : '不明';

      let content = '';

      // ArticleCaption（条文の見出し）を取得
      const articleCaption = article.getChild('ArticleCaption', namespace);
      if (articleCaption) {
        content += articleCaption.getText() + '\n';
      }

      // ArticleTitle（条の名称）を取得
      const articleTitle = article.getChild('ArticleTitle', namespace);
      if (articleTitle) {
        content += articleTitle.getText() + '\n';
      }

      // Paragraph（項）を全て取得して本文を結合
      const paragraphs = article.getChildren('Paragraph', namespace);
      paragraphs.forEach(paragraph => {
        const paragraphNum = paragraph.getChild('ParagraphNum', namespace);
        if (paragraphNum) {
          content += paragraphNum.getText() + ' ';
        }

        const paragraphSentence = paragraph.getChild('ParagraphSentence', namespace);
        if (paragraphSentence) {
          content += getTextRecursive(paragraphSentence) + '\n';
        }

        // Item（号）を全て取得
        const items = paragraph.getChildren('Item', namespace);
        items.forEach(item => {
          const itemTitle = item.getChild('ItemTitle', namespace);
          if (itemTitle) {
            content += '  ' + itemTitle.getText() + ' ';
          }

          const itemSentence = item.getChild('ItemSentence', namespace);
          if (itemSentence) {
            content += getTextRecursive(itemSentence) + '\n';
          }
        });
      });

      return [title, content.trim()];
    });

    // 再帰的にテキストを取得するヘルパー関数
    function getTextRecursive(element) {
      let text = '';
      const content = element.getAllContent();
      content.forEach(item => {
        const itemType = item.getType();
        if (itemType === XmlService.ContentTypes.TEXT) {
          text += item.getText();
        } else if (itemType === XmlService.ContentTypes.ELEMENT) {
          text += getTextRecursive(item.asElement());
        }
      });
      return text;
    }

    // 4. スプレッドシートにデータを書き込み
    sheet.getRange(1, 1, 1, 2).setValues(header).setFontWeight('bold');
    if (data.length > 0) {
      sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
    }
    
    // 列幅を自動調整
    sheet.autoResizeColumn(1);
    
    Logger.log('スプレッドシートへの書き込みが完了しました。');
    ss.toast('特許法の条文をシートに書き出しました。', '完了', 5);

  } catch (e) {
    // 5. エラー処理
    Logger.log(e);
    SpreadsheetApp.getUi().alert(`エラーが発生しました: ${e.message}`);
    ss.toast('エラーが発生しました。詳細はログを確認してください。', 'エラー', 5);
  }
}
