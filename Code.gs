/**
 * 創作ノート 同期用 Google Apps Script (Code.gs)
 *
 * このスクリプトは、Webアプリ「創作ノート」のデータをGoogleドライブに保存・同期し、
 * Googleドキュメントやスプレッドシートへ書き出すためのバックエンドプログラムです。
 *
 * 【設定手順】
 * 1. Googleドライブまたは https://script.google.com から新しいGASプロジェクトを作成し、このコードを貼り付けます。
 * 2. 画面左の歯車アイコン「プロジェクトの設定」→「スクリプト プロパティ」に以下を追加します：
 *    - プロパティ: TOKEN
 *    - 値: お好きな合言葉（パスワード）
 * 3. 画面右上の「デプロイ」→「新しいデプロイ」から種類「ウェブアプリ」を選択します：
 *    - 次のユーザーとして実行: 「自分」
 *    - アクセスできるユーザー: 「全員」
 * 4. 初回デプロイ時にGoogleドライブのアクセス権限を承認します。
 * 5. 発行された「ウェブアプリのURL」（末尾が /exec のもの）をアプリの「保存」タブに入力します。
 *
 * 【コーディングルール】
 * - 文字列は基本的にダブルクォーテーション（"）で囲む
 * - 関数と関数の間は2行あける
 * - 初心者にもわかりやすいドキュメンテーションおよび解説コメントをつける
 */


/**
 * Googleドライブ上に作成する専用フォルダの名前です。
 */
const FOLDER = "創作ノート";


/**
 * 創作ノートの保存先となるGoogleドライブのフォルダを取得します。
 * まだフォルダが存在しない場合は自動的に新しく作成します。
 * 毎回フォルダを探すと時間がかかるため、見つかったフォルダのIDをスクリプトプロパティに記憶（キャッシュ）して高速化しています。
 *
 * @returns {GoogleAppsScript.Drive.Folder} 保存先のフォルダオブジェクト
 */
function folder_() {
  // スクリプトプロパティ（キーと値を保存できる設定領域）を取得します
  const properties = PropertiesService.getScriptProperties();
  const folderId = properties.getProperty("FID");

  // すでに記憶しているフォルダIDがあれば、それを使ってフォルダを開きます
  if (folderId) {
    try {
      return DriveApp.getFolderById(folderId);
    } catch (e) {
      // フォルダが削除されていたりIDが無効な場合は、下に進んで再検索・再作成します
    }
  }

  // 「創作ノート」という名前のフォルダをGoogleドライブ全体から探します
  const folderIterator = DriveApp.getFoldersByName(FOLDER);
  const folder = folderIterator.hasNext()
    ? folderIterator.next()
    : DriveApp.createFolder(FOLDER);

  // 次回以降すばやくアクセスできるように、フォルダのIDを記憶させておきます
  properties.setProperty("FID", folder.getId());
  return folder;
}


/**
 * アプリ（フロントエンド）へ返すJSON形式のHTTPレスポンスを作成します。
 *
 * @param {Object} data 返却したいデータ（JavaScriptのオブジェクト）
 * @returns {GoogleAppsScript.Content.TextOutput} JSON形式のテキスト出力
 */
function out_(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}


/**
 * 作品やシリーズのデータの中から、「現在使われている画像ID」をすべて集めます。
 * 使われていない不要な画像を整理（ガベージコレクション）するときに呼び出されます。
 *
 * @param {Object} data 作品またはシリーズのデータオブジェクト
 * @param {Object} used 画像IDをキー、1を値として記録する辞書オブジェクト
 */
function collect_(data, used) {
  // 人物（キャラクター）に設定されている画像を収集します
  const characters = data.chars || [];
  characters.forEach(function (character) {
    const images = character.imgs || [];
    images.forEach(function (imageId) {
      used[imageId] = 1;
    });
  });

  // 資料（設定資料）に設定されている画像を収集します
  const referenceGroups = data.refs || [];
  referenceGroups.forEach(function (group) {
    const items = group.items || [];
    items.forEach(function (item) {
      const images = item.imgs || [];
      images.forEach(function (imageId) {
        used[imageId] = 1;
      });
    });
  });
}


/**
 * Googleドキュメント（原稿）またはGoogleスプレッドシート（設定表）を新規作成し、
 * 「書き出し」フォルダに保存する処理を行います。
 *
 * @param {Object} query アプリから送られてきた書き出しリクエスト情報
 * @returns {GoogleAppsScript.Content.TextOutput} 作成結果（URLや名前など）を含むJSONレスポンス
 */
function exportOp_(query) {
  try {
    // 保存先のメインフォルダを取得します
    const rootFolder = folder_();

    // メインフォルダの中に「書き出し」という名前のサブフォルダがあるか確認し、無ければ作成します
    const subFolderIterator = rootFolder.getFoldersByName("書き出し");
    const subFolder = subFolderIterator.hasNext()
      ? subFolderIterator.next()
      : rootFolder.createFolder("書き出し");

    // ファイルのタイトルを決定します（長すぎる場合は200文字までに制限します）
    const title = String(query.title || "書き出し").slice(0, 200);

    // ----------------------------------------------------
    // 1. Googleドキュメント（原稿）への書き出し処理
    // ----------------------------------------------------
    if (query.kind === "doc") {
      // 新しいGoogleドキュメントを作成します
      const doc = DocumentApp.create(title);
      const body = doc.getBody();

      // 見出しの種類を定義します
      const headingStyles = {
        title: DocumentApp.ParagraphHeading.TITLE,
        h1: DocumentApp.ParagraphHeading.HEADING1,
        h2: DocumentApp.ParagraphHeading.HEADING2,
        h3: DocumentApp.ParagraphHeading.HEADING3
      };

      // 各段落ブロックを順番に追加していきます
      const blocks = query.blocks || [];
      blocks.forEach(function (block, index) {
        const text = String(block.text || "");
        let paragraph;

        if (index === 0) {
          // 最初の段落はドキュメント作成時に自動で作られている1行目を使用します
          paragraph = body.getParagraphs()[0];
          paragraph.setText(text);
        } else {
          // 2行目以降は新しい段落を追加します
          paragraph = body.appendParagraph(text);
        }

        // 指定された見出しスタイルを適用します（指定がなければ本文スタイル）
        const heading = headingStyles[block.t] || DocumentApp.ParagraphHeading.NORMAL;
        paragraph.setHeading(heading);
      });

      // 変更を保存してドキュメントを閉じます
      doc.saveAndClose();

      // 作成されたファイルはマイドライブ直下にできるため、「書き出し」フォルダに移動します
      const file = DriveApp.getFileById(doc.getId());
      file.moveTo(subFolder);

      // 返却用のデータを用意します
      const result = {
        ok: true,
        url: doc.getUrl(),
        name: title
      };

      // PDFも同時に作成するオプションがオンの場合
      if (query.pdf) {
        const pdfBlob = file.getAs("application/pdf");
        pdfBlob.setName(title + ".pdf");
        const pdfFile = subFolder.createFile(pdfBlob);
        result.pdf = pdfFile.getUrl();
      }

      return out_(result);
    }

    // ----------------------------------------------------
    // 2. Googleスプレッドシートへの書き出し処理
    // ----------------------------------------------------
    if (query.kind === "sheet") {
      // 新しいスプレッドシートを作成します
      const spreadsheet = SpreadsheetApp.create(title);
      const firstSheet = spreadsheet.getSheets()[0];

      // 数式インジェクション（=や+で始まる文字が数式として誤認・暴発する問題）を防ぐ関数です
      const sanitizeCellValue = function (value) {
        if (typeof value === "string" && /^[=+\-@]/.test(value)) {
          return "'" + value;
        }
        return value;
      };

      const sheetList = query.sheets || [];
      sheetList.forEach(function (sheetData, sheetIndex) {
        // 1枚目のシートは既存のシートを使い、2枚目以降は新しく挿入します
        const sheet = sheetIndex === 0 ? firstSheet : spreadsheet.insertSheet();

        // シートの名前を設定します（最大90文字）
        const sheetName = String(sheetData.name || ("シート" + (sheetIndex + 1))).slice(0, 90);
        sheet.setName(sheetName);

        const rows = sheetData.rows || [];
        if (rows.length === 0) {
          return;
        }

        // 列数の最大値を調べます（すべての行の幅を揃えるため）
        let maxColumns = 1;
        rows.forEach(function (row) {
          if (row.length > maxColumns) {
            maxColumns = row.length;
          }
        });

        // データを二次元配列として整えます（短い行には空文字を埋めて長方形にします）
        const tableData = rows.map(function (row) {
          const sanitizedRow = row.map(sanitizeCellValue);
          while (sanitizedRow.length < maxColumns) {
            sanitizedRow.push("");
          }
          return sanitizedRow;
        });

        // セルにデータを一括で書き込み、折り返しと上揃えを設定します
        sheet.getRange(1, 1, tableData.length, maxColumns)
          .setValues(tableData)
          .setWrap(true)
          .setVerticalAlignment("top");

        // 1行目（ヘッダー行）を固定し、太字・薄い青色の背景にして見やすくします
        sheet.setFrozenRows(1);
        sheet.getRange(1, 1, 1, maxColumns)
          .setFontWeight("bold")
          .setBackground("#E3ECF5");

        // 列の幅を見やすく自動調整します（長い文章が入りやすい列は広めにします）
        const headerRow = rows[0];
        for (let columnIndex = 0; columnIndex < maxColumns; columnIndex++) {
          const headerText = String(headerRow[columnIndex] || "");
          const isWideColumn = /出来事|内容|メモ|登場するシーン|役割|機能|場面|追加項目/.test(headerText);
          sheet.setColumnWidth(columnIndex + 1, isWideColumn ? 260 : 130);
        }
      });

      // マイドライブ直下から「書き出し」フォルダへ移動します
      DriveApp.getFileById(spreadsheet.getId()).moveTo(subFolder);

      return out_({
        ok: true,
        url: spreadsheet.getUrl(),
        name: title
      });
    }

    // doc でも sheet でもない未知の書き出し種別の場合
    return out_({
      ok: false,
      error: "unknown kind"
    });

  } catch (error) {
    // エラーが発生した場合はエラーメッセージを返します
    return out_({
      ok: false,
      error: String((error && error.message) || error)
    });
  }
}


/**
 * ブラウザから直接URLを開いたとき（GETリクエスト）に応答する関数です。
 * 動作確認用として機能します。
 *
 * @returns {GoogleAppsScript.Content.TextOutput} 接続確認メッセージを含むJSONレスポンス
 */
function doGet() {
  return out_({
    ok: true,
    msg: "sousaku-note sync"
  });
}


/**
 * アプリからデータ同期や保存のリクエストが送信されたとき（POSTリクエスト）に呼ばれるメイン処理です。
 *
 * @param {Object} e POSTリクエストの情報（送信されたデータなどを含むイベントオブジェクト）
 * @returns {GoogleAppsScript.Content.TextOutput} 実行結果を含むJSONレスポンス
 */
function doPost(e) {
  let query;

  // 送信されたJSON文字列を解析（パース）します
  try {
    query = JSON.parse(e.postData.contents);
  } catch (parseError) {
    return out_({
      ok: false,
      error: "bad request"
    });
  }

  // スクリプトプロパティに設定した合言葉（TOKEN）を取得し、照合します
  const savedToken = PropertiesService.getScriptProperties().getProperty("TOKEN");
  if (!savedToken || query.token !== savedToken) {
    return out_({
      ok: false,
      error: "unauthorized"
    });
  }

  // 作品IDやファイル名を安全な文字（英数字・ハイフン・アンダースコア）のみに正規化します
  const id = String(query.id || "main").replace(/[^A-Za-z0-9_-]/g, "");

  // ドキュメントやスプレッドシートの書き出しは時間がかかるため、排他ロックの外側で実行します
  if (query.op === "export") {
    return exportOp_(query);
  }

  // 同時に複数の端末から書き込みが衝突してデータが壊れるのを防ぐため、スクリプトロックを取得します
  const lock = LockService.getScriptLock();
  lock.waitLock(20000); // 最大20秒間ロックを待機

  try {
    // ----------------------------------------------------
    // 操作1: gc（不要な画像のクリーンアップ・整理）
    // どの作品・シリーズからも参照されていない画像をゴミ箱に移します（24時間以内の画像は誤削除防止のため保護）
    // ----------------------------------------------------
    if (query.op === "gc") {
      const folder = folder_();
      const files = folder.getFiles();
      const usedImages = {};
      const imageFiles = [];

      // フォルダ内のファイルを順に確認します
      while (files.hasNext()) {
        const file = files.next();
        const fileName = file.getName();

        // 画像ファイル（img_*.txt）の場合は一覧に追加して後でチェックします
        if (fileName.indexOf("img_") === 0) {
          imageFiles.push(file);
          continue;
        }

        // 作品データやインデックスデータ以外のファイルはスキップします
        if (fileName === "index.json" || !/\.json$/.test(fileName)) {
          continue;
        }

        // 作品・シリーズのJSONデータから、使われている画像IDを収集します
        try {
          const content = file.getBlob().getDataAsString("UTF-8");
          const jsonData = JSON.parse(content).data || {};
          collect_(jsonData, usedImages);
        } catch (readError) {
          return out_({
            ok: false,
            error: "読み取れないファイルがあるため中止しました: " + fileName
          });
        }
      }

      // 追加から24時間以内の画像は保護（ grace 期間 ）
      const graceTime = 24 * 3600 * 1000;
      const currentTime = Date.now();
      let trashedCount = 0;
      let trashedBytes = 0;

      imageFiles.forEach(function (imageFile) {
        // ファイル名 "img_<ID>.txt" から画像IDを抽出します
        const imageId = imageFile.getName().slice(4, -4);
        const fileAge = currentTime - imageFile.getDateCreated().getTime();

        // どこからも使われておらず、かつ作成から24時間以上経過している画像のみゴミ箱に移動します
        if (!usedImages[imageId] && fileAge > graceTime) {
          trashedBytes += imageFile.getSize();
          imageFile.setTrashed(true);
          trashedCount++;
        }
      });

      return out_({
        ok: true,
        trashed: trashedCount,
        bytes: trashedBytes
      });
    }

    // ----------------------------------------------------
    // 操作2: usage（ドライブの使用量調査）
    // 創作ノートフォルダ内の合計容量・画像枚数・作品ファイル数を計算します
    // ----------------------------------------------------
    if (query.op === "usage") {
      const files = folder_().getFiles();
      let totalBytes = 0;
      let imageCount = 0;
      let docCount = 0;

      while (files.hasNext()) {
        const file = files.next();
        totalBytes += file.getSize();

        if (file.getName().indexOf("img_") === 0) {
          imageCount++;
        } else {
          docCount++;
        }
      }

      return out_({
        ok: true,
        bytes: totalBytes,
        imgs: imageCount,
        docs: docCount
      });
    }

    // ----------------------------------------------------
    // 操作3: imgput / imgget（画像の保存および取得）
    // 画像はBase64形式のテキストデータとして img_<ID>.txt に個別に保存されます
    // ----------------------------------------------------
    if (query.op === "imgput" || query.op === "imgget") {
      const imageId = String(query.iid || "").replace(/[^A-Za-z0-9_-]/g, "");
      if (!imageId) {
        return out_({
          ok: false,
          error: "bad id"
        });
      }

      const folder = folder_();
      const imageFileName = "img_" + imageId + ".txt";
      const fileIterator = folder.getFilesByName(imageFileName);

      // 画像の取得 (imgget)
      if (query.op === "imgget") {
        if (fileIterator.hasNext()) {
          const fileData = fileIterator.next().getBlob().getDataAsString("UTF-8");
          return out_({
            ok: true,
            data: fileData
          });
        }
        return out_({
          ok: false,
          error: "not found"
        });
      }

      // 画像の保存 (imgput)
      // 同じIDの画像がまだ存在しない場合のみ新規作成します（既存の画像は上書き不要）
      if (!fileIterator.hasNext()) {
        folder.createFile(imageFileName, String(query.data || ""), MimeType.PLAIN_TEXT);
      }

      return out_({
        ok: true
      });
    }

    // ----------------------------------------------------
    // 操作4: del, stat, get, put（作品データの同期・操作）
    // 作品ごとのデータは <ID>.json（例: main.json や w123.json）に保存されます
    // ----------------------------------------------------
    const folder = folder_();
    const fileName = id + ".json";
    const fileIterator = folder.getFilesByName(fileName);
    const targetFile = fileIterator.hasNext() ? fileIterator.next() : null;

    // 作品の削除 (del)
    if (query.op === "del") {
      if (targetFile) {
        targetFile.setTrashed(true);
      }
      return out_({
        ok: true
      });
    }

    // 更新日時の確認 (stat)
    // ファイルの説明欄（description）にタイムスタンプを保存しているため、中身を開かずに高速確認できます
    if (query.op === "stat" && targetFile) {
      const descriptionTimestamp = Number(targetFile.getDescription());
      if (descriptionTimestamp) {
        return out_({
          ok: true,
          updated: descriptionTimestamp
        });
      }
    }

    // ファイルが存在する場合は内容を読み込んで解析します
    const currentData = targetFile
      ? JSON.parse(targetFile.getBlob().getDataAsString("UTF-8"))
      : null;

    // ファイルの説明欄が読めなかった場合のフォールバック（ファイルの中身から updated を取得）
    if (query.op === "stat") {
      return out_({
        ok: true,
        updated: currentData ? currentData.updated : 0
      });
    }

    // 作品データの取得 (get)
    if (query.op === "get") {
      return out_({
        ok: true,
        updated: currentData ? currentData.updated : 0,
        data: currentData ? currentData.data : null
      });
    }

    // 作品データの保存 (put)
    if (query.op === "put") {
      // 競合検知（コンフリクトチェック）：
      // 端末側が持っている前回の更新時刻と、サーバー上の現在の更新時刻が一致しない場合、
      // 別の端末で先に更新されているため、誤って上書きしないよう競合エラーを返します。
      // ※強制上書き（force）の場合は query.base が -1 になります。
      if (currentData && query.base !== -1 && query.base !== currentData.updated) {
        return out_({
          ok: false,
          error: "conflict",
          updated: currentData.updated
        });
      }

      const now = Date.now();
      const payload = JSON.stringify({
        updated: now,
        data: query.data
      });

      if (targetFile) {
        // すでにファイルがある場合は内容を更新します
        targetFile.setContent(payload);
        targetFile.setDescription(String(now));
      } else {
        // 新しい作品の場合はファイルを新規作成します
        folder.createFile(fileName, payload, MimeType.PLAIN_TEXT)
          .setDescription(String(now));
      }

      return out_({
        ok: true,
        updated: now
      });
    }

    // 未知の操作コマンドの場合
    return out_({
      ok: false,
      error: "unknown op"
    });

  } finally {
    // 処理が終わったら必ずロックを解放して、他の通信を受け入れられるようにします
    lock.releaseLock();
  }
}
