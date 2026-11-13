/**
 * 創作ノート サービスワーカー (sw.js)
 *
 * このスクリプトは、Webアプリ「創作ノート」をオフラインでも起動できるようにし、
 * PWA（Progressive Web App）としてスマホやPCにインストール可能にするためのプログラムです。
 *
 * 【主な機能】
 * - アプリに必要な画面ファイル（HTMLやアイコン）をブラウザ内にキャッシュ（一時保存）
 * - 電波がない場所（オフライン）でもアプリ画面を表示可能にする
 * - 新しいバージョンが配信されたときに古いキャッシュを自動的にお掃除
 *
 * 【コーディングルール】
 * - 文字列は基本的にダブルクォーテーション（"）で囲む
 * - 関数と関数の間は2行あける
 * - 初心者にもわかりやすいドキュメンテーションおよび解説コメントをつける
 */


/**
 * キャッシュの保存先につける名前（バージョン）です。
 * アプリの画面やアイコンを大幅に更新したときは、この文字を変えると新しい内容が読み込まれます。
 */
const CACHE_NAME = "sousaku-note-cache-v1";


/**
 * オフラインでもアプリを開けるように、事前に保存しておくファイルの一覧です。
 */
const ASSETS_TO_CACHE = [
  "./",
  "./index.html",
  "./manifest.webmanifest",
  "./icon-192.png",
  "./icon-512.png"
];


/**
 * サービスワーカーが初めてインストールされたときに動く処理です。
 * アプリに必要な基本ファイルをまとめてキャッシュにダウンロード・保存します。
 *
 * @param {ExtendableEvent} event インストールイベント
 */
function handleInstall(event) {
  // キャッシュの保存が完了するまでインストール処理を待機させます
  event.waitUntil(
    caches.open(CACHE_NAME).then(function (cache) {
      return cache.addAll(ASSETS_TO_CACHE);
    }).then(function () {
      // 待機せずにすぐ有効化します
      return self.skipWaiting();
    })
  );
}


/**
 * 新しいサービスワーカーが有効化（アクティベート）されたときに動く処理です。
 * 古いバージョンのキャッシュが残っていれば自動的に削除してお掃除します。
 *
 * @param {ExtendableEvent} event アクティベートイベント
 */
function handleActivate(event) {
  event.waitUntil(
    caches.keys().then(function (cacheNames) {
      return Promise.all(
        cacheNames.map(function (cacheName) {
          // いまのバージョンと違う古いキャッシュがあれば削除します
          if (cacheName !== CACHE_NAME) {
            return caches.delete(cacheName);
          }
        })
      );
    }).then(function () {
      // すべてのタブや画面ですぐに新しいサービスワーカーを動作させます
      return self.clients.claim();
    })
  );
}


/**
 * 画面がファイルや通信を要求したとき（リクエスト送信時）に動く処理です。
 * ネットワーク優先で取得し、通信できない（オフラインの）場合はキャッシュから返します。
 * ※Google Apps ScriptへのPOST通信（データ保存）などはキャッシュを通さず直接インターネットへ流します。
 *
 * @param {FetchEvent} event フェッチイベント
 */
function handleFetch(event) {
  const request = event.request;

  // POSTリクエスト（データ送信など）やGET以外の通信はキャッシュの対象外にし、通常通り通信します
  if (request.method !== "GET") {
    return;
  }

  // Googleドライブ同期用スクリプト（script.google.com）への通信もキャッシュを通さず直接通信します
  if (request.url.includes("script.google.com")) {
    return;
  }

  // まず通常通りインターネットから取得を試み、失敗した場合（オフライン時）にキャッシュを使います
  event.respondWith(
    fetch(request)
      .then(function (response) {
        // 正常に取得できたら、次回のためにキャッシュも更新しておきます
        if (response && response.status === 200 && response.type === "basic") {
          const responseToCache = response.clone();
          caches.open(CACHE_NAME).then(function (cache) {
            cache.put(request, responseToCache);
          });
        }
        return response;
      })
      .catch(function () {
        // オフラインなどで取得できなかった場合は、キャッシュにある過去のファイルを探して返します
        return caches.match(request).then(function (cachedResponse) {
          if (cachedResponse) {
            return cachedResponse;
          }
          // ホーム画面の要求であれば index.html を返します
          if (request.mode === "navigate") {
            return caches.match("./index.html");
          }
        });
      })
  );
}


// イベントリスナーの登録（ブラウザからの通知を受け取ります）
self.addEventListener("install", handleInstall);
self.addEventListener("activate", handleActivate);
self.addEventListener("fetch", handleFetch);
