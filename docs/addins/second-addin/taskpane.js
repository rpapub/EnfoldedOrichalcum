/* global Office, getCachedToken, cacheToken, writeGraphExtension */

const CLIENT_ID      = "REPLACE_WITH_YOUR_CLIENT_ID";
const EXTENSION_NAME = "net.cprima.rpapub.SecondAddin";
const DIALOG_URL     = "https://rpapub.github.io/EnfoldedOrichalcum/shared/auth-dialog.html";

if (typeof Office !== "undefined") {
  Office.onReady(() => {
    // TODO: implement task pane logic
  });
}
