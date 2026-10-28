// ============================================================
// Firebase 設定檔 — 請照 README.md 的步驟，把你自己的設定貼進來
// 貼上後存檔，重新整理網頁即可連線
// ============================================================

const FIREBASE_CONFIG = {
  apiKey: "AIzaSyAlyudP_neS2YQZVdaV1xvXSku_AtRXu9E",
  authDomain: "quizlive-d8e26.firebaseapp.com",
  databaseURL: "https://quizlive-d8e26-default-rtdb.asia-southeast1.firebasedatabase.app",   // ← 最重要！長得像 https://xxx-default-rtdb.asia-southeast1.firebasedatabase.app
  projectId: "quizlive-d8e26",
};

function firebaseReady() {
  return (
    typeof FIREBASE_CONFIG === "object" &&
    typeof FIREBASE_CONFIG.databaseURL === "string" &&
    FIREBASE_CONFIG.databaseURL.startsWith("https://")
  );
}
