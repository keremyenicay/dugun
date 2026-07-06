/**
 * Gülçiçek & Kerem — Düğün Projesi (V5 - Galeri, Yükleme ve Ortak Anı Defteri)
 *
 * ÖNEMLİ: Bu dosyayı Google Apps Script projenize kopyalayıp YENİDEN DAĞITIN
 * (Deploy > Manage deployments > Edit > New version). Galeri fotoğraflarının ve
 * ortak anı defterinin çalışması için bu güncel sürüm gereklidir.
 */

const FOLDER_ID = "1cL9XFnM12fC4epSQ0MfWjFbLhOlILAn4";
const MESSAGES_FILE = "misafir_defteri.json"; // Anı defteri mesajlarının saklandığı dosya

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // 1. Galeri için dosya listesini çek
    if (data.action === "getFiles") {
      return _json(_getFilesFromDrive());
    }

    // 2. Anı defteri mesajlarını çek
    if (data.action === "getMessages") {
      return _json(_getMessages());
    }

    // 3. Anı defterine yeni mesaj kaydet (herkeste ortak)
    if (data.action === "saveMessage") {
      const messages = _saveMessage(data.name, data.text);
      return _json({ ok: true, messages: messages });
    }

    // 4. Yükleme oturumu oluştur
    if (data.action === "create") {
      const res = UrlFetchApp.fetch(
        "https://www.googleapis.com/upload/drive/v3/files?uploadType=resumable&supportsAllDrives=true",
        {
          method: "post",
          contentType: "application/json",
          headers: {
            Authorization: "Bearer " + ScriptApp.getOAuthToken(),
            Origin: data.origin || "https://keremyenicay.github.io"
          },
          payload: JSON.stringify({
            name: data.fileName,
            mimeType: data.mimeType,
            parents: [FOLDER_ID]
          }),
          muteHttpExceptions: true
        }
      );

      const headers = res.getAllHeaders();
      const uploadUrl = headers["Location"] || headers["location"];

      if (!uploadUrl) {
        return _json({ error: "Oturum açılamadı: " + res.getContentText() });
      }
      return _json({ uploadUrl: uploadUrl, v: 5 });
    }

    return _json({ error: "Bilinmeyen istek" });
  } catch (err) {
    return _json({ error: String(err) });
  }
}

/**
 * Drive'daki görselleri çeken yardımcı fonksiyon.
 * Oturum gerektiren getDownloadUrl() yerine herkese açık "thumbnail" adresini
 * döndürüyoruz; böylece "Bağlantıya sahip herkes" olarak paylaşılan görseller
 * sitede görünür.
 */
function _getFilesFromDrive() {
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const files = folder.getFiles();
  const fileList = [];

  while (files.hasNext()) {
    const file = files.next();
    if (file.getMimeType().indexOf("image") > -1) {
      try {
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (e) { /* zaten paylaşımlıysa veya izin yoksa geç */ }

      const id = file.getId();
      fileList.push("https://drive.google.com/thumbnail?id=" + id + "&sz=w800");
    }
  }
  return fileList;
}

/* ================= ANI DEFTERİ (ORTAK MESAJLAR) ================= */

// Mesaj dosyasını bul; yoksa oluştur.
function _getMessagesFile(folder) {
  folder = folder || DriveApp.getFolderById(FOLDER_ID);
  const it = folder.getFilesByName(MESSAGES_FILE);
  if (it.hasNext()) return it.next();
  return folder.createFile(MESSAGES_FILE, "[]", "application/json");
}

function _getMessages() {
  try {
    const file = _getMessagesFile();
    const arr = JSON.parse(file.getBlob().getDataAsString() || "[]");
    return Array.isArray(arr) ? arr : [];
  } catch (e) {
    return [];
  }
}

function _saveMessage(name, text) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000); // Aynı anda gelen yazımların birbirini ezmesini önle
  try {
    const folder = DriveApp.getFolderById(FOLDER_ID);

    // Mevcut mesajları oku
    let arr = [];
    const it = folder.getFilesByName(MESSAGES_FILE);
    if (it.hasNext()) {
      try { arr = JSON.parse(it.next().getBlob().getDataAsString() || "[]"); } catch (e) { arr = []; }
    }
    if (!Array.isArray(arr)) arr = [];

    // Girişi temizle (basit güvenlik)
    const clean = (s, max) => String(s == null ? "" : s).replace(/[<>]/g, "").trim().slice(0, max);
    const entry = { name: clean(name, 80) || "Misafir", text: clean(text, 600), ts: Date.now() };
    if (!entry.text) return arr;

    arr.unshift(entry);
    if (arr.length > 500) arr = arr.slice(0, 500);

    // Eski dosyaları çöpe at, güncel içerikle yeniden yaz (DriveApp'ta setContent yok)
    const del = folder.getFilesByName(MESSAGES_FILE);
    while (del.hasNext()) del.next().setTrashed(true);
    folder.createFile(MESSAGES_FILE, JSON.stringify(arr), "application/json");

    return arr;
  } finally {
    lock.releaseLock();
  }
}

// Kurulumun çalışıp çalışmadığını test etmek için
function doGet(e) {
  if (e.parameter && e.parameter.action === "getFiles") return _json(_getFilesFromDrive());
  if (e.parameter && e.parameter.action === "getMessages") return _json(_getMessages());
  return _json({ status: "OK", message: "Api Aktif", v: 5 });
}

function _json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Yetkilendirme için gerekli boş fonksiyon
function yetkiVer() {
  DriveApp.getFolderById(FOLDER_ID).getName();
}
