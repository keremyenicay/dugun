/**
 * Gülçiçek & Kerem — Düğün Projesi (V4 - Galeri ve Yükleme Destekli)
 *
 * ÖNEMLİ: Bu dosyayı Google Apps Script projenize kopyalayıp YENİDEN DAĞITIN
 * (Deploy > Manage deployments > Edit > New version). Galeri fotoğraflarının
 * sitede görünmesi için bu güncel sürüm gereklidir.
 */

const FOLDER_ID = "1cL9XFnM12fC4epSQ0MfWjFbLhOlILAn4";

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    // 1. Galeri için dosya listesini çek
    if (data.action === "getFiles") {
      return _json(_getFilesFromDrive());
    }

    // 2. Yükleme oturumu oluştur
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
      return _json({ uploadUrl: uploadUrl, v: 4 });
    }

    return _json({ error: "Bilinmeyen istek" });
  } catch (err) {
    return _json({ error: String(err) });
  }
}

/**
 * Drive'daki görselleri çeken yardımcı fonksiyon.
 *
 * DÜZELTME: Eski sürüm file.getDownloadUrl() döndürüyordu; bu adres oturum
 * (login) gerektirdiği için <img>/arka plan olarak GÖRÜNTÜLENEMİYORDU.
 * Bunun yerine herkese açık "thumbnail" adresini döndürüyoruz. Böylece
 * "Bağlantıya sahip herkes" olarak paylaşılan görseller sitede görünür.
 */
function _getFilesFromDrive() {
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const files = folder.getFiles();
  const fileList = [];

  while (files.hasNext()) {
    const file = files.next();
    if (file.getMimeType().indexOf("image") > -1) {
      // Dosyayı bağlantıya sahip herkesin görebilmesini garanti et
      try {
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (e) { /* zaten paylaşımlıysa veya izin yoksa geç */ }

      const id = file.getId();
      fileList.push("https://drive.google.com/thumbnail?id=" + id + "&sz=w800");
    }
  }
  return fileList;
}

// Kurulumun çalışıp çalışmadığını test etmek için
function doGet(e) {
  if (e.parameter && e.parameter.action === "getFiles") {
    return _json(_getFilesFromDrive());
  }
  return _json({ status: "OK", message: "Api Aktif", v: 4 });
}

function _json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Yetkilendirme için gerekli boş fonksiyon
function yetkiVer() {
  DriveApp.getFolderById(FOLDER_ID).getName();
}
