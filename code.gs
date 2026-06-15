const FOLDER_ID =
  "1bqo5tvKEDE2xrfvNQO0J0p8-J3y9jfAO";

const SHEET_NAME =
  "Payments";

function doPost(e){

  try{

    const data =
      JSON.parse(e.postData.contents);

    const folder =
      DriveApp.getFolderById(
        FOLDER_ID
      );

    const blob =
      Utilities.newBlob(
        Utilities.base64Decode(
          data.fileData
        ),
        data.mimeType,
        data.fileName
      );

    const uploadedFile =
      folder.createFile(blob);

    const url =
      uploadedFile.getUrl();

    const sheet =
      SpreadsheetApp
        .getActiveSpreadsheet()
        .getSheetByName(
          SHEET_NAME
        );

    sheet.appendRow([
      new Date(),
      data.nama,
      data.kategori,
      data.matrik || "",
      uploadedFile.getName(),
      url
    ]);

    return ContentService
      .createTextOutput(
        JSON.stringify({
          success:true,
          url:url
        })
      )
      .setMimeType(
        ContentService.MimeType.JSON
      );

  }catch(err){

    return ContentService
      .createTextOutput(
        JSON.stringify({
          success:false,
          error:err.toString()
        })
      )
      .setMimeType(
        ContentService.MimeType.JSON
      );
  }
}
