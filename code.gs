function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = JSON.parse(e.postData.contents);

  const rows = sheet.getDataRange().getValues();

  const newKey = data.slotKey;

  // 🔍 Check duplicate
  for (let i = 1; i < rows.length; i++) {

    const existingKey = rows[i][12]; // SlotKey column

    if (existingKey === newKey) {
      return ContentService.createTextOutput("duplicate");
    }
  }

  // ✅ Save booking
  sheet.appendRow([
    data.name,
    data.phone,
    data.email,
    data.address,
    data.district,
    data.zip,
    data.vehicle,
    data.date,
    data.time,
    data.plan,
    data.services,
    data.requirements,
    data.slotKey   // 🔥 IMPORTANT
  ]);

  return ContentService.createTextOutput("success");
}