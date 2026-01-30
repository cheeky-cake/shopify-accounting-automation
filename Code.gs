function convertShopifyToQB() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("Shopify_Data"); // 元データを貼るシート
  const targetSheet = ss.getSheetByName("QB_Import");   // 変換後が出るシート
  
  const data = sourceSheet.getDataRange().getValues();
  const result = [["Date", "Account", "Debit", "Credit", "Description"]]; // ヘッダー
  
  // 2行目からループ開始
  for (let i = 1; i < data.length; i++) {
    let orderId = data[i][0]; // Name
    let date = data[i][1];    // Created at
    let total = data[i][2];   // Total
    let subtotal = data[i][3]; // Subtotal
    
    // 1. 銀行への入金 (Debit)
    result.push([date, "Checking Account", total, 0, "Shopify Order " + orderId]);
    
    // 2. 売上 (Credit)
    result.push([date, "Sales Income", 0, subtotal, "Shopify Order " + orderId]);
    
    // ここにTaxやShippingの行も追加可能
  }
  
  targetSheet.clear();
  targetSheet.getRange(1, 1, result.length, result[0].length).setValues(result);
  SpreadsheetApp.getUi().alert("QBインポート用データの作成が完了しました！");
}
