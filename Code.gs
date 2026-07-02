/**
 * SipOrder - 飲料訂購系統 後端 Apps Script
 * 
 * 部署方式：
 * 1. 建立一個 Google 試算表。
 * 2. 點選「延伸模組」->「Apps Script」。
 * 3. 將此檔案內容覆蓋 `Code.gs`。
 * 4. 新增 `Index.html` 檔案並貼入前端代碼。
 * 5. 點選「部署」->「新增部署」->「網頁應用程式」。
 */

/**
 * 處理網頁應用程式的 GET 請求
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('SipOrder - 飲料訂購系統')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 接收前端提交的訂單資料並寫入試算表
 * 
 * @param {Object} orderData 訂單資料對象
 * @param {string} orderData.buyer 訂購者
 * @param {string} orderData.drink 飲料品項
 * @param {string} orderData.sugar 甜度
 * @param {string} orderData.ice 冰量
 * @param {number|string} orderData.quantity 數量
 * @return {Object} 提交結果狀態
 */
function submitOrder(orderData) {
  // 設定鎖定防止併發寫入衝突
  var lock = LockService.getScriptLock();
  try {
    // 嘗試取得鎖定，最多等待 15 秒
    lock.waitLock(15000);
    
    // 取得當前試算表中的第一個工作表
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[0];
    
    // 解析資料
    var buyer = orderData.buyer ? orderData.buyer.trim() : "";
    var drink = orderData.drink ? orderData.drink.trim() : "";
    var sugar = orderData.sugar || "";
    var ice = orderData.ice || "";
    var quantity = parseInt(orderData.quantity, 10);
    
    if (isNaN(quantity) || quantity <= 0) {
      quantity = 1;
    }
    
    // 驗證欄位
    if (!buyer) {
      return { success: false, message: "請輸入訂購者姓名！" };
    }
    if (!drink) {
      return { success: false, message: "請輸入或選擇飲料品項！" };
    }
    
    var timestamp = new Date();
    
    // 寫入新資料列：時間戳記、訂購者、飲料品項、甜度、冰量、數量、處理狀態
    sheet.appendRow([
      timestamp,
      buyer,
      drink,
      sugar,
      ice,
      quantity,
      "待處理"
    ]);
    
    return {
      success: true,
      message: "訂單提交成功！謝謝 " + buyer + "，阿芝已為您記錄。"
    };
    
  } catch (error) {
    return {
      success: false,
      message: "提交失敗，請稍後再試。錯誤資訊：" + error.toString()
    };
  } finally {
    // 釋放鎖定
    lock.releaseLock();
  }
}
