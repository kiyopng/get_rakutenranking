//スプレッドシート取得
var ss = SpreadsheetApp.getActiveSpreadsheet();
//シート名で指定
var sheet = ss.getSheetByName("ランキングトップ");

function myFunction() {
  
  //取得開始トースト
  ss.toast('取得開始','ランキング取得',-1);
  
  var url = sheet.getRange("C4").getValue();
  var response = UrlFetchApp.fetch(url);
  var html = response.getContentText('UTF-8');
  html = html.replace(/\r?\n/g, '');
  
  try{
    
    //page=4まで取得
    for(var pagenum=1;pagenum<=1;pagenum++){
      
      var rankNum = 1;
      var celNum = 7;
      
      //1ページ目以外はページ番号付与し再取得
      if(pagenum > 1){
        url = url + "?page=" + pagenum;
        response = UrlFetchApp.fetch(url);
        html = response.getContentText('UTF-8');
        html = html.replace(/\r?\n/g, '');
        //sleep
        Utilities.sleep(1000);
        
        rankNum = ((pagenum-1) * 80) + 1;
        celNum = ((pagenum-1) * 80) + 3;
        ss.toast(pagenum + 'ページ目取得開始','ランキング取得',-1);
      }
      
      if(url == "https://ranking.rakuten.co.jp/"){
        html = html.replace(/<script language=\"JavaScript\" type=\"text\/javascript\">(.*?)<\/script>/g,'');
        var ranking_doc = html.match(/<dl class=\"rnkRanking_upperbox\">(.*?)<\/dl>/g);
      }else{
        var ranking_doc = html.match(/<div class=\"rnkRanking_upperbox\">(.*?)<div class=\"rnkRanking_directcart\">/g);
      }
      
      var ranking_img = html.match(/<div class=\"rnkRanking_imageBox\">(.*?)<\/div>/g);
      
      var i = 0;
      
      for each(var doc in ranking_doc){
        if(url == "https://ranking.rakuten.co.jp/"){
          var itemLink = doc.match(/<dt class="rnkTop_itemName"><a href=\"(.*?)\">/)[1];
          var itemName = doc.match(/<dt class="rnkTop_itemName"><a href=\".*?\">(.*?)<\/a>/)[1];
          var sellerLink = doc.match(/<dd class=\"rnkTop_shop\"><a href=\"(.*?)\">/)[1];
          var sellerName = doc.match(/<dd class=\"rnkTop_shop\"><a href=\".*?\">(.*?)<\/a>/)[1];
          var itemPrice = doc.match(/<div class=\"rnkTop_price\">(.*?)<\/div>/)[1];
        }else{
          var itemLink = doc.match(/<div class="rnkRanking_itemName"><a href=\"(.*?)\">/)[1];
          var itemName = doc.match(/<div class="rnkRanking_itemName"><a href=\".*?\">(.*?)<\/a>/)[1];
          var sellerLink = doc.match(/<div class=\"rnkRanking_shop\"><a href=\"(.*?)\">/)[1];
          var sellerName = doc.match(/<div class=\"rnkRanking_shop\"><a href=\".*?\">(.*?)<\/a>/)[1];
          var itemPrice = doc.match(/<div class=\"rnkRanking_price\">(.*?)<\/div>/)[1];
        }
        var imageURL = ranking_img[i].match(/<img src=\"(.*?)\?/)[1];
        sheet.getRange('B'+celNum).setValue(rankNum);
        sheet.getRange('C'+celNum).setValue(itemName);
        sheet.getRange('D'+celNum).setValue(itemLink);
        sheet.getRange('E'+celNum).setValue(sellerName);
        //sheet.getRange('F'+celNum).setValue(sellerLink);
        sheet.getRange('F'+celNum).setValue(imageURL);
        sheet.getRange('G'+celNum).setValue(itemPrice);
    
        rankNum = rankNum + 1;
        celNum = celNum + 1;
        
        i++;
      }
      
    }
  
    //取得終了トースト
    ss.toast('取得終了','ランキング取得',1);
    
  }catch(e){
    //エラートースト
    ss.toast('エラー','ランキング取得',-1);
    Logger.log(e); 
  }
    
}

function clear(){
  var columnBVals = sheet.getRange('B:G').getValues();
  row = columnBVals.filter(String).length;
  //入力されているデータをクリア
  sheet.getRange(7,2,row,6).clearContent();
}
