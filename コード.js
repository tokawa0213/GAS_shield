//Sheet setting
var CHANNEL_ACCESS_TOKEN = 'RZojywHPpyKen3RODvlShZy1p4pi/NGuwQFAQzsfyrRKBMDymjSkcK6NR6dKf8Wzmh/jG28AKFVqVwIQ6irE/56glDhetExqePRBAfo4FuppbrWCyqq0kFB+ZOYxrModAszF3lcBwpv0KODbi4HSoQdB04t89/1O/w1cDnyilFU='; 
var line_endpoint = 'https://api.line.me/v2/bot/message/reply';
var spreadsheet = SpreadsheetApp.openById('15u2W565QUaBLGMlFuiTEapImDAtlUMdZhMKA68TKUZw');
var sheet = spreadsheet.getSheetByName('RemixBought');
var folder_id = "16Iilahkk4hc4nvQ9OO6gTLzlAuL6fCKu"

function addJobQue(id,result,name_to_image){
  var newQue = {
    "id": id,
    "result":result,
    "name_to_image":name_to_image
  }
  cache = CacheService.getScriptCache();
  var data = cache.get("key");
  //cacheの中身がnullならば空配列に，nullでないならstrを配列に変換する.
  if(data==null){
    data = [];
  }else{
    data = data.split(';');
  }
  //オブジェクトであるnewDataをstrに変換して配列に追加.
  data.push(JSON.stringify(newQue));
  //配列を;で分割するstrに変換.
  cache.put("key", data.join(';'), 60*2); 
  return;
}

function delete_files(){
  var today = new Date()
  var folder = DriveApp.getFolderById(folder_id)
  var files = folder.getFiles()
  while (files.hasNext()){
    var file = files.next()
    if ((today.getTime() - file.getDateCreated().getTime())/(1000*60*60*24) >= 3){
      folder.removeFile(file)
    }
  }
}

function timeDrivenFunction(){
  cache = CacheService.getScriptCache();
  var data = cache.get("key");
  cache.remove("key");
  if(data==null){
    return;
  }else{
    data = data.split(';');
  }
  //配列の中身をstrからJSON(object)に戻し，処理を実行する
  for(var i=0; i<data.length; i++){
    data[i] = JSON.parse(data[i]);
    make_image(data[i].id, data[i].result,data[i].name_to_image);
  }
  return;
}

//ポストで送られてくるので、送られてきたJSONをパース
function doPost(e) {
  var json = JSON.parse(e.postData.contents);

  //返信するためのトークン取得
  var reply_token= json.events[0].replyToken;
  if (typeof reply_token === 'undefined') {
    return;
  }

  //送られたメッセージ内容を取得
  var message = json.events[0].message.text;  
    
  // メッセージを返信
  UrlFetchApp.fetch(line_endpoint, {
    'headers': {
      'Content-Type': 'application/json; charset=UTF-8',
      'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN,
    },
    'method': 'post',
    'payload': JSON.stringify({
      'replyToken': reply_token,
      'messages': [{
        'type': 'text',
        'text': pick_card() + "\nシートは出来上がるのに時間がかかります。5分程お待ち下さい\n\n76枚印刷されている事を確認してA4で印刷してください。\n\n※スプレッドシートは3日以内に削除されます。",
      }
      ],
    }),
  });
  return ContentService.createTextOutput(JSON.stringify({'content': 'post ok'})).setMimeType(ContentService.MimeType.JSON);
}

function createSpreadsheet(ssName, folderId) {
  var folder = DriveApp.getFolderById(folderId);
  var ss = 0;
  if (folder.getFilesByName(ssName).hasNext()) {
    Logger.log(ssName+"があります");
  } else {
    Logger.log(ssName+"を作成");
    var ssId = SpreadsheetApp.create(ssName).getId();
    var file = DriveApp.getFileById(ssId);
    folder.addFile(file);
    DriveApp.getRootFolder().removeFile(file);
    var ss = SpreadsheetApp.openById(ssId);
  }
  return ss;
}

function make_image(id,result,name_to_image){
  var spreadsheet_out = SpreadsheetApp.openById(id);
  var sheet_out = spreadsheet_out.getSheetByName('シート1');
  sheet_out.insertColumnsAfter(1, 1);
  sheet_out.insertRowsAfter(1,652);
  var start_row = 1
  var start_column = 1
  var row_count = 0
  for (r in result){
    for (card_info in result[r]){
      var image = UrlFetchApp.fetch(name_to_image[card_info])
      for(i = 0;i < result[r][card_info]; i++){
        sheet_out.insertImage(name_to_image[card_info],start_column,start_row).setWidth(868.1999999999999).setHeight(1212.7238095238095)
        start_column += 9
        if (start_column == 28){
          start_row += 57
          start_column = 1
          row_count+=1
          if(row_count == 3){
            start_row += 9
            row_count = 0
          }
        }
      }
    }
  }
}

function pick_card(){
  const pack_num = 15
  const card_types = 65
  const RR_kakutei = 2
  const RR_percent = 4
  const TR_percent = 2
  const R_kakutei= 4
  const R_percent = 2
  const UC_kakutei = 12
  const UC_t_kakutei = 9
  const prize_card = {"エリカ":1}
  const card_info_columns = 12
  const name_pos = 0
  const image_pos = 11
  const rare_range = {"common":[1,28],"rare":[29,35],"double_rare":[36,39],"uncommon":[40,51],"uncommon_t":[52,60],"trainer_rare":[61,65]}

  var card_num = pack_num*5
  
  var name_to_image = {}
  var all = sheet.getRange(1,1,card_types,card_info_columns).getValues()
  for (i=0;i<card_types;i++){
    name_to_image[all[i][name_pos]] = all[i][image_pos]
  }
  
  var common = all.slice(rare_range['common'][0]-1,rare_range['common'][1])
  var common_list = {}
  for (i=0;i<rare_range['common'][1]-rare_range['common'][0]+1;i++){
    common_list[common[i][name_pos]] = 0;
  }
  
  var rare = all.slice(rare_range['rare'][0]-1,rare_range['rare'][1])
  
  var rare_list = {}
  for (i=0;i<rare_range['rare'][1]-rare_range['rare'][0]+1;i++){
    rare_list[rare[i][name_pos]] = 0;
  }

  var double_rare = all.slice(rare_range['double_rare'][0]-1,rare_range['double_rare'][1])
  
  var double_rare_list = {}
  for (i=0;i<rare_range['double_rare'][1]-rare_range['double_rare'][0]+1;i++){
    double_rare_list[double_rare[i][name_pos]] = 0;
  }

  var uncommon = all.slice(rare_range['uncommon'][0]-1,rare_range['uncommon'][1])
  var uncommon_t = all.slice(rare_range['uncommon_t'][0]-1,rare_range['uncommon_t'][1])
  
  var uncommon_list = {}
  for (i=0;i<rare_range['uncommon'][1]-rare_range['uncommon'][0]+1;i++){
    uncommon_list[uncommon[i][name_pos]] = 0;
  }
  for (i=0;i<rare_range['uncommon_t'][1]-rare_range['uncommon_t'][0]+1;i++){
    uncommon_list[uncommon_t[i][name_pos]] = 0;
  }
  
  var trainer_rare = all.slice(rare_range['trainer_rare'][0]-1,rare_range['trainer_rare'][1])

  var trainer_rare_list = {}
  for (i=0;i<rare_range['trainer_rare'][1]-rare_range['trainer_rare'][0]+1;i++){
    trainer_rare_list[trainer_rare[i][name_pos]] = 0;
  }

  var result = {"RR":double_rare_list,"TR":trainer_rare_list,"R":rare_list,"UC":uncommon_list,"C":common_list,"Other":prize_card}
  
  //確定枠から埋める
  
  //RR 2 or 3
  for (i=0;i<RR_kakutei;i++){
    random = double_rare[Math.floor(Math.random()*(rare_range["double_rare"][1]-rare_range["double_rare"][0]+1))]
    result["RR"][random[name_pos]]+=1
    card_num-=1
  }
  if(Math.floor(Math.random()*RR_percent) == 1.0){
    random = double_rare[Math.floor(Math.random()*(rare_range["double_rare"][1]-rare_range["double_rare"][0]+1))]
    result["RR"][random[name_pos]]+=1
    card_num-=1
  }
  //TR 0 or 1
  if(Math.floor(Math.random()*TR_percent) == 1.0){
    random = trainer_rare[Math.floor(Math.random()*(rare_range["trainer_rare"][1]-rare_range["trainer_rare"][0]+1))]
    result["TR"][random[name_pos]]+=1
    card_num-=1
  }
  //R 4 or 5
  for(i =0;i<R_kakutei;i++){
    random = rare[Math.floor(Math.random()*(rare_range["rare"][1]-rare_range["rare"][0]+1))]
    result["R"][random[name_pos]]+=1
    card_num-=1
  }
  if(Math.random()*R_percent == 1.0){
    random = rare[Math.floor(Math.random()*(rare_range["rare"][1]-rare_range["rare"][0]+1))]
    result["R"][random[name_pos]]+=1
    card_num-=1
  }
  
  //UC trainer=> 9 other=>12
  for(i =0;i<UC_t_kakutei;i++){
    random = uncommon_t[Math.floor(Math.random()*(rare_range["uncommon_t"][1]-rare_range["uncommon_t"][0]+1))]
    result["UC"][random[name_pos]]+=1
    card_num-=1
  }
  for(i =0;i<UC_kakutei;i++){
    random = uncommon[Math.floor(Math.random()*(rare_range["uncommon"][1]-rare_range["uncommon"][0]+1))]
    result["UC"][random[name_pos]]+=1
    card_num-=1
  }  
  //C 残り
  for(i =0;i<card_num;i++){
    random = common[Math.floor(Math.random()*(rare_range["common"][1]-rare_range["common"][0]+1))]
    result["C"][random[name_pos]]+=1
  }
  var ans = ""
  for (r in result){
    ans += "--" + r + "--" +"\n\n"
    for (card_info in result[r]){
      if (result[r][card_info] != 0.0){
        ans += card_info + " " + result[r][card_info] + "\n"
      }
    }
    ans += "\n"
  }
  var name = new Date()
  name = name.toString() + Math.floor(Math.random()*10000).toString()
  var spreadsheet_out = createSpreadsheet(name,folder_id)
  addJobQue(spreadsheet_out.getId(),result,name_to_image)
  ans += "--SpreadSheetLink--\n\nhttps://docs.google.com/spreadsheets/d/" + spreadsheet_out.getId() +"/edit#gid=0 \n\n"
  ans += "--PDF--\n\nhttps://docs.google.com/spreadsheets/d/" + spreadsheet_out.getId() + "/export?exportFormat=pdf&amp;format=pdf&size=A4&top_margin=0.00&bottom_margin=0.00&left_margin=0.00&right_margin=0.00 \n\n"
  return ans
}
