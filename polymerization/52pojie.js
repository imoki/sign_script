// 吾爱破解论坛自动签到
// 需配合“金山文档”中的表格内容

let sheetNameSubConfig = "52pojie" // 分配置表名称
let pushHeader = "【52pojie】"
let sheetNameConfig = "CONFIG"  // 总配置表
let sheetNamePush = "PUSH"  // 推送表名称
let flagSubConfig = 0; // 激活分配置工作表标志
let flagConfig = 0; // 激活主配置工作表标志
let flagPush = 0; // 激活推送工作表标志
let line = 21;  // 指定读取从第2行到第line行的内容
var message= "" // 待发送的消息
var messageOnlyError = 0;  // 0为只推送失败消息，1则为推送成功消息。
var messageNickname = 0;  // 1为用昵称替代单元格，0为不替代
var jsonPush = [
  {'name':'bark', 'key':'xxxxxx', 'flag':'0' },
  {'name':'pushplus', 'key':'xxxxxx', 'flag':'0' },
  {'name':'ServerChan', 'key':'xxxxxx', 'flag':'0' }] // 推送数据，flag=1则推送

flagConfig = ActivateSheet(sheetNameConfig);  // 激活推送表
// 主配置工作表存在
if(flagConfig == 1){
  console.log("开始读取主配置表")
  let name; // 名称
  let onlyError;
  let nickname;
  for (let i = 2; i <= 100; i++){
    // 从工作表中读取推送数据
    name = Application.Range("A" + i).Text
    onlyError = Application.Range("C" + i).Text
    nickname = Application.Range("D" + i).Text
    if(name == "")  // 如果为空行，则提前结束读取
    {
      break;  // 提前退出，提高效率
    }
    if(name == sheetNameSubConfig ){
      if(onlyError == "是"){
        messageOnlyError = 1;
        console.log("只推送错误消息")
      }

      if(nickname == "是"){
        messageNickname = 1;
        console.log("单元格用昵称替代")
      }
      
      break;  // 提前退出，提高效率
    } 

  }
}

flagPush = ActivateSheet(sheetNamePush);  // 激活推送表
// 推送工作表存在
if(flagPush == 1){
  console.log("开始读取推送工作表")
  let pushName; // 推送类型
  let pushKey;
  let pushFlag; // 是否推送标志
  for (let i = 2; i <= line; i++){
    // 从工作表中读取推送数据
    pushName = Application.Range("A" + i).Text
    pushKey = Application.Range("B" + i).Text
    pushFlag = Application.Range("C" + i).Text
    if(pushName == "")  // 如果为空行，则提前结束读取
    {
      break;
    }
    jsonPushHandle(pushName, pushFlag, pushKey)    
  }
  // console.log(jsonPush)
}

flagSubConfig =  ActivateSheet(sheetNameSubConfig);  // 激活分配置表
if(flagSubConfig == 1){
  console.log("开始读取分配置表")
  for (let i = 2; i <= line; i++){
    var cookie = Application.Range("A" + i).Text
    var exec = Application.Range("B" + i).Text
    if(cookie == "")  // 如果为空行，则提前结束读取
    {
      break;
    }
    if(exec == "是"){
      execHandle(cookie, i);
    }
  }

  push(message);  // 推送消息
}

// 总推送
function push(message){
  if(message != "")
  {
    message = pushHeader + message; // 加上推送头
    let length = jsonPush.length
    let name;
    let key;
    for(let i = 0; i < length; i++){
      if(jsonPush[i].flag == 1){
        name = jsonPush[i].name
        key = jsonPush[i].key
        if(name == "bark"){
          bark(message, key);
        }else if(name == "pushplus"){
          pushplus(message, key);
        }else if(name == "ServerChan"){
          serverchan(message, key);
        }
      }
    }
  }else{
    console.log("消息为空不推送")
  }
}

// 推送bark消息
function bark(message, key){
  if(key != ""){
    let url = 'https://api.day.app/' + key + "/" + message;
    // 若需要修改推送的分组，则将上面一行改为如下的形式
    // let url = 'https://api.day.app/' + bark_id + "/" + message + "?group=分组名";
    let resp = HTTP.get(url,
      {headers:{'Content-Type': 'application/x-www-form-urlencoded'}}
    )
    sleep(5000)
  }
}

// 推送pushplus消息
function pushplus(message, key){
  if(key != ""){
    url = 'http://www.pushplus.plus/send?token=' + key + '&content=' + message
    let resp = HTTP.fetch(url, {
      method: "get"
    })
    sleep(5000)
  }
}

// 推送serverchan消息
function serverchan(message, key){
  if(key != ""){
    url = "https://sctapi.ftqq.com/" + key + ".send"  + "?title=消息推送"  + "&desp=" + message
    let resp = HTTP.fetch(url, {
      method: "get"
    })
    sleep(5000)
  }
}

function sleep(d){
  for(var t = Date.now();Date.now() - t <= d;);
}

// 激活工作表函数
function ActivateSheet(sheetName){
  let flag = 0;
  try{
    // 激活工作表
    let sheet = Application.Sheets.Item(sheetName)
    sheet.Activate()
    console.log("激活工作表：" + sheet.Name)
    flag = 1;
  }catch{
    flag = 0;
    console.log("无法激活工作表，工作表可能不存在")
  }
  return flag;
}

// 对推送数据进行处理
function jsonPushHandle(pushName, pushFlag, pushKey){
  let length = jsonPush.length
  for(let i = 0; i < length; i++){
    if(jsonPush[i].name == pushName){
      if(pushFlag == "是"){
        jsonPush[i].flag = 1;
        jsonPush[i].key = pushKey;
      }
    }
  }
}

// 具体的执行函数
function execHandle(cookie, pos){
    let messageSuccess = "";
    let messageFail = "";
    let messageName = "";
    if(messageNickname == 1){
      messageName = Application.Range("C" + pos).Text
    }else{
      messageName = "单元格A" + pos + ""
    }

    try{
      var url1 = "https://www.52pojie.cn/CSPDREL2hvbWUucGhwP21vZD10YXNrJmRvPWRyYXcmaWQ9Mg==?wzwscspd=MC4wLjAuMA=="
      var url2 = 'https://www.52pojie.cn/home.php?mod=task&do=apply&id=2&referer=%2F'
      var url3 = 'https://www.52pojie.cn/home.php?mod=task&do=draw&id=2'
      cookie_json = cookie_to_json(cookie)
      try{
        htVC_2132_saltkey = cookie_json['htVC_2132_saltkey']
        htVC_2132_auth = cookie_json['htVC_2132_auth']
        cookie = "htVC_2132_saltkey=" + htVC_2132_saltkey + "; htVC_2132_auth=" + htVC_2132_auth + ";"
        headers={
          "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9",
          "Accept-Encoding": "gzip, deflate, br",
          "Accept-Language": "zh-CN,zh;q=0.9",
          "Cache-Control": "no-cache",
          "Connection": "keep-alive",
          "Cookie": cookie,
          "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/109.0.0.0 Safari/537.36",
        }

        let resp = HTTP.fetch(url1,{
          method: "get",
          headers: headers
        })
        cookie_set = resp.headers['set-cookie']
        cookie = cookie + cookie_set
        sleep(1000)

        headers["Cookie"] = cookie
        resp = HTTP.fetch(url2,{
          method: "get",
          headers: headers
        })
        cookie_set = resp.headers['set-cookie']
        cookie = cookie + cookie_set
        sleep(1000)

        headers["Cookie"] = cookie
        resp = HTTP.fetch(url3,{
          method: "get",
          headers: headers
        })
        console.log("签到完成")
        messageSuccess += messageName + "签到 "
      }catch{
        messageFail += messageName + "的cookie有误，请重新填写 "
        console.log(messageName + "的cookie有误，请重新填写")
      }
    }catch{
      messageFail += messageName + "失败 "
    }

    sleep(2000);
    if(messageOnlyError == 1)
    {
      message = messageFail
    }else
    {
      message = messageFail + " " + messageSuccess
    }
    console.log(message)
}



// cookie字符串转json格式
function cookie_to_json(cookies){
    var cookie_text = cookies;
    var arr = [];
    var text_to_split = cookie_text.split(";");
    for(var i in text_to_split){
        var tmp = text_to_split[i].split("=");
        arr.push('"'+tmp.shift().trim()+'":"'+tmp.join(":").trim()+'"')
    };
    var res ='{\n'+arr.join(",\n")+'\n}';
    return JSON.parse(res);
}
