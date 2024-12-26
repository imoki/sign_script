/*
    name: "vivo社区签到和抽奖"
    cron: 45 0 9 * * *
    脚本兼容: 金山文档（1.0），金山文档（2.0）
    更新时间：20241226
    环境变量名：无
    环境变量值：无
    备注：需要Cookie。
          cookie填写vivo社区网页版中获取的refresh_token。F12 -> NetWork(中文名叫"网络") -> 按一下Ctrl+R -> newbbs/ -> cookie
          vivo社区网址：https://bbs.vivo.com.cn/newbbs/
*/

var sheetNameSubConfig = "vivo"; // 分配置表名称
var pushHeader = "【vivo社区】";
var sheetNameConfig = "CONFIG"; // 总配置表
var sheetNamePush = "PUSH"; // 推送表名称
var sheetNameEmail = "EMAIL"; // 邮箱表
var flagSubConfig = 0; // 激活分配置工作表标志
var flagConfig = 0; // 激活主配置工作表标志
var flagPush = 0; // 激活推送工作表标志
var line = 21; // 指定读取从第2行到第line行的内容
var message = ""; // 待发送的消息
var messageArray = [];  // 待发送的消息数据，每个元素都是某个账号的消息。目的是将不同用户消息分离，方便个性化消息配置
var messageOnlyError = 0; // 0为只推送失败消息，1则为推送成功消息。
var messageNickname = 0; // 1为推送位置标识（昵称/单元格Ax（昵称为空时）），0为不推送位置标识
var messageHeader = []; // 存放每个消息的头部，如：单元格A3。目的是分离附加消息和执行结果消息
var messagePushHeader = pushHeader; // 存放在总消息的头部，默认是pushHeader,如：【xxxx】
var version = 1 // 版本类型，自动识别并适配。默认为airscript 1.0，否则为2.0（Beta）

var jsonPush = [
  { name: "bark", key: "xxxxxx", flag: "0" },
  { name: "pushplus", key: "xxxxxx", flag: "0" },
  { name: "ServerChan", key: "xxxxxx", flag: "0" },
  { name: "email", key: "xxxxxx", flag: "0" },
  { name: "dingtalk", key: "xxxxxx", flag: "0" },
  { name: "discord", key: "xxxxxx", flag: "0" },
]; // 推送数据，flag=1则推送
var jsonEmail = {
  server: "",
  port: "",
  sender: "",
  authorizationCode: "",
}; // 有效邮箱配置

// =================青龙适配开始===================

qlSwitch = 0

// =================青龙适配结束===================

// =================金山适配开始===================
// airscript检测版本
function checkVesion(){
  try{
    let temp = Application.Range("A1").Text;
    Application.Range("A1").Value  = temp
    console.log("😶‍🌫️ 检测到当前airscript版本为1.0，进行1.0适配")
  }catch{
    console.log("😶‍🌫️ 检测到当前airscript版本为2.0，进行2.0适配")
    version = 2
  }
}

// 推送相关
// 获取时间
function getDate(){
  let currentDate = new Date();
  currentDate = currentDate.getFullYear() + '/' + (currentDate.getMonth() + 1).toString() + '/' + currentDate.getDate().toString();
  return currentDate
}

// 将消息写入CONFIG表中作为消息队列，之后统一发送
function writeMessageQueue(message){
  // 当天时间
  let todayDate = getDate()
  flagConfig = ActivateSheet(sheetNameConfig); // 激活主配置表
  // 主配置工作表存在
  if (flagConfig == 1) {
    console.log("✨ 开始将结果写入主配置表");
    for (let i = 2; i <= 100; i++) {
      if(version == 1){
        // 找到指定的表行
        if(Application.Range("A" + (i + 2)).Value == sheetNameSubConfig){
          // 写入更新的时间
          Application.Range("F" + (i + 2)).Value = todayDate
          // 写入消息
          Application.Range("G" + (i + 2)).Value = message
          console.log("✨ 写入结果完成");
          break;
        }
      }else{
        // 找到指定的表行
        if(Application.Range("A" + (i + 2)).Value2 == sheetNameSubConfig){
          // 写入更新的时间
          Application.Range("F" + (i + 2)).Value2 = todayDate
          // 写入消息
          Application.Range("G" + (i + 2)).Value2 = message
          console.log("✨ 写入结果完成");
          break;
        }
      }
      
    }
  }
}

// 总推送
function push(message) {
  writeMessageQueue(message)  // 将消息写入CONFIG表中
  // if (message != "") {
  //   // message = messagePushHeader + message // 消息头最前方默认存放：【xxxx】
  //   let length = jsonPush.length;
  //   let name;
  //   let key;
  //   for (let i = 0; i < length; i++) {
  //     if (jsonPush[i].flag == 1) {
  //       name = jsonPush[i].name;
  //       key = jsonPush[i].key;
  //       if (name == "bark") {
  //         bark(message, key);
  //       } else if (name == "pushplus") {
  //         pushplus(message, key);
  //       } else if (name == "ServerChan") {
  //         serverchan(message, key);
  //       } else if (name == "email") {
  //         email(message);
  //       } else if (name == "dingtalk") {
  //         dingtalk(message, key);
  //       } else if (name == "discord") {
  //         discord(message, key);
  //       }
  //     }
  //   }
  // } else {
  //   console.log("🍳 消息为空不推送");
  // }
}

// 推送bark消息
function bark(message, key) {
    if (key != "") {
      message = messagePushHeader + message // 消息头最前方默认存放：【xxxx】
      message = encodeURIComponent(message)
      BARK_ICON = "https://s21.ax1x.com/2024/06/23/pkrUkfe.png"
    let url = "https://api.day.app/" + key + "/" + message + "/" + "?icon=" + BARK_ICON;
    // 若需要修改推送的分组，则将上面一行改为如下的形式
    // let url = 'https://api.day.app/' + bark_id + "/" + message + "?group=分组名";
    let resp = HTTP.get(url, {
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
    });
    sleep(5000);
    }
}

// 推送pushplus消息
function pushplus(message, key) {
  if (key != "") {
      message = encodeURIComponent(message)
    // url = "http://www.pushplus.plus/send?token=" + key + "&content=" + message;
    url = "http://www.pushplus.plus/send?token=" + key + "&content=" + message + "&title=" + pushHeader;  // 增加标题
    let resp = HTTP.fetch(url, {
      method: "get",
    });
    sleep(5000);
  }
}

// 推送serverchan消息
function serverchan(message, key) {
  if (key != "") {
    url =
      "https://sctapi.ftqq.com/" +
      key +
      ".send" +
      "?title=" + messagePushHeader +
      "&desp=" +
      message;
    let resp = HTTP.fetch(url, {
      method: "get",
    });
    sleep(5000);
  }
}

// email邮箱推送
function email(message) {
  var myDate = new Date(); // 创建一个表示当前时间的 Date 对象
  var data_time = myDate.toLocaleDateString(); // 获取当前日期的字符串表示
  let server = jsonEmail.server;
  let port = parseInt(jsonEmail.port); // 转成整形
  let sender = jsonEmail.sender;
  let authorizationCode = jsonEmail.authorizationCode;

  let mailer;
  mailer = SMTP.login({
    host: server,
    port: port,
    username: sender,
    password: authorizationCode,
    secure: true,
  });
  mailer.send({
    from: pushHeader + "<" + sender + ">",
    to: sender,
    subject: pushHeader + " - " + data_time,
    text: message,
  });
  // console.log("🍳 已发送邮件至：" + sender);
  console.log("🍳 已发送邮件");
  sleep(5000);
}

// 邮箱配置
function emailConfig() {
  console.log("🍳 开始读取邮箱配置");
  let length = jsonPush.length; // 因为此json数据可无序，因此需要遍历
  let name;
  for (let i = 0; i < length; i++) {
    name = jsonPush[i].name;
    if (name == "email") {
      if (jsonPush[i].flag == 1) {
        let flag = ActivateSheet(sheetNameEmail); // 激活邮箱表
        // 邮箱表存在
        // var email = {
        //   'email':'', 'port':'', 'sender':'', 'authorizationCode':''
        // } // 有效配置
        if (flag == 1) {
          console.log("🍳 开始读取邮箱表");
          for (let i = 2; i <= 2; i++) {
            // 从工作表中读取推送数据
            jsonEmail.server = Application.Range("A" + i).Text;
            jsonEmail.port = Application.Range("B" + i).Text;
            jsonEmail.sender = Application.Range("C" + i).Text;
            jsonEmail.authorizationCode = Application.Range("D" + i).Text;
            if (Application.Range("A" + i).Text == "") {
              // 如果为空行，则提前结束读取
              break;
            }
          }
          // console.log(jsonEmail)
        }
        break;
      }
    }
  }
}

// 推送钉钉机器人
function dingtalk(message, key) {
  message = messagePushHeader + message // 消息头最前方默认存放：【xxxx】
  let url = "https://oapi.dingtalk.com/robot/send?access_token=" + key;
  let resp = HTTP.post(url, { msgtype: "text", text: { content: message } });
  // console.log(resp.text())
  sleep(5000);
}

// 推送Discord机器人
function discord(message, key) {
  message = messagePushHeader + message // 消息头最前方默认存放：【xxxx】
  let url = key;
  let resp = HTTP.post(url, { content: message });
  //console.log(resp.text())
  sleep(5000);
}

// =================金山适配结束===================
// =================共用开始===================
// main()  // 入口

// function main(){
  checkVesion() // 版本检测，以进行不同版本的适配

  flagConfig = ActivateSheet(sheetNameConfig); // 激活推送表
  // 主配置工作表存在
  if (flagConfig == 1) {
    console.log("🍳 开始读取主配置表");
    let name; // 名称
    let onlyError;
    let nickname;
    for (let i = 2; i <= 100; i++) {
      // 从工作表中读取推送数据
      name = Application.Range("A" + i).Text;
      onlyError = Application.Range("C" + i).Text;
      nickname = Application.Range("D" + i).Text;
      if (name == "") {
        // 如果为空行，则提前结束读取
        break; // 提前退出，提高效率
      }
      if (name == sheetNameSubConfig) {
        if (onlyError == "是") {
          messageOnlyError = 1;
          console.log("🍳 只推送错误消息");
        }

        if (nickname == "是") {
          messageNickname = 1;
          console.log("🍳 单元格用昵称替代");
        }

        break; // 提前退出，提高效率
      }
    }
  }

  flagPush = ActivateSheet(sheetNamePush); // 激活推送表
  // 推送工作表存在
  if (flagPush == 1) {
    console.log("🍳 开始读取推送工作表");
    let pushName; // 推送类型
    let pushKey;
    let pushFlag; // 是否推送标志
    for (let i = 2; i <= line; i++) {
      // 从工作表中读取推送数据
      pushName = Application.Range("A" + i).Text;
      pushKey = Application.Range("B" + i).Text;
      pushFlag = Application.Range("C" + i).Text;
      if (pushName == "") {
        // 如果为空行，则提前结束读取
        break;
      }
      jsonPushHandle(pushName, pushFlag, pushKey);
    }
    // console.log(jsonPush)
  }

  // 邮箱配置函数
  emailConfig();

  flagSubConfig = ActivateSheet(sheetNameSubConfig); // 激活分配置表
  if (flagSubConfig == 1) {
    console.log("🍳 开始读取分配置表");

      if(qlSwitch != 1){  // 金山文档
          for (let i = 2; i <= line; i++) {
              var cookie = Application.Range("A" + i).Text;
              var exec = Application.Range("B" + i).Text;
              if (cookie == "") {
                  // 如果为空行，则提前结束读取
                  break;
              }
              if (exec == "是") {
                  execHandle(cookie, i);
              }
          }   
          message = messageMerge()// 将消息数组融合为一条总消息
          push(message); // 推送消息
      }else{
          for (let i = 2; i <= line; i++) {
              var cookie = Application.Range("A" + i).Text;
              var exec = Application.Range("B" + i).Text;
              if (cookie == "") {
                  // 如果为空行，则提前结束读取
                  break;
              }
              if (exec == "是") {
                  console.log("🧑 开始执行用户：" + "1" )
                  execHandle(cookie, i);
                  break;  // 只取一个
              }
          } 
      }

  }

// }

// 激活工作表函数
function ActivateSheet(sheetName) {
    let flag = 0;
    try {
      // 激活工作表
      let sheet = Application.Sheets.Item(sheetName);
      sheet.Activate();
      console.log("🥚 激活工作表：" + sheet.Name);
      flag = 1;
    } catch {
      flag = 0;
      console.log("🍳 无法激活工作表，工作表可能不存在");
    }
    return flag;
}

// 对推送数据进行处理
function jsonPushHandle(pushName, pushFlag, pushKey) {
  let length = jsonPush.length;
  for (let i = 0; i < length; i++) {
    if (jsonPush[i].name == pushName) {
      if (pushFlag == "是") {
        jsonPush[i].flag = 1;
        jsonPush[i].key = pushKey;
      }
    }
  }
}

// 将消息数组融合为一条总消息
function messageMerge(){
    // console.log(messageArray)
    let message = ""
  for(i=0; i<messageArray.length; i++){
    if(messageArray[i] != "" && messageArray[i] != null)
    {
      message += "\n" + messageHeader[i] + messageArray[i] + ""; // 加上推送头
    }
  }
  if(message != "")
  {
    console.log("✨✨✨✨✨✨✨✨✨✨✨✨✨✨✨✨✨✨")
    console.log(message + "\n")  // 打印总消息
    console.log("✨✨✨✨✨✨✨✨✨✨✨✨✨✨✨✨✨✨")
  }
  return message
}

function sleep(d) {
  for (var t = Date.now(); Date.now() - t <= d; );
}

// 获取sign，返回小写
function getsign(data) {
    var sign = Crypto.createHash("md5")
        .update(data, "utf8")
        .digest("hex")
        // .toUpperCase() // 大写
        .toString();
    return sign;
}

// =================共用结束===================

// cookie字符串转json格式
function cookie_to_json(cookies) {
  var cookie_text = cookies;
  var arr = [];
  var text_to_split = cookie_text.split(";");
  for (var i in text_to_split) {
    var tmp = text_to_split[i].split("=");
    arr.push('"' + tmp.shift().trim() + '":"' + tmp.join(":").trim() + '"');
  }
  var res = "{\n" + arr.join(",\n") + "\n}";
  return JSON.parse(res);
}

// 获取10 位时间戳
function getts10() {
  var ts = Math.round(new Date().getTime() / 1000).toString();
  return ts;
}

// 获取13位时间戳
function getts13(){
  // var ts = Math.round(new Date().getTime()/1000).toString()  // 获取10 位时间戳
  let ts = new Date().getTime()
  return ts
}

// 符合UUID v4规范的随机字符串 b9ab98bb-b8a9-4a8a-a88a-9aab899a88b9
function generateUUID() {
  return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
    var r = Math.random() * 16 | 0,
        v = c === 'x' ? r : (r & 0x3 | 0x8);
    return v.toString(16);
  });
}

function getUUIDDigits(length) {
    var uuid = generateUUID();
    return uuid.replace(/-/g, '').substr(16, length);
}

// 抽奖
function lottery(url, headers, data, count){
  messageSuccess = ""
  messageFail = ""

  resp = HTTP.post(
    url,
    JSON.stringify(data),
    { headers: headers }
  );
  
  if (resp.status == 200 ) {
    resp = resp.json();
    console.log(resp)
    code = resp["code"]
    
    if(code == 0)
    {
      // 抽奖
      // {"code":0,"msg":"成功","toast":{},"data":{"leftTime":2,"totalTime":3,"participateTimes":1,"data":{"prizeId":15,"prizeName":"谢谢参与","picture":{},"prizeType":7},"points":1,"goldBean":{},"prizeNumber":0},"serverTime":"1700000000000"}
      // {"code":100006,"msg":"抽奖机会不足","toast":"抽奖机会不足","data":{},"serverTime":"1700000000000"}
      prizeName = resp["data"]["data"]["prizeName"]
  
      content = "🌈 " + "第" + count + "抽奖：" + prizeName + "\n"
      messageSuccess += content;
      console.log(content)
    }else
    {
      respmsg = resp["msg"]
      content = "📢 " + "第" + count + "抽奖"+ respmsg + "\n"
      messageFail += content;
      console.log(content);
    }
  } else {
    content = "📢 " + "第" + count + "抽奖失败\n"
    messageFail += content;
    console.log(content);
  }

  msg = [messageSuccess, messageFail]
  return msg

}

// 具体的执行函数
function execHandle(cookie, pos) {
  let messageSuccess = "";
  let messageFail = "";
  let messageName = "";

  // 推送昵称或单元格，还是不推送位置标识
  if (messageNickname == 1) {
    // 推送昵称或单元格
    messageName = Application.Range("C" + pos).Text;
    if(messageName == "")
    {
      messageName = "单元格A" + pos + "";
    }
  }

  posLabel = pos-2 ;  // 存放下标，从0开始
  messageHeader[posLabel] = "👨‍🚀 " + messageName
  // try {
    var url1 = "https://bbs.vivo.com.cn/api/community/signIn/lotteryList"; // 抽奖列表
    var url2 = "https://bbs.vivo.com.cn/api/community/signIn/querySignInfo"; // 签到
    var url3 = "https://bbs.vivo.com.cn/api/community/signIn/signInLottery";  // 抽奖
    lotteryNum = 3 ;  // 抽奖次数，默认3次

    headers = {
      "Host":"bbs.vivo.com.cn",
      "Cookie": cookie,
      "Content-Type": "application/json;charset=utf-8",
      "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/46.0.2486.0 Safari/537.36 Edge/13.10586"
    };

    // 签到
    data = {
      "signInId":1
    }

    resp = HTTP.post(
      url2,
      JSON.stringify(data),
      { headers: headers }
    );
    
    // Application.Range("A4").Value = resp.text()
    // {"code":0,"msg":"成功","toast":null,"data":{"code":0,"msg":"","serverTime":1714974081249,"signIn":{"signInActivity":{"id":1,"activityName":"签到抽奖活动","signInType":null,"signInFlag":null,"period":14,"startTime":1578844800000,"signInRule":"","signInRuleFullPath":"","lotteryId":1,"lotteryActivityId":292,"lotteryActivityUuid":"30000000-0000-0000-0000-f1cda8700000","lotteryActivityName":"签到抽奖","signInIndex":3},"prizeList":[{"signInIndex":1,"pointPackageId":12,"giveLotteryId":1,"lotteryTimes":3,"surprisedLotteryId":null,"surprisedLotteryActivityUuid":null,"surprisedLotteryActivityName":null,"surprisedLotteryActivityId":null,"points":1,"extraLotteryTimes":0},{"signInIndex":2,"pointPackageId":12,"giveLotteryId":1,"lotteryTimes":3,"surprisedLotteryId":null,"surprisedLotteryActivityUuid":null,"surprisedLotteryActivityName":null,"surprisedLotteryActivityId":null,"points":1,"extraLotteryTimes":0},{"signInIndex":3,"pointPackageId":461,"giveLotteryId":1,"lotteryTimes":4,"surprisedLotteryId":null,"surprisedLotteryActivityUuid":null,"surprisedLotteryActivityName":"累计签到3天","surprisedLotteryActivityId":null,"points":3,"extraLotteryTimes":1},{"signInIndex":4,"pointPackageId":12,"giveLotteryId":1,"lotteryTimes":3,"surprisedLotteryId":null,"surprisedLotteryActivityUuid":null,"surprisedLotteryActivityName":null,"surprisedLotteryActivityId":null,"points":1,"extraLotteryTimes":0},{"signInIndex":5,"pointPackageId":12,"giveLotteryId":1,"lotteryTimes":3,"surprisedLotteryId":null,"surprisedLotteryActivityUuid":null,"surprisedLotteryActivityName":null,"surprisedLotteryActivityId":null,"points":1,"extraLotteryTimes":0},{"signInIndex":6,"pointPackageId":12,"giveLotteryId":1,"lotteryTimes":3,"surprisedLotteryId":null,"surprisedLotteryActivityUuid":null,"surprisedLotteryActivityName":null,"surprisedLotteryActivityId":null,"points":1,"extraLotteryTimes":0},{"signInIndex":7,"pointPackageId":459,"giveLotteryId":1,"lotteryTimes":5,"surprisedLotteryId":null,"surprisedLotteryActivityUuid":null,"surprisedLotteryActivityName":"累计签到7天","surprisedLotteryActivityId":null,"points":7,"extraLotteryTimes":2},{"signInIndex":8,"pointPackageId":12,"giveLotteryId":1,"lotteryTimes":3,"surprisedLotteryId":null,"surprisedLotteryActivityUuid":null,"surprisedLotteryActivityName":null,"surprisedLotteryActivityId":null,"points":1,"extraLotteryTimes":0},{"signInIndex":9,"pointPackageId":12,"giveLotteryId":1,"lotteryTimes":3,"surprisedLotteryId":null,"surprisedLotteryActivityUuid":null,"surprisedLotteryActivityName":null,"surprisedLotteryActivityId":null,"points":1,"extraLotteryTimes":0},{"signInIndex":10,"pointPackageId":462,"giveLotteryId":1,"lotteryTimes":6,"surprisedLotteryId":null,"surprisedLotteryActivityUuid":null,"surprisedLotteryActivityName":"累计签到10天","surprisedLotteryActivityId":null,"points":10,"extraLotteryTimes":3},{"signInIndex":11,"pointPackageId":12,"giveLotteryId":1,"lotteryTimes":3,"surprisedLotteryId":null,"surprisedLotteryActivityUuid":null,"surprisedLotteryActivityName":null,"surprisedLotteryActivityId":null,"points":1,"extraLotteryTimes":0},{"signInIndex":12,"pointPackageId":12,"giveLotteryId":1,"lotteryTimes":3,"surprisedLotteryId":null,"surprisedLotteryActivityUuid":null,"surprisedLotteryActivityName":null,"surprisedLotteryActivityId":null,"points":1,"extraLotteryTimes":0},{"signInIndex":13,"pointPackageId":12,"giveLotteryId":1,"lotteryTimes":3,"surprisedLotteryId":null,"surprisedLotteryActivityUuid":null,"surprisedLotteryActivityName":null,"surprisedLotteryActivityId":null,"points":1,"extraLotteryTimes":0},{"signInIndex":14,"pointPackageId":460,"giveLotteryId":1,"lotteryTimes":7,"surprisedLotteryId":null,"surprisedLotteryActivityUuid":null,"surprisedLotteryActivityName":"累计签到14天","surprisedLotteryActivityId":null,"points":14,"extraLotteryTimes":4}],"signInResult":{"code":0,"msg":"","signInIndex":3,"signInDay":20240506,"createTime":1714969980000,"todaySignInFlag":1,"nowSignInFlag":0,"schedulingFlag":8,"prizes":null}},"lottery":{"lotteryActivity":{"id":1,"name":"签到抽奖","lotteryTimes":2},"prizeList":[{"id":null,"prizeNumber":null,"givenNumber":null,"chance":null,"winTimes":null,"totalNumber":null,"limitTime":null,"prizeType":1,"picture":"","prizeData":null,"position":"1","prizeName":"TWS 4 耳机","isDefault":0,"remainNumber":null,"prizeId":1115},{"id":null,"prizeNumber":null,"givenNumber":null,"chance":null,"winTimes":null,"totalNumber":null,"limitTime":null,"prizeType":4,"picture":"","prizeData":null,"position":"2","prizeName":"4积分","isDefault":0,"remainNumber":null,"prizeId":23},{"id":null,"prizeNumber":null,"givenNumber":null,"chance":null,"winTimes":null,"totalNumber":null,"limitTime":null,"prizeType":1,"picture":"","prizeData":null,"position":"3","prizeName":"2A USB数据线","isDefault":0,"remainNumber":null,"prizeId":712},{"id":null,"prizeNumber":null,"givenNumber":null,"chance":null,"winTimes":null,"totalNumber":null,"limitTime":null,"prizeType":1,"picture":"","prizeData":null,"position":"4","prizeName":"大容量运动水壶","isDefault":0,"remainNumber":null,"prizeId":1051},{"id":null,"prizeNumber":null,"givenNumber":null,"chance":null,"winTimes":null,"totalNumber":null,"limitTime":null,"prizeType":1,"picture":"","prizeData":null,"position":"5","prizeName":"优酷会员周卡【电子】","isDefault":0,"remainNumber":null,"prizeId":652},{"id":null,"prizeNumber":null,"givenNumber":null,"chance":null,"winTimes":null,"totalNumber":null,"limitTime":null,"prizeType":4,"picture":"","prizeData":null,"position":"6","prizeName":"1积分","isDefault":0,"remainNumber":null,"prizeId":21},{"id":null,"prizeNumber":null,"givenNumber":null,"chance":null,"winTimes":null,"totalNumber":null,"limitTime":null,"prizeType":1,"picture":"","prizeData":null,"position":"7","prizeName":"网易云音乐月卡-电子","isDefault":0,"remainNumber":null,"prizeId":545},{"id":null,"prizeNumber":null,"givenNumber":null,"chance":null,"winTimes":null,"totalNumber":null,"limitTime":null,"prizeType":7,"picture":"","prizeData":null,"position":"8","prizeName":"谢谢参与","isDefault":0,"remainNumber":null,"prizeId":15}]},"memberAssets":{"points":209,"goldBean":7,"prizeNumber":0}},"serverTime":"1710000000000"}

    if (resp.status == 200) {
      resp = resp.json();
      console.log(resp)
      code = resp["code"]
      
      if(code == 0)
      {
        content = "🎉 " + "签到成功 "
        messageSuccess += content;
        console.log(content);
      }else
      {
        content = "❌ " + "签到失败 "
        messageFail += content;
        console.log(content);
      }
    } else {
      content = "❌ " + "签到失败 "
      messageFail += content;
      console.log(content);
    }

    // 抽奖列表
    // {"code":0,"msg":"成功","toast":{},"data":[{"nickName":"雪*孤泣","packageName":"积分定制雨伞"},{"nickName":"vi***********35","packageName":"积分定制雨伞"},{"nickName":"高*友","packageName":"积分定制雨伞"},{"nickName":"周*","packageName":"积分定制雨伞"},{"nickName":"vi**********03","packageName":"积分定制雨伞"},{"nickName":"vi*****88","packageName":"4积分"},{"nickName":"不*知恩","packageName":"1积分"},{"nickName":"vi***********73","packageName":"1积分"},{"nickName":"幸福*禾木","packageName":"4积分"},{"nickName":"vi***********73","packageName":"4积分"}],"serverTime":"1714973762971"}

    // 抽奖
    data = {
      "lotteryActivityId":1,
      "lotteryType":0
    }
    for(i=0; i<lotteryNum; i++)
    {
      console.log("🍳 第" + (i+1) + "次抽奖")
      msg = lottery(url3, headers, data, (i+1))
      messageSuccess += msg[0]
      messageFail += msg[1]
    }

  // } catch {
  //   messageFail += messageName + "失败";
  // }

  // console.log(messageSuccess)
  sleep(2000);
  if (messageOnlyError == 1) {
      messageArray[posLabel] = messageFail;
  } else {
      if(messageFail != ""){
          messageArray[posLabel] = messageFail + " " + messageSuccess;
      }else{
          messageArray[posLabel] = messageSuccess;
      }
  }

  if(messageArray[posLabel] != "")
  {
      console.log(messageArray[posLabel]);
  }
}
