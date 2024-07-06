// PUSH.js 推送脚本
// 20240706

let sheetNameConfig = "CONFIG"; // 总配置表
let sheetNamePush = "PUSH"; // 推送表名称
let sheetNameEmail = "EMAIL"; // 邮箱表
let flagSubConfig = 0; // 激活分配置工作表标志
let flagConfig = 0; // 激活主配置工作表标志
let flagPush = 0; // 激活推送工作表标志
let line = 21; // 指定读取从第2行到第line行的内容
var message = ""; // 待发送的消息
var messagePushHeader = ""; // 存放在总消息的头部，默认是pushHeader,如：【xxxx】
var pushHeader = ""

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
}; // 

// 获取时间
function getDate(){
  let currentDate = new Date();
  // 2024/07/04
  // currentDate = currentDate.getFullYear() + '' + (currentDate.getMonth() + 1).toString().padStart(2, '0') + '' + currentDate.getDate().toString().padStart(2, '0');
  currentDate = currentDate.getFullYear() + '/' + (currentDate.getMonth() + 1).toString() + '/' + currentDate.getDate().toString();
  
  return currentDate
}

// 当天时间
var todayDate = getDate()
getPush()   // 读取推送配置
var msgArray = [] // 存放消息内容
getMessage()  // 读取消息配置
sendNotify()  // 消息推送
// console.log(jsonPush)
// console.log(jsonEmail)

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
      }else{  // 不推送
        jsonPush[i].flag = 0;
        jsonPush[i].key = pushKey;
      }
    }
  }
}

// 读取推送配置
function getPush(){
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
}

// 休眠
function sleep(d) {
  for (var t = Date.now(); Date.now() - t <= d; );
}

// 读取消息配置
function getMessage(){
  flagConfig = ActivateSheet(sheetNameConfig); // 激活主配置表
  // 主配置工作表存在
  if (flagConfig == 1) {
    console.log("🍳 开始读取主配置表");
    for (let i = 2; i <= 100; i++) {
      // 从工作表中读取推送数据
      let msgDict = {
        "name": "",       // 名称
        "note": "",   // 备注
        // "onlyError":  "", // 只推送错误消息
        "update":"",       // 脚本更新时间，即脚本是否已执行
        "msg" : "",       // 待推送消息
        "date": "",       // 推送时间，即单天是否已推送
        "methodPush":"",  // 推送方式
        "flagPush" : "",  // 是否通知
        "pool":"",        // 是否加入消息池，加入消息池的都会整合为一条消息统一推送
      }

      msgDict["name"] = Application.Range("A" + i).Text;     // 工作表名称
      msgDict.note = Application.Range("B" + i).Text;     // 备注
      // msgDict.onlyError = Application.Range("C" + i).Text;     // 只推送错误消息
      msgDict.update = Application.Range("F" + i).Text;     // 脚本更新时间，即脚本是否已执行
      msgDict.msg = Application.Range("G" + i).Text;     // 待推送消息
      msgDict.date = Application.Range("H" + i).Text;     // 推送时间，即单天是否已推送
      msgDict.methodPush = Application.Range("I" + i).Text;     // 推送方式
      msgDict.flagPush = Application.Range("J" + i).Text;     // 是否通知
      msgDict.pool = Application.Range("K" + i).Text;     // 是否加入消息池，加入消息池的都会整合为一条消息统一推送
    
      if (msgDict.name == "") {
        // 如果为空行，则提前结束读取
        break; // 提前退出，提高效率
      }
      // console.log(msgDict)
      msgArray.push(msgDict)

      

    }
    // console.log(msgArray)


    
  }
}

// 发送消息
function sendNotify(){
  ActivateSheet(sheetNameConfig); // 激活主配置表

  // console.log("🍳 开始发送消息");
  let msgCurrentDict = ""
  let msgPool = ""
  for (let i = 0; i < msgArray.length; i++) {
    msgCurrentDict = msgArray[i]
    // console.log(msgCurrentDict)
    // {"name":"aliyundrive_multiuser","note":"阿里云盘（多用户版）","msg":"","date":"","methodPush":"","flagPush":"@all"}
    // 从读取推送数据
    // if(msgCurrentDict.flagPush == "是" && msgCurrentDict.update != "" && msgCurrentDict.date == ""){  // 第一次执行时更新时间不为空，推送时间为空
    // }

    // console.log(msgCurrentDict.date)
    // console.log(todayDate)
    // 消息池的先不推送，最后统一推送
    if(msgCurrentDict.pool == "否" && msgCurrentDict.flagPush == "是" && msgCurrentDict.update != msgCurrentDict.date && msgCurrentDict.msg != "" && msgCurrentDict.date != todayDate){ // 时间不一致说明未推送。消息为空不进行推送。今天未推送
      console.log("🚀 消息推送：" + msgCurrentDict.note)
      pushMessage(msgCurrentDict.msg, msgCurrentDict.methodPush, "【" + msgCurrentDict.note + "】",)

      // 写入推送的时间
      Application.Range("H" + (i + 2)).Value = todayDate

    }else{
      if(msgCurrentDict.pool == "是" && msgCurrentDict.flagPush == "是" && msgCurrentDict.update != msgCurrentDict.date && msgCurrentDict.msg != "" && msgCurrentDict.date != todayDate){
        // console.log("🧩 加入消息池：" + msgCurrentDict.note)
        msgPool += "【" + msgCurrentDict.note + "】" + msgCurrentDict.msg + "\n"

        // 写入推送的时间
        Application.Range("H" + (i + 2)).Value = todayDate

      }else{
        // console.log("🍳 不进行推送：" + msgCurrentDict.note)
      }
    }
  }
  
  // console.log(msgPool)
  // 消息池推送，消息池默认以@all方式推送
  let msgPoolJuice = msgPool.replace(/\n/g, '');  // 判断消息池内是否有数据
  // console.log(msgPoolJuice)
  if(msgPoolJuice != ""){ // 消息池内有消息才推送
    console.log("🚀 艾默库消息池推送")
    pushMessage(msgPool, "@all", "【" + "艾默库消息池" + "】")
  }

  console.log("🎉 推送结束")
}

// 消息分割，返回消息推送方式数组
function pushSplit(method){
  // console.log(method)
  let arry = []
  arry = method.split("&") // 使用&作为分隔符
  // console.log(arry)
  return arry
}

// 总推送
function pushMessage(message, method, pushHeader){
  messagePushHeader = pushHeader
  if (method == "@all") { // 所有渠道都推送
    // console.log("🚀 所有渠道都推送");
    // message = messagePushHeader + message // 消息头最前方默认存放：【xxxx】
    let length = jsonPush.length;
    let name;
    let key;
    for (let i = 0; i < length; i++) {
      if (jsonPush[i].flag == 1) {
        name = jsonPush[i].name;
        key = jsonPush[i].key;
        if (name == "bark") {
          bark(message, key);
        } else if (name == "pushplus") {
          pushplus(message, key);
        } else if (name == "ServerChan") {
          serverchan(message, key);
        } else if (name == "email") {
          email(message);
        } else if (name == "dingtalk") {
          dingtalk(message, key);
        } else if (name == "discord") {
          discord(message, key);
        }
      }
    }
  } else {
    // console.log("🚀 多消息推送");
    let arry = pushSplit(method)
    let methodCurrent = ""

    let length = jsonPush.length;
    let name;
    let key;

    for (let i = 0; i < arry.length; i++) {
      methodCurrent = arry[i]
      // console.log(methodCurrent)
      for (let i = 0; i < length; i++) {
        name = jsonPush[i].name;
        if(name == methodCurrent){
          // console.log(methodCurrent)
          if (jsonPush[i].flag == 1) {
            key = jsonPush[i].key;
            if (name == "bark") {
              bark(message, key);
            } else if (name == "pushplus") {
              pushplus(message, key);
            } else if (name == "ServerChan") {
              serverchan(message, key);
            } else if (name == "email") {
              email(message);
            } else if (name == "dingtalk") {
              dingtalk(message, key);
            } else if (name == "discord") {
              discord(message, key);
            }
          }
          break;  // 找到推送方式就提前退出
        }
      } 
    }
  }
}

// 推送bark消息
function bark(message, key) {
    if (key != "") {
      message = messagePushHeader + message // 消息头最前方默认存放：【xxxx】
      message = encodeURIComponent(message)
      // console.log(message)
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
    url = "http://www.pushplus.plus/send?token=" + key + "&content=" + message + "&title=" + messagePushHeader;  // 增加标题
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
  // console.log("🍳 开始读取邮箱配置");
  let length = jsonPush.length; // 因为此json数据可无序，因此需要遍历
  let name;
  for (let i = 0; i < length; i++) {
    name = jsonPush[i].name;
    if (name == "email") {
      if (jsonPush[i].flag == 1 || 1) { // 始终读取
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






