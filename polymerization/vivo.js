// vivo社区签到和抽奖
// 20240506

let sheetNameSubConfig = "vivo"; // 分配置表名称
let pushHeader = "【vivo社区】";
let sheetNameConfig = "CONFIG"; // 总配置表
let sheetNamePush = "PUSH"; // 推送表名称
let sheetNameEmail = "EMAIL"; // 邮箱表
let flagSubConfig = 0; // 激活分配置工作表标志
let flagConfig = 0; // 激活主配置工作表标志
let flagPush = 0; // 激活推送工作表标志
let line = 21; // 指定读取从第2行到第line行的内容
var message = ""; // 待发送的消息
var messageOnlyError = 0; // 0为只推送失败消息，1则为推送成功消息。
var messageNickname = 0; // 1为用昵称替代单元格，0为不替代
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

flagConfig = ActivateSheet(sheetNameConfig); // 激活推送表
// 主配置工作表存在
if (flagConfig == 1) {
  console.log("开始读取主配置表");
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
        console.log("只推送错误消息");
      }

      if (nickname == "是") {
        messageNickname = 1;
        console.log("单元格用昵称替代");
      }

      break; // 提前退出，提高效率
    }
  }
}

flagPush = ActivateSheet(sheetNamePush); // 激活推送表
// 推送工作表存在
if (flagPush == 1) {
  console.log("开始读取推送工作表");
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
  console.log("开始读取分配置表");
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

  push(message); // 推送消息
}

// 总推送
function push(message) {
  if (message != "") {
    message = pushHeader + message; // 加上推送头
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
    console.log("消息为空不推送");
  }
}

// 推送bark消息
function bark(message, key) {
  if (key != "") {
    let url = "https://api.day.app/" + key + "/" + message;
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
      "?title=消息推送" +
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
  // console.log("已发送邮件至：" + sender);
  console.log("已发送邮件");
  sleep(5000);
}

// 邮箱配置
function emailConfig() {
  console.log("开始读取邮箱配置");
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
          console.log("开始读取邮箱表");
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
  let url = "https://oapi.dingtalk.com/robot/send?access_token=" + key;
  let resp = HTTP.post(url, { msgtype: "text", text: { content: message } });
  // console.log(resp.text())
  sleep(5000);
}
// 推送Discord机器人
function discord(message, key) {
  let url = key;
  let resp = HTTP.post(url, { content: message });
  //console.log(resp.text())
  sleep(5000);
}
function sleep(d) {
  for (var t = Date.now(); Date.now() - t <= d; );
}

// 激活工作表函数
function ActivateSheet(sheetName) {
  let flag = 0;
  try {
    // 激活工作表
    let sheet = Application.Sheets.Item(sheetName);
    sheet.Activate();
    console.log("激活工作表：" + sheet.Name);
    flag = 1;
  } catch {
    flag = 0;
    console.log("无法激活工作表，工作表可能不存在");
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
 


// 获取sign，返回小写
function getsign(data) {
  var sign = Crypto.createHash("md5")
    .update(data, "utf8")
    .digest("hex")
    // .toUpperCase() // 大写
    .toString();
  return sign;
}

// 抽奖
function lottery(url, headers, data){
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
  
      content = "抽奖：" + prizeName + " "
      messageSuccess += content;
      console.log(content)
    }else
    {
      respmsg = resp["msg"]
      content = respmsg + "抽奖失败 "
      messageFail += content;
      console.log(content);
    }
  } else {
    content = "抽奖失败 "
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

  if (messageNickname == 1) {
    messageName = Application.Range("C" + pos).Text;
  } else {
    messageName = "单元格A" + pos + "";
  }
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
        content = "签到成功 "
        messageSuccess += content;
        console.log(content);
      }else
      {
        content = "签到失败 "
        messageFail += content;
        console.log(content);
      }
    } else {
      content = "签到失败 "
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
      console.log("第" + (i+1) + "次抽奖")
      msg = lottery(url3, headers, data)
      messageSuccess += msg[0]
      messageFail += msg[1]
    }

  // } catch {
  //   messageFail += messageName + "失败";
  // }

  // console.log(messageSuccess)
  sleep(2000);
  if (messageOnlyError == 1) {
    message += messageFail;
  } else {
    message += messageFail + " " + messageSuccess;
  }

  message = "帐号：" + messageName + message  // 附加账号信息

  console.log(message);
}
