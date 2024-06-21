// WPS权益报名和打卡、超级会员打卡(打卡版)
// 20240621

let sheetNameSubConfig = "wps"; // 分配置表名称
let sheetNameSubConfig2 = "wps_daka";
let pushHeader = "【wps打卡版】";
let sheetNameConfig = "CONFIG"; // 总配置表
let sheetNamePush = "PUSH"; // 推送表名称
let sheetNameEmail = "EMAIL"; // 邮箱表
let flagSubConfig = 0; // 激活分配置工作表标志
let flagConfig = 0; // 激活主配置工作表标志
let flagPush = 0; // 激活推送工作表标志
let line = 21; // 指定读取从第2行到第line行的内容
var message = ""; // 待发送的消息
var messageArray = [];  // 待发送的消息数据，每个元素都是某个账号的消息。目的是将不同用户消息分离，方便个性化消息配置
var messageOnlyError = 0; // 0为只推送失败消息，1则为推送成功消息。
var messageNickname = 0; // 1为推送位置标识（昵称/单元格Ax（昵称为空时）），0为不推送位置标识
var messageHeader = []; // 存放每个消息的头部，如：单元格A3。目的是分离附加消息和执行结果消息
var messagePushHeader = pushHeader; // 存放在总消息的头部，默认是pushHeader,如：【xxxx】

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
// 总推送
function push(message) {
  if (message != "") {
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
    console.log("🍳 消息为空不推送");
  }
}

// 推送bark消息
function bark(message, key) {
    if (key != "") {
      message = messagePushHeader + message // 消息头最前方默认存放：【xxxx】
      message = encodeURIComponent(message)
      BARK_ICON = "https://s21.ax1x.com/2024/06/21/pkDYtK0.png"
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

// 打卡渠道2
function daka2(cookie, Signature){
  messageSuccess = ""
  messageFail = ""
  msg = []

  // 查询获得的奖励
  // url = "https://personal-bus.wps.cn/activity/clock_in/v1/info?client_type=1&page_index=0&page_size=2"
  // 签到
  url = "https://personal-bus.wps.cn/activity/clock_in/v1/clock_in"
  headers = {
    "Host": "personal-bus.wps.cn",
    "Content-Type": "application/json",
    "Cookie": "csrf=1234567890;wps_sid=" + cookie,
    "sid": cookie,
    "Date": "Wed, 15 May 2024 02:20:22 GMT",
    "Signature": Signature,
    "X-CSRFToken": 1234567890,
  }

  data = {
    "client_type":1
  }

  // {"result":"ok","msg":"","data":{"reward_list":{"list":[],"total_num":0},"clock_in_total_num":,"continuous_days":0,"s_key":""}}
  // {"result":"ok","msg":"","data":{"reward_list":{"list":[{"reward_id":5990777,"reward_status":1,"clock_in_time":1715998599,"reward_type":2,"sku_name":"图片权益包1天","mb_name":"","mb_id":0,"mb_img_url":""},{"reward_id":5897293,"reward_status":3,"clock_in_time":1715825063,"reward_type":4,"sku_name":"","mb_name":"蓝色简约大气商务模板","mb_id":,"mb_img_url":""}],"total_num":3},"clock_in_total_num":18040065,"continuous_days":1,"s_key":""}}
  // resp = HTTP.fetch(url, {
  //   method: "post",
  //   headers: headers,
  //   data : data,
  // });

  // {"result":"error","msg":"already clocked in today","data":{}}
  resp = HTTP.post(url,
    data = data,
    {headers : headers}
  )

  resp = resp.json();
  console.log(resp);
  result = resp["result"]
  continuous_days = resp["data"]["continuous_days"]
  if(result == "ok")
  {
    
    // clock_in_status = resp["data"]["clock_in_status"]

    // right = resp["data"]["reward_list"]["list"][0]["sku_name"]
    // if(right == "" || right == "undefined")
    // {
    //   right = "打卡成功"
    // }

    content = "打卡渠道2：周签" + continuous_days + "天，需手动领取奖励 "
    messageSuccess += content;
    console.log(content);
  }else
  {
    msg = resp["msg"]
    content = "打卡渠道2：" + msg + " "
    // messageFail += content;
    messageSuccess += content;
    console.log(content);
  }

  sleep(2000); 
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
    url1 = "https://docs.wps.cn/2c/kdocsclock/api/v1/clock/handle"; // 打卡
    url2 = "https://docs.wps.cn/2c/kdocsclock/api/v1/clock/attend"; // 报名
    content = ""

    headers = {
      Cookie: "wps_sid=" + cookie,
      "User-Agent":
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/46.0.2486.0 Safari/537.36 Edge/13.10586",
    };
    data = {};

    // 渠道1是否打卡。渠道1打卡，此渠道自动领取奖励
    flagExec1 = Application.Range("E" + pos).Text;
    if(flagExec1 == '是')
    {
       console.log("🍳 进行渠道1打卡，此渠道自动领取奖励")
      // 打卡
      // {"code":0,"msg":"ok","data":{"equity":"1天PDF权益包即将到账","right":"1天PDF权益包","writer":"即将到账!"},"request_id":""}
      // {"code":20002,"msg":"打卡失败","request_id":""}
      resp = HTTP.fetch(url1, {
        method: "post",
        headers: headers,
        data: data,
      });

      resp = resp.json();
      console.log(resp);
      code = resp["code"]
      if(code == 0)
      {
        right = resp["data"]["right"]
        content = "🎉 " + "打卡渠道1：" + right + "\n"
        messageSuccess += content;
        console.log(content);
      }else
      {
        respmsg = resp["msg"]
        content = "📢 " + "打卡渠道1：" + respmsg + "\n"
        messageFail += content;
        console.log(content);
      }

      sleep(2000); 
    
      // 报名
      // {"code":0,"msg":"ok","data":{"subscribe":true},"request_id":""}
      // {"code":20001,"msg":"报名失败","request_id":""}
      resp = HTTP.fetch(url2, {
        method: "post",
        headers: headers,
        data: data,
      });

      resp = resp.json();
      console.log(resp);
      code = resp["code"]
      if(code == 0)
      {
        respmsg = resp["msg"]
        if(respmsg == "ok"){
          respmsg = "报名成功"
          content = "🎉 " + "渠道1报名情况：" + respmsg + "\n"
        }else{
          content = "📢 " + "渠道1报名情况：" + respmsg + "\n"
        }
        
        messageSuccess += content;
        console.log(content);
      }else
      {
        respmsg = resp["msg"]
        content = "📢 " + "渠道1报名情况：" + respmsg + "\n"
        messageFail += content;
        console.log(content);
      }

    }else
    {
       console.log("🍳 不进行渠道1打卡")
    }
    
    sleep(2000);

    // 渠道2是否打卡。渠道1打卡，此渠道需手动领取奖励
    flagExec2 = Application.Range("F" + pos).Text;
    if(flagExec2 == '是')
    {
      // 打卡渠道2
       console.log("🍳 进行渠道2打卡，此渠道需手动领取奖励")
      Signature = Application.Range("G" + pos).Text;
      msg = daka2(cookie, Signature)
      messageSuccess += msg[0]
      messageFail += msg[1]
    }else{
       console.log("🍳 不进行渠道2打卡")
    }
    
  // } catch {
  //   messageFail += messageName + "失败";
  // }

  
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