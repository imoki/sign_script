/*
    name: "中兴商城"
    cron: 45 0 9 * * *
    脚本兼容: 金山文档（1.0），金山文档（2.0）
    更新时间：20241226
    环境变量名：无
    环境变量值：无
    备注：需要Cookie。
          抓包工具抓取所需的值或者网页端获取所需的值。浏览器访问网页版中兴社区，F12 -> "NetWork"(中文名叫"网络") -> 按一下Ctrl+R -> /cn -> cookie
          中兴商城网址：https://www.ztemall.com/cn/
*/

var sheetNameSubConfig = "ztemall"; // 分配置表名称
var pushHeader = "【中兴商城】";
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
 


// 获取sign，返回小写
function getsign(data) {
  var sign = Crypto.createHash("md5")
    .update(data, "utf8")
    .digest("hex")
    // .toUpperCase() // 大写
    .toString();
  return sign;
}

function execHandle(cookie,pos,_0x5956b,_0xee06de,_0x00231f){var _0x981a;_0x5956b="";_0x981a=(326479^326473)+(710007^710015);_0xee06de="";var _0x270f=(534665^534669)+(705930^705929);_0x00231f="";_0x270f=(891284^891286)+(864556^864549);if(messageNickname==(776298^776299)){_0x00231f=Application['\u0052\u0061\u006E\u0067\u0065']("\u0043"+pos)['\u0054\u0065\u0078\u0074'];if(_0x00231f==""){_0x00231f="\u5355\u5143\u683C\u0041"+pos+"";}}posLabel=pos-(499418^499416);messageHeader[posLabel]=" \uDE80\uD83D\u200D\uDC68\uD83D".split("").reverse().join("")+_0x00231f+"\u000A";var _0x3e7e="ipapot/php.xedni/moc.llametz.www//:sptth".split("").reverse().join("");try{cookie_json=cookie_to_json(cookie);accessToken=cookie_json["\u0061\u0075\u0074\u0068\u005F\u0074\u006F\u006B\u0065\u006E\u005F\u0070\u0063"];if(accessToken==undefined){accessToken=cookie;console['\u006C\u006F\u0067']("nekoTssecca\u4E3A\u4F5Ceikooc\u5C06\uFF0Ccp_nekot_htua\u65E0\u4E2Deikooc \uDF73\uD83C".split("").reverse().join(""));}}catch{accessToken=cookie;}console['\u006C\u006F\u0067']("\uD83C\uDF73\u0020\u5DF2\u8BFB\u53D6\u5230\u0061\u0063\u0063\u0065\u0073\u0073\u0054\u006F\u006B\u0065\u006E\u003A"+accessToken);params="=nekoTssecca&=ngis&1v=v&nosj=tamrof&dda.nIkcehc.rebmem=dohtem?".split("").reverse().join("")+accessToken;headers={"Host":"\u0077\u0077\u0077\u002E\u007A\u0074\u0065\u006D\u0061\u006C\u006C\u002E\u0063\u006F\u006D"};resp=HTTP['\u0067\u0065\u0074'](_0x3e7e+params,{'\u0068\u0065\u0061\u0064\u0065\u0072\u0073':headers});if(resp['\u0073\u0074\u0061\u0074\u0075\u0073']==(287825^287897)){resp=resp['\u006A\u0073\u006F\u006E']();errorcode=resp["\u0065\u0072\u0072\u006F\u0072\u0063\u006F\u0064\u0065"];if(errorcode==(538208^538208)){msg=resp["msg"];checkin_days=resp["data"]["\u0063\u0068\u0065\u0063\u006B\u0069\u006E\u005F\u0064\u0061\u0079\u0073"];currentCheckInPoint=resp["\u0064\u0061\u0074\u0061"]["\u0063\u0075\u0072\u0072\u0065\u006E\u0074\u0043\u0068\u0065\u0063\u006B\u0049\u006E\u0050\u006F\u0069\u006E\u0074"];point=resp["\u0064\u0061\u0074\u0061"]["\u0070\u006F\u0069\u006E\u0074"];content=" \uDF89\uD83C".split("").reverse().join("")+msg+"\u0020\u8FDE\u7B7E"+checkin_days+"\u5F97\u83B7,\u5929".split("").reverse().join("")+currentCheckInPoint+"\u79EF\u5206\u000A";_0x5956b+=content;console['\u006C\u006F\u0067'](content);}else if(errorcode==(106496^100112)){content="\uD83D\uDCE2\u0020"+"\u53EF\u80FD\u5DF2\u7ECF\u7B7E\u8FC7\u4E86\u000A";_0x5956b+=content;console['\u006C\u006F\u0067'](content);}else{content="\u274C\u0020"+"\u7B7E\u5230\u5931\u8D25\u000A";_0xee06de+=content;console['\u006C\u006F\u0067'](content);}}else{content="\u274C\u0020"+"\u7B7E\u5230\u5931\u8D25\u000A";_0xee06de+=content;console['\u006C\u006F\u0067'](content);}var _0x1a16b;var _0x9da43b="=nekoTssecca&=ngis&1v=v&nosj=tamrof&tsil.ksat=dohtem&5h=mroftalp?ipapot/php.xedni/moc.llametz.www//:sptth".split("").reverse().join("");_0x1a16b=(642197^642193)+(969247^969242);var _0xffc;var _0xc83e="ipapot/php.xedni/moc.llametz.www//:sptth".split("").reverse().join("");_0xffc=(414436^414437)+(254108^254110);console['\u006C\u006F\u0067']("\u8868\u5217di\u52A1\u4EFB\u53D6\u83B7 \uDF73\uD83C".split("").reverse().join(""));resp=HTTP['\u0067\u0065\u0074'](_0x9da43b+params,{'\u0068\u0065\u0061\u0064\u0065\u0072\u0073':headers});task_id_list=[];resp=resp['\u006A\u0073\u006F\u006E']();tasks=resp["data"]["tasks"];for(key=468191^468191;key<tasks['\u006C\u0065\u006E\u0067\u0074\u0068'];key++){desc=tasks[key]["btn"]["desc"];if(desc=="\u4E0E\u53C2\u53BB".split("").reverse().join("")||desc=="\u7ACB\u5373\u9886\u53D6"){dict={"\u0074\u0069\u0074\u006C\u0065":tasks[key]["title"],"\u0074\u0061\u0073\u006B\u005F\u0069\u0064":tasks[key]["task_id"],"page_ids":tasks[key]["\u0074\u0061\u0073\u006B\u005F\u0064\u0061\u0074\u0061"]["page_ids"]};task_id_list['\u0070\u0075\u0073\u0068'](dict);}}data={"\u0070\u0061\u0067\u0065\u005F\u0069\u0064":0,"task_id":"","\u006D\u0065\u0074\u0068\u006F\u0064":"\u0074\u0061\u0073\u006B\u002E\u0073\u0074\u0061\u0072\u0074","format":"json","\u0076":"\u0076\u0031","sign":"","accessToken":accessToken};headers={"\u0048\u006F\u0073\u0074":"\u0077\u0077\u0077\u002E\u007A\u0074\u0065\u006D\u0061\u006C\u006C\u002E\u0063\u006F\u006D","Content-Type":"\u0061\u0070\u0070\u006C\u0069\u0063\u0061\u0074\u0069\u006F\u006E\u002F\u0078\u002D\u0077\u0077\u0077\u002D\u0066\u006F\u0072\u006D\u002D\u0075\u0072\u006C\u0065\u006E\u0063\u006F\u0064\u0065\u0064"};for(i=653450^653450;i<task_id_list['\u006C\u0065\u006E\u0067\u0074\u0068'];i++){title=task_id_list[i]["\u0074\u0069\u0074\u006C\u0065"];task_id=task_id_list[i]["task_id"];data["\u0074\u0061\u0073\u006B\u005F\u0069\u0064"]=task_id;try{page_id=task_id_list[i]["\u0070\u0061\u0067\u0065\u005F\u0069\u0064\u0073"][243754^243754];if(page_id!=undefined&&page_id!=""&&page_id!="denifednu".split("").reverse().join("")){data["\u0070\u0061\u0067\u0065\u005F\u0069\u0064"]=page_id;}else{data["\u0070\u0061\u0067\u0065\u005F\u0069\u0064"]=663558^663558;}}catch{data["\u0070\u0061\u0067\u0065\u005F\u0069\u0064"]=572460^572460;}data["\u006D\u0065\u0074\u0068\u006F\u0064"]="\u0074\u0061\u0073\u006B\u002E\u0073\u0074\u0061\u0072\u0074";console['\u006C\u006F\u0067']("\uD83C\uDF73\u0020\u5F00\u59CB\u505A\u4EFB\u52A1\uFF1A"+title);resp=HTTP['\u0070\u006F\u0073\u0074'](_0xc83e,data,{'\u0068\u0065\u0061\u0064\u0065\u0072\u0073':headers});resp=resp['\u006A\u0073\u006F\u006E']();errorcode=resp["\u0065\u0072\u0072\u006F\u0072\u0063\u006F\u0064\u0065"];if(errorcode!=(715684^715684)){content=" \u274C".split("").reverse().join("")+title+"\n\u52A1\u4EFB\u4E2A\u4E00\u4E0B\u5165\u8FDB\u63A5\u76F4\uFF0C\u59CB\u5F00\u6CD5\u65E0".split("").reverse().join("");_0xee06de+=content;console['\u006C\u006F\u0067'](content);sleep(461729^461937);continue;}sleep(658537^665233);data["\u006D\u0065\u0074\u0068\u006F\u0064"]="\u0074\u0061\u0073\u006B\u002E\u0066\u0069\u006E\u0069\u0073\u0068";resp=HTTP['\u0070\u006F\u0073\u0074'](_0xc83e,data,{"headers":headers});resp=resp['\u006A\u0073\u006F\u006E']();errorcode=resp["errorcode"];if(errorcode!=(535127^535127)){content=" \u274C".split("").reverse().join("")+title+"\u65E0\u6CD5\u5B8C\u6210\uFF0C\u76F4\u63A5\u8FDB\u5165\u4E0B\u4E00\u4E2A\u4EFB\u52A1\u000A";_0xee06de+=content;console['\u006C\u006F\u0067'](content);sleep(906971^905483);continue;}sleep(223773^224717);params="=di_ksat?".split("").reverse().join("")+task_id+"=nekoTssecca&=ngis&1v=v&nosj=tamrof&kcehc.ksat=dohtem&".split("").reverse().join("")+accessToken;resp=HTTP['\u0067\u0065\u0074'](_0xc83e+params,{"headers":headers});resp=resp['\u006A\u0073\u006F\u006E']();if(errorcode!=(716025^716025)){content="\u274C\u0020"+title+"\u65E0\u6CD5\u9886\u53D6\u5956\u52B1\uFF0C\u8BF7\u624B\u52A8\u9886\u53D6\u5956\u52B1\u000A";_0xee06de+=content;console['\u006C\u006F\u0067'](content);sleep(947675^946699);continue;}else{content="\u2728\u0020"+title+"\u5DF2\u5B8C\u6210\u000A";_0xee06de+=content;console['\u006C\u006F\u0067'](content);}sleep(306087^306295);}url5="=nekoTssecca&=ngis&1v=v&nosj=tamrof&liated.tniop.rebmem=dohtem&1=ezis_egap&1=on_egap?ipapot/php.xedni/moc.llametz.www//:sptth".split("").reverse().join("")+accessToken;resp=HTTP['\u0067\u0065\u0074'](url5,{"headers":headers});resp=resp['\u006A\u0073\u006F\u006E']();errorcode=resp["errorcode"];if(errorcode==(175451^175451)){point_count=resp["data"]["\u0070\u006F\u0069\u006E\u0074\u005F\u0074\u006F\u0074\u0061\u006C"]["\u0070\u006F\u0069\u006E\u0074\u005F\u0063\u006F\u0075\u006E\u0074"];content="\uD83C\uDF89\u0020"+":\u5206\u79EF\u524D\u5F53".split("").reverse().join("")+point_count+"\u000A";_0x5956b+=content;console['\u006C\u006F\u0067'](content);sleep(924127^925199);}else{content="\u274C\u0020"+"\n\u5206\u79EF\u524D\u5F53\u5230\u8BE2\u67E5\u6CD5\u65E0".split("").reverse().join("");console['\u006C\u006F\u0067'](content);}sleep(524801^525777);if(messageOnlyError==(916096^916097)){messageArray[posLabel]=_0xee06de;}else{if(_0xee06de!=""){messageArray[posLabel]=_0xee06de+"\u0020"+_0x5956b;}else{messageArray[posLabel]=_0x5956b;}}if(messageArray[posLabel]!=""){console['\u006C\u006F\u0067'](messageArray[posLabel]);}}

