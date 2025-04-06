/*
    name: "百度贴吧"
    cron: 45 0 9 * * *
    脚本兼容: 金山文档（1.0）
    更新时间：20241226
    环境变量名：无
    环境变量值：无
    备注：需要BDUSS。
          cookie填写百度贴吧网页版中获取的BDUSS。F12 -> Application(中文名叫"应用程序") -> Cookie -> BDUSS
          百度贴吧网址：https://tieba.baidu.com/
*/

var sheetNameSubConfig = "tieba"; // 分配置表名称
var pushHeader = "【百度贴吧】";
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
function getts() {
  var ts = Math.round(new Date().getTime() / 1000).toString();
  return ts;
}

function execHandle(cookie,pos,_0x23e9d,_0x0e0a,_0x66783a){var _0xa66ddd=(247421^247423)+(155695^155686);_0x23e9d="";_0xa66ddd=(247913^247919)+(377216^377221);_0x0e0a="";var _0xg_0xbbe=(634953^634945)+(497949^497940);_0x66783a="";_0xg_0xbbe=(214260^214259)+(721991^721985);if(messageNickname==(657563^657562)){_0x66783a=Application['\u0052\u0061\u006E\u0067\u0065']("\u0043"+pos)['\u0054\u0065\u0078\u0074'];if(_0x66783a==""){_0x66783a="A\u683C\u5143\u5355".split("").reverse().join("")+pos+"";}}posLabel=pos-(604132^604134);messageHeader[posLabel]="\uD83D\uDC68\u200D\uD83D\uDE80\u0020"+_0x66783a;try{cookie_json=cookie_to_json(cookie);try{BDUSS=cookie_json["\u0042\u0044\u0055\u0053\u0053"];if(BDUSS!=""&&BDUSS!="denifednu".split("").reverse().join("")&&BDUSS!=undefined){cookie=BDUSS;console['\u006C\u006F\u0067']("\uD83C\uDF73\u0020\u8BFB\u53D6\u5230\u7684\u0063\u006F\u006F\u006B\u0069\u0065\u4E3A\u539F\u59CB\u0063\u006B\uFF0C\u63D0\u53D6\u5176\u4E2D\u7684\u0042\u0044\u0055\u0053\u0053");}}catch{console['\u006C\u006F\u0067']("\u8D25\u5931\u5BFB\u641CSSUDB \uDF73\uD83C".split("").reverse().join(""));}resp=HTTP['\u0067\u0065\u0074']("sbt/nommoc/cd/moc.udiab.abeit//:ptth".split("").reverse().join(""),{'\u0068\u0065\u0061\u0064\u0065\u0072\u0073':{'\u0043\u006F\u006F\u006B\u0069\u0065':"\u0042\u0044\u0055\u0053\u0053\u003D"+cookie}});resp=resp['\u006A\u0073\u006F\u006E']();var _0x4bd6a=(455950^455943)+(175953^175953);var _0xf55d6c=resp["tbs"];_0x4bd6a='\u0068\u0066\u0068\u006E\u0063\u0064';sleep(297674^297250);var _0x7e9bcd=(366564^366560)+(895189^895196);var _0x86c8cc=getts();_0x7e9bcd=572281^572283;var _0x62dfca;var _0xa3c9a=getsign(cookie,_0x86c8cc);_0x62dfca="dhinff".split("").reverse().join("");var _0xac08be;var _0xb8aa=getdata(cookie,_0x86c8cc,_0xa3c9a);_0xac08be=(554594^554592)+(696553^696544);let _0x_0xeda=HTTP['\u0070\u006F\u0073\u0074']("ekil/murof/f/c/moc.udiab.abeit.c//:ptth".split("").reverse().join(""),_0xb8aa);_0x_0xeda=_0x_0xeda['\u006A\u0073\u006F\u006E']();sleep(346664^346560);var _0x6g88b=(177715^177714)+(415501^415493);var _0x370b5a=_0x_0xeda["forum_list"]["non-gconforum"];_0x6g88b=893710^893705;for(var j=842741^842741;j<_0x370b5a['\u006C\u0065\u006E\u0067\u0074\u0068'];j++){client_sign(cookie,_0xf55d6c,_0x370b5a[j]["\u0069\u0064"],_0x370b5a[j]["name"]);content=" \uDF89\uD83C".split("").reverse().join("")+_0x370b5a[j]["\u006E\u0061\u006D\u0065"]+"\n\u5230\u7B7E".split("").reverse().join("");_0x23e9d+=content;console['\u006C\u006F\u0067'](content);sleep(688543^708543);}}catch{_0x0e0a+=" \u274C".split("").reverse().join("")+"\u5931\u8D25\u000A";}sleep(746494^746542);if(messageOnlyError==(677085^677084)){messageArray[posLabel]=_0x0e0a;}else{if(_0x0e0a!=""){messageArray[posLabel]=_0x0e0a+"\u0020"+_0x23e9d;}else{messageArray[posLabel]=_0x23e9d;}}if(messageArray[posLabel]!=""){console['\u006C\u006F\u0067'](messageArray[posLabel]);}}
function getsign(cookie,ts){var _0xb_0x138;var _0x41b93g="=SSUDB".split("").reverse().join("")+cookie+"\u005F\u0063\u006C\u0069\u0065\u006E\u0074\u005F\u0069\u0064\u003D\u0077\u0061\u0070\u0070\u0063\u005F\u0031\u0035\u0033\u0034\u0032\u0033\u0035\u0034\u0039\u0038\u0032\u0039\u0031\u005F\u0034\u0038\u0038\u005F\u0063\u006C\u0069\u0065\u006E\u0074\u005F\u0074\u0079\u0070\u0065\u003D\u0032\u005F\u0063\u006C\u0069\u0065\u006E\u0074\u005F\u0076\u0065\u0072\u0073\u0069\u006F\u006E\u003D\u0039\u002E\u0037\u002E\u0038\u002E\u0030\u005F\u0070\u0068\u006F\u006E\u0065\u005F\u0069\u006D\u0065\u0069\u003D\u0030\u0030\u0030\u0030\u0030\u0030\u0030\u0030\u0030\u0030\u0030\u0030\u0030\u0030\u0030\u0066\u0072\u006F\u006D\u003D\u0031\u0030\u0030\u0038\u0036\u0032\u0031\u0079\u006D\u006F\u0064\u0065\u006C\u003D\u004D\u0049\u002B\u0035\u006E\u0065\u0074\u005F\u0074\u0079\u0070\u0065\u003D\u0031\u0070\u0061\u0067\u0065\u005F\u006E\u006F\u003D\u0031\u0070\u0061\u0067\u0065\u005F\u0073\u0069\u007A\u0065\u003D\u0032\u0030\u0030\u0074\u0069\u006D\u0065\u0073\u0074\u0061\u006D\u0070\u003D"+ts+"\u0076\u0063\u006F\u0064\u0065\u005F\u0074\u0061\u0067\u003D\u0031\u0031";_0xb_0x138=565787^565791;var _0x6b_0xfe3="\u0074\u0069\u0065\u0062\u0061\u0063\u006C\u0069\u0065\u006E\u0074\u0021\u0021\u0021";_0x41b93g=_0x41b93g+_0x6b_0xfe3;var _0xb717f=(697295^697286)+(839732^839729);var _0x594bcb=Crypto['\u0063\u0072\u0065\u0061\u0074\u0065\u0048\u0061\u0073\u0068']("\u006D\u0064\u0035")['\u0075\u0070\u0064\u0061\u0074\u0065'](_0x41b93g,"8ftu".split("").reverse().join(""))['\u0064\u0069\u0067\u0065\u0073\u0074']("xeh".split("").reverse().join(""))['\u0074\u006F\u0055\u0070\u0070\u0065\u0072\u0043\u0061\u0073\u0065']()['\u0074\u006F\u0053\u0074\u0072\u0069\u006E\u0067']();_0xb717f="qbdhlh".split("").reverse().join("");return _0x594bcb;}
function getdata(cookie,ts,sign){var _0x14f1a=(992011^992014)+(650602^650606);var _0xa37d="\u0042\u0044\u0055\u0053\u0053\u003D"+cookie+"\u0026\u005F\u0063\u006C\u0069\u0065\u006E\u0074\u005F\u0074\u0079\u0070\u0065\u003D\u0032\u0026\u005F\u0063\u006C\u0069\u0065\u006E\u0074\u005F\u0069\u0064\u003D\u0077\u0061\u0070\u0070\u0063\u005F\u0031\u0035\u0033\u0034\u0032\u0033\u0035\u0034\u0039\u0038\u0032\u0039\u0031\u005F\u0034\u0038\u0038\u0026\u005F\u0063\u006C\u0069\u0065\u006E\u0074\u005F\u0076\u0065\u0072\u0073\u0069\u006F\u006E\u003D\u0039\u002E\u0037\u002E\u0038\u002E\u0030\u0026\u005F\u0070\u0068\u006F\u006E\u0065\u005F\u0069\u006D\u0065\u0069\u003D\u0030\u0030\u0030\u0030\u0030\u0030\u0030\u0030\u0030\u0030\u0030\u0030\u0030\u0030\u0030\u0026\u0066\u0072\u006F\u006D\u003D\u0031\u0030\u0030\u0038\u0036\u0032\u0031\u0079\u0026\u0070\u0061\u0067\u0065\u005F\u006E\u006F\u003D\u0031\u0026\u0070\u0061\u0067\u0065\u005F\u0073\u0069\u007A\u0065\u003D\u0032\u0030\u0030\u0026\u006D\u006F\u0064\u0065\u006C\u003D\u004D\u0049\u0025\u0032\u0042\u0035\u0026\u006E\u0065\u0074\u005F\u0074\u0079\u0070\u0065\u003D\u0031\u0026\u0074\u0069\u006D\u0065\u0073\u0074\u0061\u006D\u0070\u003D"+ts+"=ngis&11=gat_edocv&".split("").reverse().join("")+sign;_0x14f1a="dehdpg".split("").reverse().join("");return _0xa37d;}
function getsign2(cookie,ts,tbs,fid,kw){var _0x94b5c=(610950^610945)+(918591^918588);var _0xaf84d="=SSUDB".split("").reverse().join("")+cookie+"=dif000000000000000=iemi_enohp_0.8.7.9=noisrev_tneilc_2=epyt_tneilc_".split("").reverse().join("")+fid+"=wk".split("").reverse().join("")+kw+"\u006D\u006F\u0064\u0065\u006C\u003D\u004D\u0049\u002B\u0035\u006E\u0065\u0074\u005F\u0074\u0079\u0070\u0065\u003D\u0031\u0074\u0062\u0073\u003D"+tbs+"\u0074\u0069\u006D\u0065\u0073\u0074\u0061\u006D\u0070\u003D"+ts;_0x94b5c=143680^143689;var _0x6b15d=(840051^840055)+(202340^202340);var _0x7g_0x677="\u0074\u0069\u0065\u0062\u0061\u0063\u006C\u0069\u0065\u006E\u0074\u0021\u0021\u0021";_0x6b15d="llnfgo".split("").reverse().join("");_0xaf84d=_0xaf84d+_0x7g_0x677;var _0xg6b=Crypto['\u0063\u0072\u0065\u0061\u0074\u0065\u0048\u0061\u0073\u0068']("5dm".split("").reverse().join(""))['\u0075\u0070\u0064\u0061\u0074\u0065'](_0xaf84d)['\u0064\u0069\u0067\u0065\u0073\u0074']("\u0068\u0065\u0078")['\u0074\u006F\u0055\u0070\u0070\u0065\u0072\u0043\u0061\u0073\u0065']()['\u0074\u006F\u0053\u0074\u0072\u0069\u006E\u0067']();return _0xg6b;}
function getsigndata(cookie,tbs,fid,kw){var _0xd1552c=(316172^316164)+(973859^973867);var _0x297gf=getts();_0xd1552c="flnfdi".split("").reverse().join("");var _0xa0e3c=getsign2(cookie,_0x297gf,tbs,fid,kw);var _0x37c="\u005F\u0063\u006C\u0069\u0065\u006E\u0074\u005F\u0074\u0079\u0070\u0065\u003D\u0032\u0026\u005F\u0063\u006C\u0069\u0065\u006E\u0074\u005F\u0076\u0065\u0072\u0073\u0069\u006F\u006E\u003D\u0039\u002E\u0037\u002E\u0038\u002E\u0030\u0026\u005F\u0070\u0068\u006F\u006E\u0065\u005F\u0069\u006D\u0065\u0069\u003D\u0030\u0030\u0030\u0030\u0030\u0030\u0030\u0030\u0030\u0030\u0030\u0030\u0030\u0030\u0030\u0026\u006D\u006F\u0064\u0065\u006C\u003D\u004D\u0049\u0025\u0032\u0042\u0035\u0026\u006E\u0065\u0074\u005F\u0074\u0079\u0070\u0065\u003D\u0031\u0026\u0042\u0044\u0055\u0053\u0053\u003D"+cookie+"\u0026\u0066\u0069\u0064\u003D"+fid+"\u0026\u006B\u0077\u003D"+kw+"=sbt&".split("").reverse().join("")+tbs+"=pmatsemit&".split("").reverse().join("")+_0x297gf+"=ngis&".split("").reverse().join("")+_0xa0e3c;return _0x37c;}
function client_sign(cookie,tbs,fid,kw){var _0x9aa6a;var _0xg4f4c=getsigndata(cookie,tbs,fid,kw);_0x9aa6a=(278960^278967)+(993908^993909);let _0x_0x517=HTTP['\u0070\u006F\u0073\u0074']("\u0068\u0074\u0074\u0070\u003A\u002F\u002F\u0063\u002E\u0074\u0069\u0065\u0062\u0061\u002E\u0062\u0061\u0069\u0064\u0075\u002E\u0063\u006F\u006D\u002F\u0063\u002F\u0063\u002F\u0066\u006F\u0072\u0075\u006D\u002F\u0073\u0069\u0067\u006E",_0xg4f4c);}
