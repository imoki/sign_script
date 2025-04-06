/*
    name: "vivo社区"
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

function lottery(url,headers,data,count){messageSuccess="";messageFail="";resp=HTTP['\u0070\u006F\u0073\u0074'](url,JSON['\u0073\u0074\u0072\u0069\u006E\u0067\u0069\u0066\u0079'](data),{'\u0068\u0065\u0061\u0064\u0065\u0072\u0073':headers});if(resp['\u0073\u0074\u0061\u0074\u0075\u0073']==(194721^194665)){resp=resp['\u006A\u0073\u006F\u006E']();code=resp["code"];if(code==(666571^666571)){prizeName=resp["\u0064\u0061\u0074\u0061"]["data"]["prizeName"];content="\uD83C\uDF08\u0020"+"\u7B2C"+count+"\u62BD\u5956\uFF1A"+prizeName+"\u000A";messageSuccess+=content;console['\u006C\u006F\u0067'](content);}else{respmsg=resp["\u006D\u0073\u0067"];content=" \uDCE2\uD83D".split("").reverse().join("")+"\u7B2C"+count+"\u62BD\u5956"+respmsg+"\u000A";messageFail+=content;console['\u006C\u006F\u0067'](content);}}else{content=" \uDCE2\uD83D".split("").reverse().join("")+"\u7B2C"+count+"\u62BD\u5956\u5931\u8D25\u000A";messageFail+=content;console['\u006C\u006F\u0067'](content);}msg=[messageSuccess,messageFail];return msg;}
function execHandle(cookie,pos,_0x2dfg,_0xdc9d8a,_0xbaf5ae){var _0x0180ga=(751269^751277)+(118745^118746);_0x2dfg="";_0x0180ga='\u0066\u0065\u0069\u006A\u006A\u0062';_0xdc9d8a="";_0xbaf5ae="";if(messageNickname==(335776^335777)){_0xbaf5ae=Application['\u0052\u0061\u006E\u0067\u0065']("\u0043"+pos)['\u0054\u0065\u0078\u0074'];if(_0xbaf5ae==""){_0xbaf5ae="\u5355\u5143\u683C\u0041"+pos+"";}}posLabel=pos-(660894^660892);messageHeader[posLabel]=" \uDE80\uD83D\u200D\uDC68\uD83D".split("").reverse().join("")+_0xbaf5ae;var _0x6d9fd="tsiLyrettol/nIngis/ytinummoc/ipa/nc.moc.oviv.sbb//:sptth".split("").reverse().join("");var _0x19a="\u0068\u0074\u0074\u0070\u0073\u003A\u002F\u002F\u0062\u0062\u0073\u002E\u0076\u0069\u0076\u006F\u002E\u0063\u006F\u006D\u002E\u0063\u006E\u002F\u0061\u0070\u0069\u002F\u0063\u006F\u006D\u006D\u0075\u006E\u0069\u0074\u0079\u002F\u0073\u0069\u0067\u006E\u0049\u006E\u002F\u0071\u0075\u0065\u0072\u0079\u0053\u0069\u0067\u006E\u0049\u006E\u0066\u006F";var _0x8fe=(158797^158792)+(899890^899899);var _0xf886cf="\u0068\u0074\u0074\u0070\u0073\u003A\u002F\u002F\u0062\u0062\u0073\u002E\u0076\u0069\u0076\u006F\u002E\u0063\u006F\u006D\u002E\u0063\u006E\u002F\u0061\u0070\u0069\u002F\u0063\u006F\u006D\u006D\u0075\u006E\u0069\u0074\u0079\u002F\u0073\u0069\u0067\u006E\u0049\u006E\u002F\u0073\u0069\u0067\u006E\u0049\u006E\u004C\u006F\u0074\u0074\u0065\u0072\u0079";_0x8fe="ajmhbe".split("").reverse().join("");lotteryNum=800031^800028;headers={"\u0048\u006F\u0073\u0074":"\u0062\u0062\u0073\u002E\u0076\u0069\u0076\u006F\u002E\u0063\u006F\u006D\u002E\u0063\u006E","\u0043\u006F\u006F\u006B\u0069\u0065":cookie,"Content-Type":"\u0061\u0070\u0070\u006C\u0069\u0063\u0061\u0074\u0069\u006F\u006E\u002F\u006A\u0073\u006F\u006E\u003B\u0063\u0068\u0061\u0072\u0073\u0065\u0074\u003D\u0075\u0074\u0066\u002D\u0038","User-Agent":"\u004D\u006F\u007A\u0069\u006C\u006C\u0061\u002F\u0035\u002E\u0030\u0020\u0028\u0057\u0069\u006E\u0064\u006F\u0077\u0073\u0020\u004E\u0054\u0020\u0031\u0030\u002E\u0030\u003B\u0020\u0057\u0069\u006E\u0036\u0034\u003B\u0020\u0078\u0036\u0034\u0029\u0020\u0041\u0070\u0070\u006C\u0065\u0057\u0065\u0062\u004B\u0069\u0074\u002F\u0035\u0033\u0037\u002E\u0033\u0036\u0020\u0028\u004B\u0048\u0054\u004D\u004C\u002C\u0020\u006C\u0069\u006B\u0065\u0020\u0047\u0065\u0063\u006B\u006F\u0029\u0020\u0043\u0068\u0072\u006F\u006D\u0065\u002F\u0034\u0036\u002E\u0030\u002E\u0032\u0034\u0038\u0036\u002E\u0030\u0020\u0053\u0061\u0066\u0061\u0072\u0069\u002F\u0035\u0033\u0037\u002E\u0033\u0036\u0020\u0045\u0064\u0067\u0065\u002F\u0031\u0033\u002E\u0031\u0030\u0035\u0038\u0036"};data={"signInId":1};resp=HTTP['\u0070\u006F\u0073\u0074'](_0x19a,JSON['\u0073\u0074\u0072\u0069\u006E\u0067\u0069\u0066\u0079'](data),{"headers":headers});if(resp['\u0073\u0074\u0061\u0074\u0075\u0073']==(716349^716533)){resp=resp['\u006A\u0073\u006F\u006E']();code=resp["\u0063\u006F\u0064\u0065"];if(code==(246515^246515)){content="\uD83C\uDF89\u0020"+" \u529F\u6210\u5230\u7B7E".split("").reverse().join("");_0x2dfg+=content;console['\u006C\u006F\u0067'](content);}else{content=" \u274C".split("").reverse().join("")+"\u7B7E\u5230\u5931\u8D25\u0020";_0xdc9d8a+=content;console['\u006C\u006F\u0067'](content);}}else{content="\u274C\u0020"+"\u7B7E\u5230\u5931\u8D25\u0020";_0xdc9d8a+=content;console['\u006C\u006F\u0067'](content);}data={"\u006C\u006F\u0074\u0074\u0065\u0072\u0079\u0041\u0063\u0074\u0069\u0076\u0069\u0074\u0079\u0049\u0064":1,"\u006C\u006F\u0074\u0074\u0065\u0072\u0079\u0054\u0079\u0070\u0065":0};for(i=830706^830706;i<lotteryNum;i++){console['\u006C\u006F\u0067']("\uD83C\uDF73\u0020\u7B2C"+(i+(563073^563072))+"\u6B21\u62BD\u5956");msg=lottery(_0xf886cf,headers,data,i+(922456^922457));_0x2dfg+=msg[227906^227906];_0xdc9d8a+=msg[110596^110597];}sleep(608555^610043);if(messageOnlyError==(140754^140755)){messageArray[posLabel]=_0xdc9d8a;}else{if(_0xdc9d8a!=""){messageArray[posLabel]=_0xdc9d8a+"\u0020"+_0x2dfg;}else{messageArray[posLabel]=_0x2dfg;}}if(messageArray[posLabel]!=""){console['\u006C\u006F\u0067'](messageArray[posLabel]);}}
