/*
    name: "熊猫代理"
    cron: 45 0 9 * * *
    脚本兼容: 金山文档
    更新时间：20241106
    环境变量名：xmdl
    环境变量值：http://www.xiongmaodaili.com?invitationCode=1368A6DA-2960-4070-9F9B-4ABACC8D678D 需要用户名和密码
*/

let sheetNameSubConfig = "xmdl"; // 分配置表名称
let pushHeader = "【熊猫代理】";
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
// 推送相关
// 获取时间
function getDate() {
  let currentDate = new Date();
  currentDate = currentDate.getFullYear() + '/' + (currentDate.getMonth() + 1).toString() + '/' + currentDate.getDate().toString();
  return currentDate
}

// 将消息写入CONFIG表中作为消息队列，之后统一发送
function writeMessageQueue(message) {
  // 当天时间
  let todayDate = getDate()
  flagConfig = ActivateSheet(sheetNameConfig); // 激活主配置表
  // 主配置工作表存在
  if (flagConfig == 1) {
    console.log("✨ 开始将结果写入主配置表");
    for (let i = 2; i <= 100; i++) {
      // 找到指定的表行
      if (Application.Range("A" + (i + 2)).Value2 == sheetNameSubConfig) {
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

  if (qlSwitch != 1) {  // 金山文档
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
  } else {
    for (let i = 2; i <= line; i++) {
      var cookie = Application.Range("A" + i).Text;
      var exec = Application.Range("B" + i).Text;
      if (cookie == "") {
        // 如果为空行，则提前结束读取
        break;
      }
      if (exec == "是") {
        console.log("🧑 开始执行用户：" + "1")
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
function messageMerge() {
  // console.log(messageArray)
  let message = ""
  for (i = 0; i < messageArray.length; i++) {
    if (messageArray[i] != "" && messageArray[i] != null) {
      message += "\n" + messageHeader[i] + messageArray[i] + ""; // 加上推送头
    }
  }
  if (message != "") {
    console.log("✨✨✨✨✨✨✨✨✨✨✨✨✨✨✨✨✨✨")
    console.log(message + "\n")  // 打印总消息
    console.log("✨✨✨✨✨✨✨✨✨✨✨✨✨✨✨✨✨✨")
  }
  return message
}

function sleep(d) {
  for (var t = Date.now(); Date.now() - t <= d;);
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
function getts13() {
  // var ts = Math.round(new Date().getTime()/1000).toString()  // 获取10 位时间戳
  let ts = new Date().getTime()
  return ts
}

// 符合UUID v4规范的随机字符串 b9ab98bb-b8a9-4a8a-a88a-9aab899a88b9
function generateUUID() {
  return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
    var r = Math.random() * 16 | 0,
      v = c === 'x' ? r : (r & 0x3 | 0x8);
    return v.toString(16);
  });
}

function getUUIDDigits(length) {
  var uuid = generateUUID();
  return uuid.replace(/-/g, '').substr(16, length);
}

function login(_0x2fbd38,_0xc0d2c5,_0x42817e){messageSuccess="".split("").reverse().join("");messageFail="".split("").reverse().join("");token="".split("").reverse().join("");resp=HTTP["\u0070\u006f\u0073\u0074"](_0x2fbd38,_0x42817e,{"\u0068\u0065\u0061\u0064\u0065\u0072\u0073":_0xc0d2c5,"\u0076\u0065\u0072\u0069\u0066\u0079":"\u0046\u0061\u006c\u0073\u0065"});console["\u006c\u006f\u0067"](resp);if(resp["\u0073\u0074\u0061\u0074\u0075\u0073"]==(0xc01b3^0xc017b)){respj=resp["\u006a\u0073\u006f\u006e"]();console["\u006c\u006f\u0067"](respj);code=respj["\u0063\u006f\u0064\u0065"];console["\u006c\u006f\u0067"](code);if(code==(0x7a616^0x7a616)){token=resp['msg']['headers']["\u0053\u0065\u0074\u002d\u0043\u006f\u006f\u006b\u0069\u0065"];console["\u006c\u006f\u0067"](token);cookies=token['toString']()["\u0073\u0070\u006c\u0069\u0074"](';');token=cookies[0x3835b^0x3835b];console['log'](token);content='📢\x20'+"\n\u529F\u6210\u5F55\u767B".split("").reverse().join("");messageSuccess+=content;console["\u006c\u006f\u0067"](content);}else{respmsg=respj["\u006d\u0073\u0067"];content='📢\x20'+respmsg+'\x0a';messageFail+=content;console["\u006c\u006f\u0067"](content);}}else{content=" \u274C".split("").reverse().join("")+"\n\u8D25\u5931\u5F55\u767B".split("").reverse().join("");messageFail+=content;console["\u006c\u006f\u0067"](content);}msg=[messageSuccess,messageFail,token];return msg;}function sign(_0x382352,_0xa987ea,_0x1cc39c){messageSuccess="".split("").reverse().join("");messageFail="";flagstatus=0xa3870^0xa3870;headers1={"\u0055\u0073\u0065\u0072\u002d\u0041\u0067\u0065\u006e\u0074":'Mozilla/5.0\x20(Windows\x20NT\x2010.0;\x20Win64;\x20x64)\x20AppleWebKit/537.36\x20(KHTML,\x20like\x20Gecko)\x20Chrome/122.0.6261.95\x20Safari/537.36',"\u0043\u006f\u006e\u0074\u0065\u006e\u0074\u002d\u0054\u0079\u0070\u0065":"\u0061\u0070\u0070\u006c\u0069\u0063\u0061\u0074\u0069\u006f\u006e\u002f\u0078\u002d\u0077\u0077\u0077\u002d\u0066\u006f\u0072\u006d\u002d\u0075\u0072\u006c\u0065\u006e\u0063\u006f\u0064\u0065\u0064",'Cookie':_0x1cc39c};resp=HTTP['get']("emiTyaDnIngiSteg/stniop/bew-oamgnoix/moc.iliadoamgnoix.www//:ptth".split("").reverse().join(""),{'headers':headers1});console['log'](resp);respj=resp["\u006a\u0073\u006f\u006e"]();console['log'](respj);console['log'](respj['msg']);if(respj["\u0063\u006f\u0064\u0065"]!=(0x31ee9^0x31ee9)){content='❌\x20'+"\n\u8D25\u5931\u5230\u7B7E".split("").reverse().join("");messageFail+=content;console['log'](content);msg=[messageSuccess,messageFail,flagstatus];return msg;}resp=HTTP["\u0067\u0065\u0074"](_0x382352+(respj["\u006f\u0062\u006a"]+(0xc9be8^0xc9be9))['toString'](),{"\u0068\u0065\u0061\u0064\u0065\u0072\u0073":headers1});console["\u006c\u006f\u0067"](resp);console['log'](resp['msg']);if(resp['status']==(0x95688^0x95640)){resp=resp['json']();console["\u006c\u006f\u0067"](resp);code=resp["\u0063\u006f\u0064\u0065"];resp1=HTTP["\u0067\u0065\u0074"]("stnioPresUteg/stniop/bew-oamgnoix/moc.iliadoamgnoix.www//:ptth".split("").reverse().join(""),{'headers':headers1});resp1=resp1['json']();console['log'](resp1);console["\u006c\u006f\u0067"](resp1['obj']);if(code==(0xc2cb4^0xc2cb4)){respmsg=resp['msg'];score=resp1['obj'];content='🎉\x20'+respmsg+"\u5F97\u83B7".split("").reverse().join("")+score+'积分\x0a';messageSuccess+=content;console["\u006c\u006f\u0067"](content);flagstatus=0x4b0ee^0x4b0ef;}else{respmsg=resp['msg'];if(respmsg=='今日已签到！'){content='📢\x20'+respmsg+'\x0a';messageSuccess+=content;console["\u006c\u006f\u0067"](content);flagstatus=0x6b164^0x6b165;}else{content='📢\x20'+respmsg+'\x0a';messageFail+=content;console['log'](content);}}}else{content='❌\x20'+"\n\u8D25\u5931\u5230\u7B7E".split("").reverse().join("");messageFail+=content;console['log'](content);}msg=[messageSuccess,messageFail,flagstatus];return msg;}function execHandle(_0xe9bbb2,_0x21c1c2){let _0xe7398f='';let _0x36dd22="";let _0x4ed9bd='';if(messageNickname==(0x24907^0x24906)){_0x4ed9bd=Application["\u0052\u0061\u006e\u0067\u0065"]("\u0043"+_0x21c1c2)["\u0054\u0065\u0078\u0074"];if(_0x4ed9bd==''){_0x4ed9bd="A\u683C\u5143\u5355".split("").reverse().join("")+_0x21c1c2+"".split("").reverse().join("");}}posLabel=_0x21c1c2-0x2;messageHeader[posLabel]='👨‍🚀\x20'+_0x4ed9bd;var _0x19231c="nigol/resu/bew-oamgnoix/moc.iliadoamgnoix.www//:ptth".split("").reverse().join("");var _0x40e07f="=yaDnIngis?stnioPeviecer/stniop/bew-oamgnoix/moc.iliadoamgnoix.www//:ptth".split("").reverse().join("");token=_0xe9bbb2;username=Application['Range']('D'+_0x21c1c2)["\u0054\u0065\u0078\u0074"];password=Application['Range']('E'+_0x21c1c2)['Text'];headers={"\u0055\u0073\u0065\u0072\u002d\u0041\u0067\u0065\u006e\u0074":'Mozilla/5.0\x20(Windows\x20NT\x2010.0;\x20Win64;\x20x64)\x20AppleWebKit/537.36\x20(KHTML,\x20like\x20Gecko)\x20Chrome/122.0.6261.95\x20Safari/537.36','Content-Type':"\u0061\u0070\u0070\u006c\u0069\u0063\u0061\u0074\u0069\u006f\u006e\u002f\u0078\u002d\u0077\u0077\u0077\u002d\u0066\u006f\u0072\u006d\u002d\u0075\u0072\u006c\u0065\u006e\u0063\u006f\u0064\u0065\u0064"};flagstatus=0x5d8b7^0x5d8b7;if(token!=''){data={'SESSION':"\u0033\u0031\u0066\u0034\u0066\u0030\u0033\u0062\u002d\u0037\u0033\u0035\u0032\u002d\u0034\u0065\u0030\u0038\u002d\u0039\u0062\u0061\u0030\u002d\u0037\u0061\u0031\u0036\u0064\u0062\u0038\u0038\u0035\u0061\u0065\u0062"};msg=sign(_0x40e07f,headers,token);if(msg[0xa4252^0xa4250]==0x1){flagstatus=0x1;content=msg[0x0];_0xe7398f+=content;console['log'](content);}else{console['log']('🍳\x20此token签到失败，尝试登录获取新token');}}else{console['log']('🍳\x20token为空，开始进行登录获取token');}if(flagstatus!=(0x44211^0x44210)){data={'account':username,"\u0070\u0061\u0073\u0073\u0077\u006f\u0072\u0064":password,'originType':'1'};msg=login(_0x19231c,headers,data);if(msg[0x2]!=''){token=msg[0xa2fd0^0xa2fd2];console["\u006c\u006f\u0067"]('🍳\x20登录成功，已获得最新token:'+token);Application["\u0052\u0061\u006e\u0067\u0065"]('A'+_0x21c1c2)['Value2']=token;data={'SESSION':token};msg=sign(_0x40e07f,headers,token);if(msg[0x5aebe^0x5aebc]==(0x99c56^0x99c57)){content=msg[0x0];_0xe7398f+=content;console["\u006c\u006f\u0067"](content);}else{content=msg[0x65027^0x65026];_0x36dd22+=content;console["\u006c\u006f\u0067"](content);}}else{console["\u006c\u006f\u0067"]('❌\x20获取最新token为空');content=msg[0x41a6e^0x41a6f];_0x36dd22+=content;console["\u006c\u006f\u0067"](content);}}sleep(0x9ecd0^0x9eb00);if(messageOnlyError==(0xed41a^0xed41b)){messageArray[posLabel]=_0x36dd22;}else{if(_0x36dd22!=''){messageArray[posLabel]=_0x36dd22+'\x20'+_0xe7398f;}else{messageArray[posLabel]=_0xe7398f;}}if(messageArray[posLabel]!=""){console['log'](messageArray[posLabel]);}}