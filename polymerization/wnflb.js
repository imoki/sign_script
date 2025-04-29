/*
    name: "万能福利吧"
    cron: 45 0 9 * * *
    脚本兼容: 金山文档（1.0），金山文档（2.0）
    更新时间：20241226
    环境变量名：无
    环境变量值：无
    备注：需要cookie。F12 -> "NetWork"(中文名叫"网络") -> 按一下Ctrl+R -> www.wnflb2023.com -> cookie
          有门槛，没账号的不用折腾。
          万能福利吧网址：https://www.wnflb2023.com
*/

var sheetNameSubConfig = "wnflb"; // 分配置表名称
var pushHeader = "【万能福利吧】";
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
var separator = "##########MOKU##########" // 分割符，分割消息。可用于PUSH.js灵活推送
var maxMessageLength = 400;  // 设置最大长度，超过这个长度则分片发送
var messageDistance = 100; // 消息距离，用于匹配100字符内最近的行

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
    var url1 = "https://www.wnflb2023.com/plugin.php?id=fx_checkin:list"; // 获取formhash、判断成功+获取积分
    var url2 = "https://www.wnflb2023.com/plugin.php?id=fx_checkin%3Acheckin&infloat=yes&handlekey=fx_checkin&inajax=1&ajaxtarget=fwin_content_fx_checkin" // 签到

    headers={
      "Host": "www.wnflb2023.com", 
      // "Content-Type": "application/x-www-form-urlencoded",
      "Cookie":cookie,
      // "Cookie":""
    }

    // 获取formhash
    resp = HTTP.get(
      url1,
      { headers: headers }
    );

    // 正则匹配
    formhash = ""
    // const Reg = /你已经连续签到(.*?)天，再接再厉！/i;
    Reg = [
      /formhash=(.+?)&/i, 
      /showmenu">积分: (.+?)<\/a>/i,
    ]

    valueName = [
      "formhash", "签到前积分",
    ]

    html = resp.text();
    // console.log(html)
    for(i=0; i< Reg.length; i++)
    {
      flagTrue = Reg[i].test(html); // 判断是否存在字符串
      if (flagTrue == true) {
        let result = Reg[i].exec(html); // 提取匹配的字符串，["你已经连续签到 1 天，再接再厉！"," 1 "]
        // result = result[0];
        result = result[1];
        if(i == 1){
          content = "🎉 " +  valueName[i] + ":" + result + " "
          messageSuccess += content;
        }else
        {
          formhash = result
          content = "🍳 formhash:" + result + " "
        }
        console.log(content)
      } else {
        content = "❌ " + "formhash获取失败 "
        messageFail += content;
      }
    }

    // 签到
    headers={
      "Host": "www.wnflb2023.com", 
      "Accept-Encoding": "gzip, deflate, br",
      "Accept-Language": "zh-CN,zh;q=0.9",
      // "Content-Type": "application/x-www-form-urlencoded",
      "X-Requested-With":"XMLHttpRequest",
      "Cookie":cookie,
      "Referer":"https://www.wnflb2023.com/",
      "DNT":1,
    }
    params = "&formhash=" + formhash + "&" + formhash
    url2 = url2 + params
    // console.log(url2)

    resp = HTTP.get(
      url2,
      { headers: headers }
    );
    
    // console.log(resp.text())

//     <?xml version="1.0" encoding="utf-8"?>
// <root><![CDATA[<h3 class="flb"><em>提示信息</em><span><a href="javascript:;" class="flbc" onclick="hideWindow('fx_checkin');" title="关闭">关闭</a></span></h3>
// <div class="c altw">
// <div class="alert_right"><script type="text/javascript" reload="1">if(typeof errorhandle_fx_checkin=='function') {errorhandle_fx_checkin('签到成功,您今日第3874个签到,累计签到2851天!', {});}hideWindow('fx_checkin');showDialog('签到成功,您今日第3874个签到,累计签到2851天!', 'right', null, null, 0, null, null, null, null, null, null);</script><script type="text/javascript">fx_chk_menu=true;$('fx_checkin_topb').innerHTML="<a href=\"plugin.php?id=fx_checkin:list\" onmouseover=\"fx_checkin_menu('fx_checkin_topb');\"><img id=\"fx_checkin_b\" src=\"source/plugin/fx_checkin/images/mini2.gif\"  style=\"position:relative;top:5px;height:18px;\"></a>";$('fx_checkin_menut').innerHTML="<em>签到成功!</em><p>您今天第<i>3874</i>个签到，签到排名竞争激烈，记得每天都来签到哦！</p>";$('fx_checkin_menub').innerHTML="已连续签到:<i>8</i>天，累计签到:<i>2851</i>天";</script></div>
// </div>
// <p class="o pns">
// <button type="button" class="pn pnc" id="closebtn" onclick="hideWindow('fx_checkin');"><strong>确定</strong></button>
// <script type="text/javascript" reload="1">if($('closebtn')) {$('closebtn').focus();}</script>
// </p>
// ]]></root>

// <?xml version="1.0" encoding="utf-8"?>
// <root><![CDATA[<h3 class="flb"><em>提示信息</em><span><a href="javascript:;" class="flbc" onclick="hideWindow('fx_checkin');" title="关闭">关闭</a></span></h3>
// <div class="c altw">
// <div class="alert_right"><script type="text/javascript" reload="1">if(typeof errorhandle_fx_checkin=='function') {errorhandle_fx_checkin('签名出错-2,请重新登陆后签到1!', {});}hideWindow('fx_checkin');showDialog('签名出错-2,请重新登陆后签到1!', 'right', null, null, 0, null, null, null, null, null, null);</script></div>
// </div>
// <p class="o pns">
// <button type="button" class="pn pnc" id="closebtn" onclick="hideWindow('fx_checkin');"><strong>确定</strong></button>
// <script type="text/javascript" reload="1">if($('closebtn')) {$('closebtn').focus();}</script>
// </p>
// ]]></root>

    // 获取签到天数数据、获取积分
    headers={
      "Host": "www.wnflb2023.com", 
      // "Content-Type": "application/x-www-form-urlencoded",
      "Cookie":cookie,
      // "Cookie":""
    }

    resp = HTTP.get(
      url1,
      { headers: headers }
    );
    // console.log(resp.text())
    // 正则匹配
    // const Reg = /你已经连续签到(.*?)天，再接再厉！/i;
    Reg = [
      /累计签到:<i>(.+?)<\/i>天/i, 
      /已连续签到:<i>(.+?)<\/i>天/i,
      /showmenu">积分: (.+?)<\/a>/i,
    ]

    valueName = [
      "累计签天数", "已连签天数","当前积分",
    ]

    html = resp.text();
    // console.log(html)
    resultall = ""
    for(i=0; i< Reg.length; i++)
    {
      flagTrue = Reg[i].test(html); // 判断是否存在字符串
      if (flagTrue == true) {
        let result = Reg[i].exec(html); // 提取匹配的字符串，["你已经连续签到 1 天，再接再厉！"," 1 "]
        // result = result[0];
        result = result[1];
        if(result == "{days}" || result == "{constant}" ){

        }else{
          content = "🎉 " + valueName[i] + ":" + result + " "
          messageSuccess += content;
          console.log(content)
        }
        
      } else {
        content = "❌ " +"签到数据获取失败 "
        messageFail += content;
      }
    }

    // // 获取积分
    // // 正则匹配,获取formhash
    // formhash = ""
    // // const Reg = /你已经连续签到(.*?)天，再接再厉！/i;
    // Reg = [
    //   /formhash=(.+?)&/i, 
    //   /showmenu">积分: (.+?)<\/a>/i,
      
    // ]

    // valueName = [
    //   "formhash", "当前积分",
    // ]

    // // html = resp.text();
    // // console.log(html)
    // for(i=0; i< Reg.length; i++)
    // {
    //   flagTrue = Reg[i].test(html); // 判断是否存在字符串
    //   if (flagTrue == true) {
    //     let result = Reg[i].exec(html); // 提取匹配的字符串，["你已经连续签到 1 天，再接再厉！"," 1 "]
    //     // result = result[0];
    //     result = result[1];
    //     formhash = result
    //     if(i == 1){
    //       content = valueName[i] + ":" + result + " "
    //       messageSuccess += content;
    //     }else
    //     {
    //       content = "formhash:" + result + " "
    //     }
        
    //     console.log(content)
    //   } else {
    //     content = "formhash获取失败 "
    //     messageFail += content;
    //   }
    // }

  // } catch {
  //   messageFail += messageName + "失败";
  // }

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
