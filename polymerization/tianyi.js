// 天翼云盘签到
// 20240621
// 文中引用代码改编自公众号"Jerry不是猫"
// 获取加密后的手机号和密码请访问：https://as.js.cool/reference/ecloud/

pushData = []
let sheetNameSubConfig = "tianyi"; // 分配置表名称
let pushHeader = "【天翼云盘】";
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

// 引用开始
// 52论坛kittylang https://www.52pojie.cn/thread-1919639-1-1.html

// ../../packages/utils-pure/index.ts
function getXmlElement(xml, tag) {
    if (!xml.match) return "";
    let m = xml.match(`<${tag}>(.*)</${tag}>`);
    return m ? m[1] : "";
}
function createLogger(options) {
    let wrap = (type, ...args)=>{
        if (options && options.pushData) {
            let msg = args.reduce((str, cur)=>`${str} ${cur}`, "").substring(1);
            options.pushData.push({
                msg: msg,
                type: type,
                date: /* @__PURE__ */ new Date()
            });
        }
        // console[type](...args);
    }, info = (...args)=>wrap("info", ...args), error = (...args)=>wrap("error", ...args);
    return {
        info: info,
        error: error,
        fatal: error,
        debug: info,
        start: info,
        success: info,
        fail: info,
        trace: info,
        warn: (...args)=>wrap("warn", ...args)
    };
}
function toLowerCaseHeaders(headers) {
    return headers ? Object.entries(headers).reduce((acc, [key, value])=>(acc[key.toLowerCase()] = value, acc), {}) : {};
}
function isPlainObject(obj) {
    return Array.isArray(obj) || Object.prototype.toString.call(obj) === "[object Object]";
}
function sortStringify(obj) {
    return Object.entries(obj).sort().map(([key, value])=>`${key}=${value}`).join("&");
}
function getBeijingTime(timestamp) {
    let time = timestamp ? new Date(timestamp) : /* @__PURE__ */ new Date(), [year, monthStr, dayStr] = time.toLocaleDateString("zh-CN").split(/[/:-]/);
    return {
        year: parseInt(year, 10),
        month: parseInt(monthStr, 10),
        day: parseInt(dayStr, 10)
    };
}
function getInToday(timestamp) {
    let today = getBeijingTime(), time = getBeijingTime(timestamp);
    return today.day === time.day && today.month === time.month && today.year === time.year;
}
// ../../core/ecloud/constant.ts
var Appkey = "600100422";
// ../../core/ecloud/api.ts
function createApi(http) {
    let logboxUrl = "https://open.e.189.cn/api/logbox", oauth2Url = logboxUrl + "/oauth2";
    return {
        loginSubmit: function({ username, password, returnUrl, paramId, pre = "{NRP}" }) {
            return http.post(oauth2Url + "/loginSubmit.do", {
                version: "v2.0",
                appKey: "cloud",
                apToken: "",
                accountType: "01",
                userName: `${pre}${username}`,
                epd: `${pre}${password}`,
                validateCode: "",
                captchaToken: "",
                captchaType: "",
                dynamicCheck: "FALSE",
                clientType: 1,
                returnUrl: returnUrl,
                mailSuffix: "@189.cn",
                paramId: paramId,
                cb_SaveName: 3
            });
        },
        loginRedirect: function() {
            let loginUrl = http.get("https://cloud.189.cn/api/portal/loginUrl.action", {
                native: !0,
                followRedirect: !1
            }).headers.location;
            if (!loginUrl) throw new Error("\u83B7\u53D6\u767B\u5F55\u94FE\u63A5loginUrl\u5931\u8D25");
            let htmlUrl = http.get(loginUrl, {
                native: !0,
                followRedirect: !1
            }).headers.location;
            if (!htmlUrl) throw new Error("\u83B7\u53D6\u767B\u5F55\u94FE\u63A5htmlUrl\u5931\u8D25");
            return htmlUrl.split(".html?")[1].split("&").reduce((acc, cur)=>{
                let [key, value] = cur.split("=");
                return acc[key] = value, acc;
            }, {});
        },
        appConf: function() {
            return http.post(oauth2Url + "/appConf.do", {
                appKey: "cloud"
            });
        },
        encryptConf: function() {
            return http.post(logboxUrl + "/config/encryptConf.do", "appId=cloud");
        },
        loginCallback: function(toUrl) {
            return http.get(toUrl);
        },
        drawPrizeMarket: function(taskId) {
            return http.get(`https://m.cloud.189.cn/v2/drawPrizeMarketDetails.action?taskId=${taskId}&activityId=ACT_SIGNIN&noCache=${Math.random()}`);
        },
        userSign: function(version = "10.1.4", model = "iPhone14") {
            return http.get(`https://api.cloud.189.cn/mkt/userSign.action?rand=${/* @__PURE__ */ new Date().getTime()}&clientType=TELEANDROID&version=${version}&model=${model}`, {
                headers: {
                    HOST: "m.cloud.189.cn"
                }
            });
        },
        getUserBriefInfo: function() {
            return http.get("https://cloud.189.cn/api/portal/v2/getUserBriefInfo.action");
        },
        getAccessTokenBySsKey: function(sessionKey, headers) {
            return http.get(`https://cloud.189.cn/api/open/oauth2/getAccessTokenBySsKey.action?sessionKey=${sessionKey}`, {
                headers: headers
            });
        },
        getFamilyList: function(headers) {
            return http.get("https://api.cloud.189.cn/open/family/manage/getFamilyList.action", {
                headers: headers
            });
        },
        familySign: function(familyId, headers) {
            return http.get(`https://api.cloud.189.cn/open/family/manage/exeFamilyUserSign.action?familyId=${familyId}`, {
                headers: headers
            });
        },
        getListGrow: function() {
            return http.get(`https://cloud.189.cn/api/portal/listGrow.action?noCache=${Math.random()}`);
        },
        http: http
    };
}
// ../../core/ecloud/index.ts
function signIn($) {
    try {
        let { isSign, netdiskBonus } = $.api.userSign();
        if (!netdiskBonus) {
            $.logger.error("\u7B7E\u5230\u5931\u8D25\uFF0C\u8BF7\u624B\u52A8\u7B7E\u5230");
            return;
        }
        if (isSign) {
            $.logger.info("\u4ECA\u65E5\u5DF2\u7B7E\u5230\uFF0C\u83B7\u5F97", netdiskBonus, "MB");
            return;
        }
        $.logger.info("\u7B7E\u5230\u6210\u529F\uFF0C\u83B7\u5F97", netdiskBonus, "MB");
    } catch (error) {
        $.logger.error("\u7B7E\u5230\u5F02\u5E38", error.message);
    }
}
function drawPrizeMarket($, taskId) {
    try {
        let data = $.api.drawPrizeMarket(taskId), { errorCode, prizeName, errorMsg } = data;
        if (errorCode) return errorCode === "User_Not_Chance" ? $.logger.info(`${taskId}\u65E0\u62BD\u5956\u6B21\u6570`) : $.logger.info("\u62BD\u5956\u5931\u8D25", errorCode, errorMsg);
        if (prizeName) return $.logger.info("\u62BD\u5956\u6210\u529F\uFF0C\u83B7\u5F97", prizeName);
        $.logger.info("\u62BD\u5956\u5931\u8D25", JSON.stringify(data));
    } catch (error) {
        $.logger.error("\u62BD\u5956\u5F02\u5E38", error.message);
    }
}
function getEncryptConf($) {
    try {
        let encryptConf = $.api.encryptConf();
        if (encryptConf.data.pubKey) return encryptConf.data;
    } catch (error) {
        $.logger.debug("\u83B7\u53D6\u52A0\u5BC6\u914D\u7F6E\u5F02\u5E38", error);
    }
    return {
        pubKey: "MIGfMA0GCSqGSIb3DQEBAQUAA4GNADCBiQKBgQCZLyV4gHNDUGJMZoOcYauxmNEsKrc0TlLeBEVVIIQNzG4WqjimceOj5R9ETwDeeSN3yejAKLGHgx83lyy2wBjvnbfm/nLObyWwQD/09CmpZdxoFYCH6rdDjRpwZOZ2nXSZpgkZXoOBkfNXNxnN74aXtho2dqBynTw3NFTWyQl8BQIDAQAB",
        pre: "{NRP}"
    };
}
function login($) {
    let { lt, reqId } = $.api.loginRedirect();
    $.api.http.setHeader("Lt", lt).setHeader("reqId", reqId);
    let appConfig = $.api.appConf(), { pre, pubKey } = getEncryptConf($), publicKey = `-----BEGIN PUBLIC KEY-----
${pubKey}
-----END PUBLIC KEY-----`, username = $.config.username, password = $.config.password;
    $.rsaEncrypt && ({ username, password } = $.rsaEncrypt(publicKey, $.config.username, $.config.password));
    let { result, toUrl, msg } = $.api.loginSubmit({
        username: username,
        password: password,
        paramId: appConfig.data.paramId,
        returnUrl: appConfig.data.returnUrl,
        pre: pre
    });
    if (result !== 0) throw $.logger.error("\u767B\u5F55\u5931\u8D25", result, msg), new Error(msg);
    $.api.loginCallback(toUrl);
}
function drawPrizeTask($) {
    $.logger.debug("\u5F00\u59CB\u6267\u884C\u62BD\u5956\u4EFB\u52A1");
    try {
        drawPrizeMarket($, "TASK_SIGNIN"), $.sleep(6e3), drawPrizeMarket($, "TASK_SIGNIN_PHOTOS"), $.sleep(6e3), drawPrizeMarket($, "TASK_2022_FLDFS_KJ");
    } catch (error) {
        $.logger.error("\u62BD\u5956\u4EFB\u52A1\u5F02\u5E38", error.message);
    }
}
function getSessionKey($) {
    try {
        let { res_code, res_message, sessionKey } = $.api.getUserBriefInfo();
        if (res_code === 0) return sessionKey;
        // $.logger.error("\u83B7\u53D6 sessionKey \u5931\u8D25", res_code, res_message);
    } catch (error) {
        // $.logger.error("\u83B7\u53D6 sessionKey \u5F02\u5E38", error.message);
    }
}
function getAccessToken($, sessionKey) {
    try {
        let data = $.api.getAccessTokenBySsKey(sessionKey, {
            Appkey: Appkey,
            ...getSignature($.md5, {
                Appkey: Appkey,
                sessionKey: sessionKey
            })
        });
        if (data.errorCode) {
            $.logger.error("\u83B7\u53D6 accessToken \u5931\u8D25", data.errorCode, data.errorMsg);
            return;
        }
        return data;
    } catch (error) {
        $.logger.error("\u83B7\u53D6 sessionKey \u5F02\u5E38", error.message);
    }
}
function getFamilyId($, AccessToken) {
    try {
        let xml = $.api.getFamilyList({
            ...getSignature($.md5, {
                AccessToken: AccessToken
            })
        });
        if (isPlainObject(xml)) {
            $.logger.info("\u83B7\u53D6\u5BB6\u5EAD\u51FA\u9519", JSON.stringify(xml));
            return;
        }
        return getXmlElement(xml, "familyId");
    } catch (error) {
        $.logger.error("\u83B7\u53D6\u5BB6\u5EAD\u4FE1\u606F\u5F02\u5E38", error.message);
    }
}
function signInFamily($, AccessToken) {
    if (!AccessToken) 
      // return $.logger.info("\u8BF7\u63D0\u4F9B AccessToken");
      return $.logger.info("");
    try {
        let familyId = getFamilyId($, AccessToken);
        if (!familyId) return $.logger.info("\u7B7E\u5230\u5931\u8D25");
        let xml = $.api.familySign(familyId, {
            ...getSignature($.md5, {
                AccessToken: AccessToken,
                familyId: familyId
            })
        });
        if (isPlainObject(xml)) {
            $.logger.info("\u7B7E\u5230\u5F02\u5E38", JSON.stringify(xml));
            return;
        }
        let bonusSpace = getXmlElement(xml, "bonusSpace");
        $.logger.info(`\u7B7E\u5230\u6210\u529F\uFF0C\u83B7\u5F97${bonusSpace}M\u5BB6\u5EAD\u7A7A\u95F4`);
    } catch (error) {
        $.logger.error("\u5BB6\u5EAD\u7B7E\u5230\u5F02\u5E38", error.message);
    }
}
function getTaskProgress($, times) {
    let returnWarp = (sign = 0, draw = !1)=>({
            sign: sign,
            draw: draw
        });
    try {
        let data = $.api.getListGrow();
        if (data.errorCode) return $.logger.info("\u83B7\u53D6\u7A7A\u95F4\u8BB0\u5F55\u5931\u8D25", data.errorCode, data.errorMsg), returnWarp();
        let todayList = data.growList.filter((item)=>getInToday(item.changeTime) && (item.taskName === "\u6BCF\u65E5\u7B7E\u5230\u9001\u7A7A\u95F4" || item.taskName === "\u5929\u7FFC\u4E91\u76D850M\u7A7A\u95F4"));
        if (todayList.length === 0) return returnWarp();
        let todaySign = todayList.find((item)=>item.taskName === "\u6BCF\u65E5\u7B7E\u5230\u9001\u7A7A\u95F4");
        if (todaySign) return todayList.length >= 3 ? returnWarp(todaySign.changeValue, !0) : returnWarp(todaySign.changeValue, !1);
        if (todayList.length >= 2) return returnWarp(0, !0);
    } catch (error) {
        $.logger.error("\u83B7\u53D6\u7A7A\u95F4\u8BB0\u5F55\u5F02\u5E38", error.message);
    }
    return returnWarp();
}
function appLogin($) {
    try {
        let sessionKey = getSessionKey($);
        if (!sessionKey) return;
        let data = getAccessToken($, sessionKey);
        return data ? data.accessToken : void 0;
    } catch (error) {
        $.logger.error("appLogin \u5F02\u5E38", error.message);
    }
}
function init($) {
    return login($), $.store = {
        AccessToken: appLogin($)
    }, getTaskProgress($, 1);
}
function run($) {
    // $.logger.info("\u5F00\u59CB\u7B7E\u5230", $.config.username);
    try {
        let { draw, sign } = init($);
        sign ? $.logger.info(`\u4ECA\u65E5\u5DF2\u7B7E\u5230\uFF0C\u83B7\u5F97${sign}M\u7A7A\u95F4`) : ($.sleep(2e3), signIn($)), draw ? $.logger.info("\u4ECA\u65E5\u5DF2\u65E0\u62BD\u5956\u6B21\u6570") : ($.sleep(5e3), drawPrizeTask($)), $.sleep(5e3), signInFamily($, $.store && $.store.AccessToken);
    } catch (error) {
        $.logger.error("\u8FD0\u884C\u5F02\u5E38", error.message);
    }
}
function getSignature(md52, data) {
    let Timestamp = String(Date.now()), d = {};
    return data.AccessToken && (d.AccessToken = data.AccessToken), {
        signature: md52(sortStringify({
            Timestamp: Timestamp,
            ...data
        })),
        Timestamp: Timestamp,
        "Sign-Type": "1",
        ...d
    };
}
// ../../core/push/index.ts
function _send({ logger, http }, name = "\u81EA\u5B9A\u4E49\u6D88\u606F", options) {
    try {
        let requestOptions = {
            method: "POST",
            headers: {
                "Content-Type": "application/json"
            },
            timeout: 1e4,
            ...options
        };
        Reflect.has(options, "data") && Reflect.has(options.data, "agent") && (requestOptions.agent = options.data.agent, delete options.data.agent);
        let data = http.fetch(requestOptions), { errcode, code, err } = data;
        if (errcode || err || ![
            0,
            200,
            void 0
        ].some((c)=>code === c)) return logger.error(`${name}\u53D1\u9001\u5931\u8D25`, JSON.stringify(data));
        logger.info(`${name}\u5DF2\u53D1\u9001\uFF01`);
    } catch (error) {
        logger.info(`${name}\u53D1\u9001\u5931\u8D25: ${error.message}`);
    }
}
// function pushplus(apiOption, { token, ...option }, title, text) {
//     return _send(apiOption, "pushplus", {
//         url: "http://www.pushplus.plus/send",
//         method: "POST",
//         data: {
//             token: token,
//             title: title,
//             content: text,
//             ...option
//         }
//     });
// }
// function serverChan(apiOption, { token, ...option }, title, text) {
//     return _send(apiOption, "Server\u9171", {
//         url: `https://sctapi.ftqq.com/${token}.send`,
//         headers: {
//             "content-type": "application/x-www-form-urlencoded"
//         },
//         data: {
//             text: title,
//             desp: text.replace(/\n/g, `

// `),
//             ...option
//         }
//     });
// }
function workWeixin(apiOption, { msgtype = "text", touser = "@all", agentid, corpid, corpsecret, ...option }, title, text) {
    try {
        let { access_token } = apiOption.http.fetch({
            url: "https://qyapi.weixin.qq.com/cgi-bin/gettoken",
            method: "POST",
            data: {
                corpid: corpid,
                corpsecret: corpsecret
            },
            headers: {
                "Content-Type": "application/json"
            }
        });
        return _send(apiOption, "\u4F01\u4E1A\u5FAE\u4FE1\u63A8\u9001", {
            url: `https://qyapi.weixin.qq.com/cgi-bin/message/send?access_token=${access_token}`,
            data: {
                touser: touser,
                msgtype: msgtype,
                agentid: agentid,
                [msgtype]: {
                    content: `${title}

${text}`
                },
                ...option
            }
        });
    } catch (error) {
        apiOption.logger.error("\u4F01\u4E1A\u5FAE\u4FE1\u63A8\u9001\u5931\u8D25"), apiOption.logger.error(error);
    }
}
function workWeixinBot(apiOption, { url, msgtype = "text", ...option }, title, text) {
    return _send(apiOption, "\u4F01\u4E1A\u5FAE\u4FE1Bot\u63A8\u9001", {
        url: url,
        data: {
            msgtype: msgtype,
            [msgtype]: {
                centent: `${title}

${text}`
            },
            ...option
        }
    });
}
// function bark(apiOption, { key, level = "passive", ...options }, title, text) {
//     return _send(apiOption, "Bark ios \u63A8\u9001", {
//         url: `https://api.day.app/${key}`,
//         data: {
//             body: text,
//             title: title,
//             level: level,
//             ...options
//         }
//     });
// }
// ../utils/utils.ts
function arrayMap(arr, cb) {
    let _arr = [];
    for(let i = 0; i < arr.length; i++)_arr.push(cb(arr[i], i, arr));
    return _arr;
}
// ../utils/cookie.ts
function parseCookie(cookie) {
    return cookie.split(/;\s?/).reduce((t, cur, i)=>{
        let a = cur.split("=");
        return a[0] === "" ? t : i === 0 ? {
            key: a[0],
            value: a[1]
        } : {
            ...t,
            [a[0]]: a[1]
        };
    }, {});
}
function parseCookies(cookies) {
    return arrayMap(cookies, parseCookie);
}
function stringifyCookies(cookies) {
    return arrayMap(cookies, stringifyCookie).join(";");
}
function stringifyCookie(cookie) {
    let [[, key], [, value]] = Object.entries(cookie);
    return key + "=" + value;
}
function createCookieJar() {
    let _ = {
        store: [],
        setCookies: function(cookieStrings) {
            cookieStrings && cookieStrings.length !== 0 && (_.store = [
                ..._.store,
                ...parseCookies(cookieStrings)
            ]);
        },
        getCookieString: function() {
            return stringifyCookies(_.store);
        }
    };
    return _;
}
// ../utils/http.ts
function destr(str) {
    try {
        return JSON.parse(str);
    } catch  {
        return str;
    }
}
function stringify(obj) {
    return Object.entries(obj).map(([key, value])=>Array.isArray(value) ? key + "=" + value.join(",") : `${key}=${value}`).join("&");
}
function mergeOptions(options, globalOptions) {
    options.headers = toLowerCaseHeaders(options.headers);
    let newHeaders = {
        ...globalOptions.headers,
        ...options.headers
    }, _options = {
        ...globalOptions,
        ...options,
        headers: newHeaders
    };
    return _options.data && (_options.body = _options.data, delete _options.data), _options.body && _options.headers["content-type"] && _options.headers["content-type"].includes("form-urlencoded") ? _options.body = encodeURI(stringify(_options.body)) : isPlainObject(options.body) && (_options.body = JSON.stringify(options.body)), _options;
}
function createRequest(options = {}) {
    options.headers = toLowerCaseHeaders(options.headers), options.headers["user-agent"] || (options.headers["user-agent"] = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/109.0.0.0 Safari/537.36");
    let globalOptions = {
        method: "POST",
        timeout: 3e4,
        ...options
    };
    globalOptions.cookieJar || (globalOptions.cookieJar = createCookieJar());
    let request = (options2)=>{
        options2 = mergeOptions(options2, globalOptions);
        let hooks = options2.hooks || {};
        hooks.beforeRequest && (options2 = hooks.beforeRequest(options2)), options2.cookieJar && (options2.headers.cookie = options2.cookieJar.getCookieString());
        let resp = HTTP.fetch(options2.url, options2), respHeaders = resp.headers;
        if (globalOptions.cookieJar.setCookies(respHeaders["set-cookie"]), options2.native) return resp;
        switch(options2.responseType){
            case "buffer":
                return resp.binary();
            case "text":
                return resp.text();
            default:
                return destr(resp.text());
        }
    }, http = {
        request: request,
        get: (url, options2 = {})=>request({
                ...options2,
                method: "GET",
                url: url
            }),
        post: (url, body, options2 = {})=>request({
                body: body,
                ...options2,
                method: "POST",
                url: url
            }),
        setOptions: function(options2) {
            return options2.headers = toLowerCaseHeaders(options2.headers), globalOptions.headers = {
                ...globalOptions.headers,
                ...options2.headers
            }, options2.timeout && (globalOptions.timeout = options2.timeout), options2.cookieJar && (globalOptions.cookieJar = options2.cookieJar), http;
        },
        setHeader: function(key, value) {
            return globalOptions.headers[key.toLowerCase()] = value, http;
        },
        setCookie: function(key, value, currentUrl) {
            return globalOptions.cookieJar.setCookies([
                `${key}=${value}`
            ]), http;
        }
    };
    return http;
}
// ../utils/index.ts
function getPushConfig() {
    let usedRange2 = Application.Sheets.Item("\u63A8\u9001").UsedRange;
    if (!usedRange2) return console.log("🍳 \u672A\u5F00\u542F\u63A8\u9001"), {};
    let cells = usedRange2.Columns.Cells, columnEnd = Math.min(50, usedRange2.ColumnEnd), rowEnd = Math.min(50, usedRange2.RowEnd), pushConfig = {};
    for(let option = usedRange2.Column; option <= columnEnd; option++){
        let t = {}, item = cells.Item(option);
        if (item.Text) {
            pushConfig[item.Text] = t;
            for(let kv = 1; kv <= rowEnd; kv++){
                let key = item.Offset(kv).Text;
                key.trim() && (t[key] = valueHandle(item.Offset(kv, 1).Text.trim()));
            }
        }
    }
    let base = pushConfig.base;
    if (!base) return pushConfig;
    return delete pushConfig.base, {
        ...pushConfig,
        ...base
    };
    function valueHandle(value) {
        return value === "TRUE" || value === "\u662F" ? !0 : value === "FALSE" || value === "\u5426" ? !1 : value;
    }
}
function email({ logger }, email2, title, text) {
    try {
        if (!email2 || !email2.pass || !email2.from || !email2.host) return;
        let port = email2.port || 465, toUser = email2.to || email2.from;
        SMTP.login({
            host: email2.host,
            // 域名
            port: port,
            // 端口
            secure: port === 465,
            // TLS
            username: email2.from,
            // 账户名
            password: email2.pass
        }).send({
            from: `${title} <${email2.from}>`,
            to: toUser,
            subject: title,
            text: text.replace(/\n/g, `\r
`)
        }), logger.info("\u90AE\u4EF6\u6D88\u606F\u5DF2\u53D1\u9001");
    } catch (error) {
        logger.error("\u90AE\u4EF6\u6D88\u606F\u53D1\u9001\u5931\u8D25", error.message);
    }
}
function sendNotify(op, data, title, text) {
    let cbs = {
        pushplus: pushplus,
        serverChan: serverChan,
        workWeixin: workWeixin,
        email: email,
        workWeixinBot: workWeixinBot,
        bark: bark
    };
    for (let [name, d] of Object.entries(data))try {
        let cb = cbs[name];
        if (!cb) continue;
        cb(op, d, title, text);
    } catch (error) {
        op.logger.error("\u672A\u77E5\u5F02\u5E38", error.message);
    }
}
function sendWpsNotify(pushData2, pushConfig) {
    if (pushData2.length && pushConfig && !(pushConfig.onlyError && !pushData2.some((el)=>el.type === "error"))) {
        let msg = pushData2.map((m)=>`[${m.type} ${m.date.toLocaleTimeString()}]${m.msg}`).join(`
`);
        msg && sendNotify({
            logger: createLogger(),
            http: {
                fetch: (op)=>(op.data && typeof op.data != "string" && (op.body = JSON.stringify(op.data)), HTTP.fetch(op.url, op).json())
            }
        }, pushConfig, pushConfig.title || "asign \u8FD0\u884C\u63A8\u9001", msg);
    }
}
function _hash(algorithm, input) {
    return Crypto.createHash(algorithm).update(input).digest("hex");
}
function md5(input) {
    return _hash("md5", input);
}
// index.ts
function main(config, option) {
    if (!config) return;
    let logger = createLogger({
        pushData: option && option.pushData
    }), $ = {
        api: createApi(createRequest({
            headers: {
                "content-type": "application/x-www-form-urlencoded",
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:125.0) Gecko/20100101 Firefox/123.0",
                Referer: "https://open.e.189.cn/",
                accept: "application/json"
            }
        })),
        logger: logger,
        sleep: Time.sleep,
        md5: md5,
        config: config
    };
    // $.logger.info("--------------"),
    run($);
}
// var sheet = Application.Sheets.Item("\u5929\u7FFC\u4E91\u76D8") || Application.Sheets.Item("ecloud") || ActiveSheet, usedRange = sheet.UsedRange, columnA = sheet.Columns("A"), len = usedRange.Row + usedRange.Rows.Count - 1, pushData = [];
// for(let i = 1; i <= len; i++){
//     let cell = columnA.Rows(i);
//     cell.Text && (console.log(`\u6267\u884C\u7B2C ${i} \u884C`), main({
//         username: cell.Text,
//         password: cell.Offset(0, 1).Text
//     }, {
//         pushData: pushData
//     }));
// }
// sendWpsNotify(pushData, getPushConfig());

// 引用结束



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
    
    
    // var sheet = Application.Sheets.Item("\u5929\u7FFC\u4E91\u76D8") || Application.Sheets.Item("ecloud") || ActiveSheet, usedRange = sheet.UsedRange, columnA = sheet.Columns("A"), len = usedRange.Row + usedRange.Rows.Count - 1, pushData = [];
    // for(let i = 1; i <= len; i++){
    //     let cell = columnA.Rows(i);
    //     cell.Text && (console.log(`\u6267\u884C\u7B2C ${i} \u884C`), main({
    //         username: cell.Text,
    //         password: cell.Offset(0, 1).Text
    //     }, {
    //         pushData: pushData
    //     }));
    // }
    // sendWpsNotify(pushData, getPushConfig());


    // var sheet = Application.Sheets.Item("\u79FB\u52A8\u4E91\u76D8") || Application.Sheets.Item("caiyun") || ActiveSheet, usedRange = sheet.UsedRange, AColumn = sheet.Columns("A"), len = usedRange.Row + usedRange.Rows.Count - 1, BColumn = sheet.Columns("B"), pushData = [];
    // for(let i = 1; i <= len; i++){
    //     let cell = AColumn.Rows(i);
    //     cell.Text && (console.log(`\u6267\u884C\u7B2C ${i} \u884C`), runMain(i, cell), console.log(`\u7B2C ${i} \u884C\u6267\u884C\u7ED3\u675F`));
    // }
    // sendWpsNotify(pushData, getPushConfig());

    cell = cookie
    i = 0

    username = Application.Range("D" + pos).Text;
    password = Application.Range("E" + pos).Text;
    
    main({
        username: username,
        password: password
    }, {
        pushData: pushData
    });

    console.log(pushData)
    msg = ""
    for(i = 0; i<pushData.length; i++)
    {
      // console.log(pushData[i]["msg"])
      try{
        type = pushData[i]["type"]
        if(type == "info")
        {
          msg += pushData[i]["msg"] + " \n"
        }else
        {
          // console.log("📢 " + pushData[i]["msg"])
        }
        
      }catch{

      }
      
    }
    // console.log("🎉" + msg)
    messageSuccess += "🎉" + msg
    
    // messageSuccess += msg[0]
    // messageFail += msg[1]

    // i = 0
    // const pushData = [];
    // phone = ""
    // auth_token = ""
    // authorization = cookie
    // msg = main(i, phone, auth_token, authorization, {
    //   pushData: pushData,
    // });


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

