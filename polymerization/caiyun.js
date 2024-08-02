// 中国移动云盘签到
// 20240706
// 文中引用代码改编自公众号"Jerry不是猫"
/*
备注：和彩云已更名为中国移动云盘。cookie填写中国移动云盘网页版中获取的authorization。F12 -> Application(中文名叫"应用程序") -> Cookie -> authorization
中国移动云盘网址:https://yun.139.com
*/

pushData = []
let sheetNameSubConfig = "caiyun"; // 分配置表名称
let pushHeader = "【中国移动云盘】";
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
      // 找到指定的表行
      if(Application.Range("A" + (i + 2)).Value == sheetNameSubConfig){
        // 写入更新的时间
        Application.Range("F" + (i + 2)).Value = todayDate
        // 写入消息
        Application.Range("G" + (i + 2)).Value = message
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
 
// 引用开始
function randomHex(length, pad = "-") {
        return Array.isArray(length) ? length.map((l)=>randomHex(l, pad)).join(pad) : Array.from({
            length: length
        }).map(()=>Math.floor(Math.random() * 16).toString(16)).join("");
    }
    function randomNumber(low, high = low) {
        return Math.floor(Math.random() * (high - low) + low);
    }
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
            console[type](...args);
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
    function getHostname(url) {
        return url.split("/")[2].split("?")[0];
    }
    function asyncForEach(array, task, cb) {
        let len2 = array.length;
        for(let index = 0; index < len2; index++){
            let item = array[index];
            task(item), index < len2 - 1 && cb && cb();
        }
    }
    function setStoreArray(store, key, values) {
        return Reflect.has(store, key) ? Reflect.set(store, key, Reflect.get(store, key).concat(values)) : Reflect.set(store, key, values);
    }
    function getAuthInfo(basicToken) {
        basicToken = basicToken.replace("Basic ", "");
        let rawToken = Buffer.from(basicToken, "base64").toString("utf-8"), [platform, phone, token] = rawToken.split(":");
        return {
            phone: phone,
            token: token,
            auth: `Basic ${basicToken}`,
            platform: platform
        };
    }
    function hashCode(str) {
        if (typeof str != "string") return 0;
        let hash = 0, char = null;
        if (str.length == 0) return hash;
        for(let i = 0; i < str.length; i++)char = str.charCodeAt(i), hash = (hash << 5) - hash + char, hash = hash & hash;
        return hash;
    }
    function isWps() {
        return globalThis.setTimeout === void 0 && globalThis.HTTP;
    }
    function createTime() {
        return /* @__PURE__ */ new Date().toLocaleString("zh-CN").split(/[/\W\s:-]/).reduce((str, cur)=>`${str}${cur.length === 1 ? 0 + cur : cur}`, "");
    }
    function toLowerCaseHeaders(headers) {
        return headers ? Object.entries(headers).reduce((acc, [key, value])=>(acc[key.toLowerCase()] = value, acc), {}) : {};
    }
    function isPlainObject(obj) {
        return Array.isArray(obj) || Object.prototype.toString.call(obj) === "[object Object]";
    }
    // ../../core/caiyun/constant/taskList.ts
    var emailTaskList = {
        /**
     * 从固定入口签到
     */ 1005: {
            id: 1005,
            group: "time"
        },
        /**
     * 去“发现广场”浏览精彩内容
     */ 1008: {
            id: 1008,
            group: "month"
        },
        /**
     * 前往“云盘”查看个人动态
     */ 1009: {
            id: 1009,
            group: "month"
        },
        /**
     * 浏览限免影视大片
     */ 1010: {
            id: 1010,
            group: "month"
        },
        /**
     * 查看“我的附件”
     */ 1013: {
            id: 1013,
            group: "month"
        },
        /**
     * 体验“PDF转换”功能
     */ 1014: {
            id: 1014,
            group: "month"
        },
        /**
     * 体验“云笔记”功能
     */ 1016: {
            id: 1016,
            group: "month"
        },
        /**
     * 登录移动云盘APP云朵中心
     */ 1017: {
            id: 1017,
            group: "month"
        }
    }, cloudTaskList = {
        /**
     * 手动上传一个文件
     */ 106: {
            id: 106,
            runner: !0,
            group: "day"
        },
        /**
     * 创建一篇云笔记
     */ 107: {
            id: 107,
            runner: !0,
            group: "day"
        },
        /**
     * 当月上传容量满1G
     */ 110: {
            id: 110,
            runner: !0,
            group: "month"
        },
        /**
     * 使用PC客户端
     */ 113: {
            id: 113,
            group: "month"
        },
        /**
     * 浏览APP-我的-活动中心
     */ 115: {
            id: 115,
            group: "month"
        },
        /**
     * 从固定入口访问云朵中心
     */ 409: {
            id: 409,
            group: "month"
        },
        /**
     * 去发现频道看大片
     */ 426: {
            id: 426,
            group: "month"
        },
        /**
     * 使用移动云手机
     */ 431: {
            id: 431,
            group: "month"
        },
        /**
     * 分享文件有好礼
     */ 434: {
            id: 434,
            runner: void 0,
            group: "day"
        },
        /**
     * 体验云盘AI功能
     */ 435: {
            id: 435,
            group: "month"
        }
    }, hotTaskList = {
        /**
     * 分享文件有好礼
     */ 434: {
            id: 434,
            group: "time"
        },
        /**
     * 去中国移动APP领好礼
     */ 447: {
            id: 447,
            group: "time"
        }
    }, TASK_LIST = {
        ...emailTaskList,
        ...cloudTaskList,
        ...hotTaskList
    }, SKIP_TASK_LIST = [
        /** 邀请新用户 */ 117,
        /** 邀请好友看电影 */ 437,
        /** 组团领现金 */ 991,
        /** 开启通知领云朵 */ 406,
        /** 月月备份领好礼 */ 385,
        /** 30G流量0元领 */ 120,
        /** 月月开通 月月有礼 */ 1021
    ];
    // ../../core/caiyun/utils/index.ts
    function request($, api, options, ...args) {
        let name = typeof options == "string" ? options : options.name;
        try {
            let { code, message, msg, result } = api(...args);
            if (code !== 0) $.logger.error(`${name}\u5931\u8D25`, code, message || msg);
            else return result;
        } catch (error) {
            $.logger.error(`${name}\u5F02\u5E38`, error.message);
        }
        return typeof options == "object" ? options.isArray === !1 ? {} : [] : {};
    }
    // ../../core/caiyun/service/backupGift.ts
    function backupGiftTask($) {
        // $.logger.info("------\u3010\u5907\u4EFD\u597D\u793C\u3011------");
        try {
            let { state, curMonth } = request($, $.api.getBackupGift, "backupGiftTask");
            if (!curMonth) return;
            switch(state){
                case -1:
                    $.logger.warn("\u672A\u5F00\u542F\u5907\u4EFD\uFF0C\u8BF7\u524D\u5F80 APP \u624B\u52A8\u5F00\u542F");
                    return;
                case 1:
                    $.logger.info("\u672C\u6708\u5DF2\u9886\u53D6");
                    return;
                case 0:
                    $.logger.info("\u9886\u53D6\u5907\u4EFD\u5956\u52B1");
                    let result = request($, $.api.receiveBackupGift, "backupGiftTask");
                    typeof result == "number" && $.logger.info("\u9886\u53D6\u6210\u529F\uFF0C\u83B7\u5F97\u4E91\u6735", result);
                    return;
                default:
                    $.logger.warn("\u672A\u77E5\u72B6\u6001", state);
            }
        } catch (error) {
            $.logger.error("\u5907\u4EFD\u597D\u793C", error.message);
        }
    }
    // ../../core/caiyun/constant/index.ts
    var caiyunUrl = "https://caiyun.feixin.10086.cn", marketUrl = "https://caiyun.feixin.10086.cn/market", mw2TogetherUrl = "https://html5.mail.10086.cn/mw2/together/s", riddlesUrl = "https://caiyun.feixin.10086.cn/market/lanternriddles/answeredPuzzles";
    // ../../core/caiyun/service/aiRedPack.ts
    function aiRedPackTask($) {
        if ($.logger.info("------\u3010AI\u7EA2\u5305\u3011------"), isWps()) {
            $.logger.info("AI\u7EA2\u5305", "\u5F53\u524D\u73AF\u5883\u4E3AWPS\uFF0C\u8DF3\u8FC7");
            return;
        }
        try {
            blindboxJournaling($);
            let sid = init($);
            if (!sid) return;
            let count = 4;
            for(; !_task($, sid) && count > 0;)count--;
        } catch (error) {
            $.logger.error("AI\u7EA2\u5305", error.message);
        }
    }
    function blindboxJournaling({ api, sleep }) {
        let marketName = "&marketName=National_LanternRiddlesal_LanternRiddles";
        api.journaling("National_LanternRiddles_client_all", 1008, marketName), sleep(200), api.journaling("National_LanternRiddles_pv", 1008, marketName), sleep(200), api.journaling("National_LanternRiddles_client_isApp", 1008, marketName), sleep(200);
    }
    function _task($, sid) {
        let sleep = $.sleep;
        sleep(1e3);
        let puzzleCards = getPuzzleCards($);
        if (!puzzleCards) return;
        let puzzleCard = puzzleCards[randomNumber(0, puzzleCards.length - 1)];
        $.logger.debug("\u8C1C\u9762", puzzleCard.puzzle), sleep(200);
        let taskId = getMailChatTaskId($, sid, puzzleCard.puzzle);
        if (!taskId) return;
        sleep(1e3);
        let msg = getMailChatMsg($, sid, taskId);
        if (!msg) return;
        let tip = matchResult(msg);
        if (!tip) return $.logger.info("AI\u7EA2\u5305", "\u672A\u83B7\u53D6\u5230\u8C1C\u5E95");
        $.logger.debug("\u8C1C\u5E95", tip), sleep(400);
        let answered = submitAnswered($, puzzleCard.id, tip);
        if (answered === -1) return !0;
        if (answered !== 0) return;
        sleep(400);
        let prizeName = openRedPack($, puzzleCard.id);
        if (prizeName === -1) return !0;
        if (typeof prizeName != "number" && ($.logger.info("\u83B7\u5F97", prizeName), prizeName !== "\u8C22\u8C22\u53C2\u4E0E")) return !0;
    }
    function init($) {
        try {
            let loginInfo = loginEmail($);
            if (!loginInfo) return;
            let { sid, rmkey } = loginInfo;
            if (sid) return $.http.setCookie("sid", sid, caiyunUrl), $.http.setCookie("RMKEY", rmkey, mw2TogetherUrl), sid;
        } catch (error) {
            $.logger.error("\u90AE\u7BB1\u767B\u5F55\u5F02\u5E38", error.message);
        }
    }
    function getPuzzleCards($) {
        let puzzleCards = request($, $.api.getIndexPuzzleCard, {
            name: "\u83B7\u53D6\u8C1C\u5E95\u5217\u8868"
        });
        if (!puzzleCards.length) {
            $.logger.info("\u672A\u83B7\u53D6\u5230\u8C1C\u5E95\u5217\u8868\uFF0C\u8DF3\u8FC7\u4EFB\u52A1");
            return;
        }
        return puzzleCards.filter((item)=>item.puzzleTitleContext).map((item)=>({
                puzzle: item.puzzleTitleContext + "---" + item.puzzleTipContext,
                id: item.id
            }));
    }
    function getMailChatTaskId($, sid, q) {
        try {
            let { code, taskId, summary } = $.api.mailChatTask(sid, q);
            if (code !== "S_OK") {
                $.logger.info("\u83B7\u53D6AI\u804A\u5929\u4EFB\u52A1ID\u5931\u8D25", summary);
                return;
            }
            return taskId;
        } catch (error) {
            $.logger.error("\u83B7\u53D6AI\u804A\u5929\u4EFB\u52A1ID\u5F02\u5E38", error.message);
        }
    }
    function getMailChatMsg($, sid, id) {
        try {
            let { code, summary, var: data } = $.api.mailChatMsg(sid, id);
            if (code !== "S_OK" || data && data.result === "") {
                $.logger.info("\u83B7\u53D6AI\u804A\u5929\u6D88\u606F\u5931\u8D25", summary);
                return;
            }
            return data.result;
        } catch (error) {
            $.logger.error("\u83B7\u53D6AI\u804A\u5929\u6D88\u606F\u5F02\u5E38", error.message);
        }
    }
    function matchResult(result) {
        let match = result.split(/[—-]/);
        return match[match.length - 1];
    }
    function submitAnswered($, id, a) {
        try {
            let { code, msg } = $.api.submitAnswered(id, a);
            switch(code){
                case 0:
                    return $.logger.debug("\u56DE\u7B54\u95EE\u9898\u6210\u529F"), 0;
                case 201:
                    return $.logger.info("\u56DE\u7B54\u95EE\u9898\u6210\u529F\uFF0C\u4F46", code, msg), -1;
                default:
                    return $.logger.info("\u56DE\u7B54\u95EE\u9898\u5931\u8D25", code, msg), 1;
            }
        } catch (error) {
            $.logger.error("\u56DE\u7B54\u7B54\u6848\u5F02\u5E38", error.message);
        }
        return 2;
    }
    function openRedPack($, puzzleId) {
        try {
            let { code, msg, result } = $.api.getAwarding(puzzleId);
            switch(code){
                case 0:
                    return result.prizeName;
                case 10010020:
                    return $.logger.info("\u53EF\u80FD\u4F60\u9700\u8981\u53BB APP \u624B\u52A8\u5B8C\u6210\u4E00\u6B21", code, msg), -1;
                default:
                    return $.logger.info("\u6253\u5F00\u7EA2\u5305\u5931\u8D25", code, msg), 1;
            }
        } catch (error) {
            $.logger.error("\u6253\u5F00\u7EA2\u5305", error.message);
        }
        return 2;
    }
    // ../../core/caiyun/service/index.ts
    function uploadFileRequest($, parentCatalogID, { ext = ".png", digest = randomHex(32).toUpperCase(), contentSize = randomNumber(1, 1e3), manualRename = 2, contentName = "asign-" + randomHex(4) + ext, createTime: createTime2 = createTime() } = {}, needUpload) {
        try {
            let xml = $.api.uploadFileRequest({
                phone: $.config.phone,
                parentCatalogID: parentCatalogID,
                contentSize: contentSize,
                createTime: createTime2,
                digest: digest,
                manualRename: manualRename,
                contentName: contentName
            }), isNeedUpload = getXmlElement(xml, "isNeedUpload"), contentID = getXmlElement(xml, "contentID");
            if (isNeedUpload === "1") return needUpload ? {
                redirectionUrl: getXmlElement(xml, "redirectionUrl"),
                uploadTaskID: getXmlElement(xml, "uploadTaskID"),
                contentName: getXmlElement(xml, "contentName"),
                contentID: contentID
            } : ($.logger.info("\u672A\u627E\u5230\u8BE5\u6587\u4EF6\uFF0C\u8BE5\u6587\u4EF6\u9700\u8981\u624B\u52A8\u4E0A\u4F20"), {});
            if (contentID) return contentID && setStoreArray($.store, "files", [
                contentID
            ]), {
                contentID: contentID
            };
            $.logger.error("\u4E0A\u4F20\u6587\u4EF6\u8BF7\u6C42\u5931\u8D25", xml);
        } catch (error) {
            $.logger.error("\u4E0A\u4F20\u6587\u4EF6\u8BF7\u6C42\u5F02\u5E38", error.message);
        }
        return {};
    }
    function uploadFile($, parentCatalogID, { ext = ".png", digest = randomHex(32).toUpperCase(), contentSize = randomNumber(1, 1e3), manualRename = 2, contentName = "asign-" + randomHex(4) + ext, createTime: createTime2 = createTime() } = {}, file) {
        try {
            let { redirectionUrl, uploadTaskID, contentID } = uploadFileRequest($, parentCatalogID, {
                ext: ext,
                digest: digest,
                contentSize: contentSize,
                manualRename: manualRename,
                contentName: contentName,
                createTime: createTime2
            }, !0);
            if (!redirectionUrl || !file) return !!contentID;
            let size = typeof file == "string" ? file.length : file.byteLength;
            $.logger.debug("\u522B\u7740\u6025\uFF0C\u6587\u4EF6\u4E0A\u4F20\u4E2D\u3002\u3002\u3002");
            let xml = $.api.uploadFile(redirectionUrl.replace(/&amp;/g, "&"), uploadTaskID, file, size);
            switch(getXmlElement(xml, "resultCode")){
                case "0":
                    return contentID && setStoreArray($.store, "files", [
                        contentID
                    ]), !0;
                case "9119":
                    return $.logger.info("\u4E0A\u4F20\u6587\u4EF6\u5931\u8D25\uFF1Amd5\u6821\u9A8C\u5931\u8D25"), !1;
                default:
                    return $.logger.error("\u4E0A\u4F20\u6587\u4EF6\u5931\u8D25", xml), !1;
            }
        } catch (error) {
            $.logger.error("\u4E0A\u4F20\u6587\u4EF6\u5F02\u5E38", error.message);
        }
        return !1;
    }
    function pcUploadFileRequest($, path) {
        try {
            let { success, message, data } = $.api.pcUploadFileRequest($.config.phone, path, 0, "asign" + randomHex(4) + ".png", "d41d8cd98f00b204e9800998ecf8427e");
            if (success && data && data.uploadResult) return data.uploadResult.newContentIDList.map(({ contentID })=>contentID);
            $.logger.error("\u4E0A\u4F20\u6587\u4EF6\u8BF7\u6C42\u5931\u8D25", message);
        } catch (error) {
            $.logger.error("\u4E0A\u4F20\u6587\u4EF6\u8BF7\u6C42\u5F02\u5E38", error.message);
        }
    }
    function getParentCatalogID() {
        return "00019700101000000001";
    }
    function getBackParentCatalogID() {
        return "00019700101000000043";
    }
    // ../../core/caiyun/service/garden.ts
    function request2($, api, options, ...args) {
        let name, defu;
        typeof options == "string" ? (name = options, defu = {}) : (name = options.name, defu = options.defu);
        try {
            let { success, msg, result } = api(...args);
            if (!success) $.logger.error(`${name}\u5931\u8D25`, msg);
            else return result;
        } catch (error) {
            $.logger.error(`${name}\u5F02\u5E38`, error.message);
        }
        return defu;
    }
    function loginGarden($, token, phone) {
        try {
            $.gardenApi.login(token, phone);
        } catch (error) {
            $.logger.error("\u767B\u5F55\u679C\u56ED\u5931\u8D25", error.message);
        }
    }
    function getTodaySign($) {
        let { todayCheckin } = request2($, $.gardenApi.checkinInfo, "\u83B7\u53D6\u679C\u56ED\u7B7E\u5230\u4FE1\u606F");
        return todayCheckin;
    }
    function initTree($) {
        let { collectWater, treeLevel, nickName } = request2($, $.gardenApi.initTree, "\u521D\u59CB\u5316\u679C\u56ED");
        return nickName ? ($.logger.info(`${nickName}\u62E5\u6709${treeLevel}\u7EA7\u679C\u6811\uFF0C\u5F53\u524D\u6C34\u6EF4${collectWater}`), !0) : !1;
    }
    function signInGarden($) {
        let todaySign = getTodaySign($);
        if (todaySign !== void 0) {
            if (todaySign) return $.logger.info("\u4ECA\u65E5\u679C\u56ED\u5DF2\u7B7E\u5230");
            try {
                let { code, msg } = request2($, $.gardenApi.checkin, "\u679C\u56ED\u7B7E\u5230");
                code !== 1 && $.logger.error("\u679C\u56ED\u7B7E\u5230\u5931\u8D25", code, msg);
            } catch (error) {
                $.logger.error("\u679C\u56ED\u7B7E\u5230\u5F02\u5E38", error.message);
            }
        }
    }
    function clickCartoon($) {
        let cartoonTypes = request2($, $.gardenApi.getCartoons, {
            name: "\u83B7\u53D6\u5DF2\u7ECF\u5B8C\u6210\u7684\u573A\u666F\u5217\u8868",
            defu: []
        });
        asyncForEach([
            "cloud",
            "color",
            "widget",
            "mail"
        ].filter((cart)=>!cartoonTypes.includes(cart)), (cartoonType)=>{
            let { msg, code } = request2($, $.gardenApi.clickCartoon, "\u9886\u53D6\u573A\u666F\u6C34\u6EF4", cartoonType);
            [
                1,
                -1,
                -2
            ].includes(code) ? $.logger.debug(`\u9886\u53D6\u573A\u666F\u6C34\u6EF4${cartoonType}`) : $.logger.error(`\u9886\u53D6\u573A\u666F\u6C34\u6EF4${cartoonType}\u5931\u8D25`, code, msg);
        }, ()=>$.sleep(5e3));
    }
    function getTaskList($, headers) {
        return request2($, $.gardenApi.getTaskList, {
            name: "\u83B7\u53D6\u4EFB\u52A1\u5217\u8868",
            defu: []
        }, headers);
    }
    function getTaskStateList($, headers) {
        return request2($, $.gardenApi.getTaskStateList, {
            name: "\u83B7\u53D6\u4EFB\u52A1\u5B8C\u6210\u60C5\u51B5\u8868",
            defu: []
        }, headers);
    }
    function doTask($, tasks, headers) {
        let fileInfo = {
            contentSize: "133967",
            digest: $.config.garden.digest
        }, taskMap = {
            2002: ()=>{
                if (uploadFileRequest($, getParentCatalogID(), {
                    ext: ".png",
                    ...fileInfo
                })) return $.logger.debug("\u4E0A\u4F20\u56FE\u7247\u6210\u529F"), $.sleep(2e3), !0;
            },
            2003: ()=>{
                if (uploadFileRequest($, getParentCatalogID(), {
                    ext: ".mp4",
                    ...fileInfo
                })) return $.logger.debug("\u4E0A\u4F20\u89C6\u9891\u6210\u529F"), $.sleep(2e3), !0;
            }
        };
        asyncForEach(tasks, ({ taskId, taskName })=>{
            let { code, summary } = request2($, $.gardenApi.doTask, `\u63A5\u6536${taskName}\u4EFB\u52A1`, taskId, headers);
            switch(code){
                case -1:
                case 1:
                    $.logger.debug(`${taskName}\u4EFB\u52A1\u5DF2\u9886\u53D6`), taskMap[taskId] && taskMap[taskId]();
                    return;
                default:
                    $.logger.error(`\u9886\u53D6${taskName}\u5931\u8D25`, code, summary);
                    return;
            }
        }, ()=>$.sleep(6e3));
    }
    function doTaskByHeaders($, headers) {
        try {
            let taskList = getTaskList($, headers), _stateList = getTaskStateList($, headers), stateList = _stateList.reduce((arr, { taskId, taskState })=>taskState === 0 ? [
                    ...arr,
                    taskId
                ] : arr, []);
            if (_stateList.length && !stateList.length) return;
            return $.sleep(2e3), _run(stateList.length ? taskList.filter((task)=>stateList.indexOf(task.taskId) !== -1) : taskList);
            function _run(_taskList) {
                $.sleep(5e3), doTask($, _taskList, headers), $.sleep(4e3);
                let stateList2 = getTaskStateList($, headers).filter(({ taskState })=>taskState === 1).map(({ taskId })=>taskId);
                $.sleep(200), givenWater($, taskList.filter((task)=>stateList2.indexOf(task.taskId) !== -1), headers);
            }
        } catch (error) {
            $.logger.error("\u4EFB\u52A1\u5F02\u5E38", error.message);
        }
    }
    function givenWater($, tasks, headers) {
        asyncForEach(tasks, ({ taskName, taskId })=>{
            let { water, msg } = request2($, $.gardenApi.givenWater, `\u9886\u53D6${taskName}\u6C34\u6EF4`, taskId, headers);
            water === 0 ? $.logger.error(`\u9886\u53D6${taskName}\u5956\u52B1\u5931\u8D25`, msg) : $.logger.info(`\u9886\u53D6${taskName}\u5956\u52B1`);
        }, ()=>$.sleep(6e3));
    }
    function gardenTask($) {
        try {
            $.logger.info("------\u3010\u679C\u56ED\u3011------");
            let token = getSsoTokenApi($, $.config.phone);
            if (!token) return $.logger.error("\u8DF3\u8FC7\u679C\u56ED\u4EFB\u52A1");
            if (loginGarden($, token, $.config.phone), !initTree($)) {
                $.logger.warn("\u83B7\u53D6\u679C\u56ED\u4FE1\u606F\u5931\u8D25\uFF0C\u8BF7\u786E\u8BA4\u5DF2\u7ECF\u6FC0\u6D3B\u679C\u56ED");
                return;
            }
            signInGarden($), $.sleep(2e3), $.logger.info("\u9886\u53D6\u573A\u666F\u6C34\u6EF4"), clickCartoon($), $.logger.info("\u5B8C\u6210\u90AE\u7BB1\u4EFB\u52A1"), doTaskByHeaders($, {
                "user-agent": $.DATA.baseUA + $.DATA.mailUaEnd,
                "x-requested-with": $.DATA.mailRequested
            }), $.sleep(2e3), $.logger.info("\u5B8C\u6210\u4E91\u76D8\u4EFB\u52A1"), doTaskByHeaders($, {
                "user-agent": $.DATA.baseUA,
                "x-requested-with": $.DATA.mcloudRequested
            });
        } catch (error) {
            $.logger.error("\u679C\u56ED\u4EFB\u52A1\u5F02\u5E38", error.message);
        }
    }
    // ../../core/caiyun/service/msgPush.ts
    function msgPushOnTask($) {
        $.logger.info("------\u3010\u6D88\u606F\u63A8\u9001\u5956\u52B1\u3011------");
        try {
            let { onDuaration, pushOn, secondTaskStatus, firstTaskStatus } = request($, $.api.getMsgPushOn, "\u83B7\u53D6\u6D88\u606F\u901A\u77E5\u72B6\u6001");
            if (pushOn === 0) {
                $.logger.error("\u6D88\u606F\u901A\u77E5\u5DF2\u5173\u95ED\uFF0C\u8BF7\u524D\u5F80 APP \u624B\u52A8\u6253\u5F00");
                return;
            }
            if (firstTaskStatus !== 3 && $.logger.info("\u9996\u6B21\u5956\u52B1\u672A\u9886\u53D6\uFF0C\u8BF7\u524D\u5F80 APP \u624B\u52A8\u5B8C\u6210"), secondTaskStatus === 2) {
                $.logger.info("\u9886\u53D6\u5956\u52B1"), request($, $.api.obtainMsgPushOn, "\u9886\u53D6\u6D88\u606F\u901A\u77E5\u5956\u52B1");
                return;
            }
            $.logger.info("\u5DF2\u7ECF\u5F00\u542F", onDuaration, "\u5929");
        } catch (error) {
            $.logger.error("\u6D88\u606F\u63A8\u9001\u5956\u52B1", error.message);
        }
    }
    // ../../core/caiyun/service/taskExpansion.ts
    function taskExpansionTask($) {
        return _taskExpansion($, 1);
    }
    function backFile($) {
        try {
            let buffer = randomHex(32), digest = $.md5(buffer), success = uploadFile($, getBackParentCatalogID(), {
                manualRename: 5,
                digest: digest
            }, buffer);
            return success ? $.logger.debug("\u6587\u4EF6\u5907\u4EFD\u6210\u529F") : $.logger.debug("\u6587\u4EF6\u5907\u4EFD\u5931\u8D25"), success;
        } catch (error) {
            $.logger.error("\u6587\u4EF6\u5907\u4EFD\u5F02\u5E38", error.message);
        }
    }
    function _taskExpansion($, timer = 1) {
        let { curMonthBackup, curMonthTaskRecordCount, acceptDate, nextMonthTaskRecordCount } = request($, $.api.getTaskExpansion, "\u83B7\u53D6\u5907\u4EFD\u989D\u5916\u5956\u52B1");
        if (curMonthBackup === !1) {
            // 关闭文件备份
            // if (timer > 0) return backFile($), $.logger.debug("\u7B49\u5F85\u4E00\u6BB5\u65F6\u95F4\u540E\u91CD\u8BD5"), $.sleep($.config.backupWaitTime * 1e3), _taskExpansion($, --timer);
            // $.logger.warn("\u672C\u6708\u672A\u5F00\u542F\u5907\u4EFD\uFF0C\u5C06\u65E0\u6CD5\u83B7\u53D6\u7FFB\u500D\u5956\u52B1\uFF01\uFF01\uFF01\u9700\u8981\u624B\u52A8\u5F00\u542F"), $.store.curMonthBackup = !1;
        }
        if (curMonthTaskRecordCount > 0) {
            let { cloudCount } = request($, $.api.receiveTaskExpansion, "\u9886\u53D6\u7FFB\u500D\u5956\u52B1", acceptDate);
            cloudCount && $.logger.info(`\u9886\u53D6\u5230${cloudCount}\u4E2A\u4E91\u6735`);
        }
        nextMonthTaskRecordCount && $.logger.debug(`\u4E0B\u6708\u53EF\u9886\u53D6${nextMonthTaskRecordCount}\u4E2A\u4E91\u6735`);
    }
    // ../../core/caiyun/api/aiRedPack.ts
    function createAiRedPackApi(http) {
        return {
            submitAnswered: function(puzzleId, answered) {
                return http.get(`${riddlesUrl}/submitAnswered?puzzleId=${puzzleId}&answered=${answered}`);
            },
            getIndexPuzzleCard: function() {
                return http.get(`${riddlesUrl}/getIndexPuzzleCard`);
            },
            getAwarding: function(puzzleId) {
                return http.post(`${riddlesUrl}/awarding`, {
                    puzzleId: puzzleId
                });
            }
        };
    }
    // ../../core/caiyun/api/backupGift.ts
    function createBackupGiftApi(http) {
        return {
            getBackupGift: function() {
                return http.get(`${marketUrl}/backupgift/info`);
            },
            receiveBackupGift: function() {
                return http.get(`${marketUrl}/backupgift/receive`);
            }
        };
    }
    // ../../core/caiyun/api/mailChat.ts
    function createMailChatApi(http) {
        function _together(name, sid, data) {
            return http.post(`${mw2TogetherUrl}?func=together:${name}&sid=${sid}&behaviorData=&rnd=${Math.random()}&cguid=2348352175888&k=5921&comefrom=2066`, data);
        }
        return {
            mailChatTask: function(sid, question) {
                return _together("mailChatTask", sid, `<object>
      <string name="content">\u4F60\u662F\u4F4D\u731C\u706F\u8C1C\u5927\u5E08\uFF0C\u8BF7\u6839\u636E\u6211\u63D0\u4F9B\u7684\u706F\u8C1C\u8C1C\u9762\uFF0C\u521B\u4F5C\u5BF9\u5E94\u8C1C\u5E95\uFF0C\u8FD4\u56DE\u7684\u683C\u5F0F\u8981\u6C42\uFF1A\u9700\u8981\u8FD4\u56DE\u8C1C\u9762\u548C\u8C1C\u5E95\uFF0C\u706F\u8C1C\u8C1C\u9762\u4E3A\uFF1A${question}</string>
      <string name="clientId">10109</string>
      <string name="configId">707</string>
      <string name="model">blian</string>
      <string name="talkType">1</string>
      <string name="createSession" />
      <string name="title">\u731C\u706F\u8C1C\u2014\u2014${question}</string>
      <string name="defaultQuestion">{&quot;question&quot;:&quot;${question}&quot;,&quot;orderTip&quot;:&quot;\u731C\u706F\u8C1C&quot;,&quot;businessCode&quot;:&quot;14&quot;}</string>
      <string name="userActiveType">0</string>
    </object>`);
            },
            mailChatMsg: function(sid, taskId) {
                return _together("mailChatMsg", sid, `<object>
      <string name="taskId">${taskId}</string>
      <string name="model">blian</string>
    </object>`);
            }
        };
    }
    // ../../core/caiyun/api/msgPush.ts
    function createMsgPushApi(http) {
        return {
            getMsgPushOn: function() {
                return http.get(`${marketUrl}/msgPushOn/task/status`);
            },
            obtainMsgPushOn: function() {
                return http.post(`${marketUrl}/msgPushOn/task/obtain`, {
                    type: 2
                });
            }
        };
    }
    // ../../core/caiyun/api/signin.ts
    function createSignInApi(http) {
        let caiyunUrl2 = "https://caiyun.feixin.10086.cn/market/signin/";
        return {
            getTaskExpansion: function() {
                return http.get(`${caiyunUrl2}page/taskExpansion`);
            },
            receiveTaskExpansion: function(acceptDate) {
                return http.get(`${caiyunUrl2}page/receiveTaskExpansion?acceptDate=${acceptDate}`);
            }
        };
    }
    // ../../core/caiyun/api/garden.ts
    function createGardenApi(http) {
        let gardenUrl = "https://happy.mail.10086.cn/jsp/cn/garden";
        return {
            login: function(token, account) {
                return http.get(`${gardenUrl}/login/caiyunsso.do?token=${token}&account=${account}&targetSourceId=001208&sourceid=1014&enableShare=1`, {
                    followRedirect: !1,
                    native: !0
                });
            },
            checkinInfo: function() {
                return http.get(`${gardenUrl}/task/checkinInfo.do`);
            },
            initTree: function() {
                return http.get(`${gardenUrl}/user/initTree.do`);
            },
            /**
       * 需要对应 ua
       */ getTaskList: function(headers = {}) {
                return http.get(`${gardenUrl}/task/taskList.do?clientType=PE`, {
                    headers: headers
                });
            },
            getTaskStateList: function(headers = {}) {
                return http.get(`${gardenUrl}/task/taskState.do`, {
                    headers: headers
                });
            },
            checkin: function() {
                return http.get(`${gardenUrl}/task/checkin.do`);
            },
            clickCartoon: function(cartoonType) {
                return http.get(`${gardenUrl}/user/clickCartoon.do?cartoonType=${cartoonType}`);
            },
            doTask: function(taskId, headers = {}) {
                return http.get(`${gardenUrl}/task/doTask.do?taskId=${taskId}`, {
                    headers: headers
                });
            },
            givenWater: function(taskId, headers = {}) {
                return http.get(`${gardenUrl}/task/givenWater.do?taskId=${taskId}`, {
                    headers: headers
                });
            },
            getCartoons: function(headers = {}) {
                return http.get(`${gardenUrl}/user/gotCartoons.do`, {
                    headers: headers
                });
            }
        };
    }
    // ../../core/caiyun/api.ts
    function createApi(http) {
        let yun139Url = "https://yun.139.com", caiyunUrl2 = "https://caiyun.feixin.10086.cn", mnoteUrl = "https://mnote.caiyun.feixin.10086.cn";
        return {
            querySpecToken: function(account, toSourceId = "001005") {
                return http.post(`${yun139Url}/orchestration/auth-rebuild/token/v1.0/querySpecToken`, {
                    toSourceId: toSourceId,
                    account: String(account),
                    commonAccountInfo: {
                        account: String(account),
                        accountType: 1
                    }
                }, {
                    headers: {
                        referer: "https://yun.139.com/w/",
                        accept: "application/json, text/plain, */*",
                        "content-type": "application/json",
                        "accept-language": "zh-CN,zh;q=0.9"
                    }
                });
            },
            authTokenRefresh: function(token, account) {
                return http.post("https://aas.caiyun.feixin.10086.cn/tellin/authTokenRefresh.do", `<?xml version="1.0" encoding="utf-8"?><root><token>${token}</token><account>${account}</account><clienttype>656</clienttype></root>`, {
                    headers: {
                        accept: "*/*",
                        "content-type": "application/json"
                    },
                    responseType: "text"
                });
            },
            getNoteAuthToken: function(token, account) {
                let headers = http.post(`${mnoteUrl}/noteServer/api/authTokenRefresh.do`, {
                    authToken: token,
                    userPhone: String(account)
                }, {
                    headers: {
                        APP_CP: "pc",
                        APP_NUMBER: String(account),
                        CP_VERSION: "7.7.1.20240115"
                    },
                    native: !0
                }).headers;
                if (headers.app_auth) return {
                    app_auth: headers.app_auth,
                    app_number: headers.app_number,
                    note_token: headers.note_token
                };
            },
            syncNoteBook: function(headers) {
                return http.post(`${mnoteUrl}/noteServer/api/syncNotebook.do `, {
                    addNotebooks: [],
                    delNotebooks: [],
                    updateNotebooks: []
                }, {
                    headers: {
                        APP_CP: "pc",
                        CP_VERSION: "7.7.1.20240115",
                        ...headers
                    }
                });
            },
            createNote: function(noteId, title, account, headers, tags = []) {
                return http.post(`${mnoteUrl}/noteServer/api/createNote.do`, {
                    archived: 0,
                    attachmentdir: "",
                    attachmentdirid: "",
                    attachments: [],
                    contentid: "",
                    contents: [
                        {
                            data: "<span></span>",
                            noteId: noteId,
                            sortOrder: 0,
                            type: "TEXT"
                        }
                    ],
                    cp: "",
                    createtime: String(Date.now()),
                    description: "",
                    expands: {
                        noteType: 0
                    },
                    landMark: [],
                    latlng: "",
                    location: "",
                    noteid: noteId,
                    remindtime: "",
                    remindtype: 0,
                    revision: "1",
                    system: "",
                    tags: tags,
                    title: title,
                    topmost: "0",
                    updatetime: String(Date.now()),
                    userphone: String(account),
                    version: "",
                    visitTime: String(Date.now())
                }, {
                    headers: {
                        APP_CP: "pc",
                        APP_NUMBER: String(account),
                        CP_VERSION: "7.7.1.20240115",
                        ...headers
                    }
                });
            },
            deleteNote: function(noteid, headers) {
                return http.post(`${mnoteUrl}/noteServer/api/moveToRecycleBin.do`, {
                    noteids: [
                        {
                            noteid: noteid
                        }
                    ]
                }, {
                    headers: {
                        APP_CP: "pc",
                        CP_VERSION: "7.7.1.20240115",
                        ...headers
                    }
                });
            },
            tyrzLogin: function(ssoToken) {
                return http.get(`${caiyunUrl2}/portal/auth/tyrzLogin.action?ssoToken=${ssoToken}`);
            },
            signInInfo: function() {
                return http.get(`${caiyunUrl2}/market/signin/page/info?client=app`);
            },
            getDrawInWx: function() {
                return http.get(`${caiyunUrl2}/market/playoffic/drawInfo`);
            },
            drawInWx: function() {
                return http.get(`${caiyunUrl2}/market/playoffic/draw`);
            },
            signInfoInWx: function() {
                return http.get(`${caiyunUrl2}/market/playoffic/followSignInfo?isWx=true`);
            },
            getDisk: function(account, catalogID) {
                return http.post(`${yun139Url}/orchestration/personalCloud/catalog/v1.0/getDisk`, {
                    commonAccountInfo: {
                        account: String(account)
                    },
                    catalogID: catalogID,
                    catalogType: -1,
                    sortDirection: 1,
                    catalogSortType: 0,
                    contentSortType: 0,
                    filterType: 1,
                    startNumber: 1,
                    endNumber: 40
                });
            },
            queryBatchList: function() {
                return http.post("https://grdt.middle.yun.139.com/openapi/pDynamicInfo/queryBatchList", {
                    encodeData: "WBvKN8KKSLovAM=",
                    encodeType: 2,
                    pageSize: 3,
                    dynamicType: 2
                });
            },
            uploadFileRequest: function(options) {
                return http.post("https://ose.caiyun.feixin.10086.cn/richlifeApp/devapp/IUploadAndDownload", getUploadXml(options), {
                    headers: {
                        // 'hcy-cool-flag': '1',
                        "x-huawei-uploadSrc": "2",
                        "x-huawei-channelSrc": "10000023",
                        "Content-Type": "text/xml; charset=UTF-8"
                    }
                });
            },
            uploadFile: function(url, id, file, size) {
                return http.post(url, file, {
                    headers: {
                        UploadtaskID: id + "-",
                        "x-huawei-uploadSrc": "1",
                        "Content-Type": "application/octet-stream",
                        "x-huawei-channelSrc": "10000023",
                        "User-Agent": "okhttp/3.11.0",
                        contentSize: size.toString(),
                        Range: `bytes=0-${(size - 1).toString()}`
                    }
                });
            },
            pcUploadFileRequest: function(account, parentCatalogID, contentSize, contentName, digest) {
                return http.post(`${yun139Url}/orchestration/personalCloud/uploadAndDownload/v1.0/pcUploadFileRequest`, {
                    commonAccountInfo: {
                        account: String(account)
                    },
                    fileCount: 1,
                    totalSize: contentSize,
                    uploadContentList: [
                        {
                            contentName: contentName,
                            contentSize: contentSize,
                            comlexFlag: 0,
                            digest: digest
                        }
                    ],
                    newCatalogName: "",
                    parentCatalogID: parentCatalogID,
                    operation: 0,
                    path: "",
                    manualRename: 2,
                    autoCreatePath: [],
                    tagID: "",
                    tagType: "",
                    seqNo: ""
                });
            },
            createBatchOprTask: function(account, contentIds) {
                return http.post(`${yun139Url}/orchestration/personalCloud/batchOprTask/v1.0/createBatchOprTask`, {
                    createBatchOprTaskReq: {
                        taskType: 2,
                        actionType: 201,
                        taskInfo: {
                            contentInfoList: contentIds,
                            catalogInfoList: [],
                            newCatalogID: ""
                        },
                        commonAccountInfo: {
                            account: account,
                            accountType: 1
                        }
                    }
                });
            },
            queryBatchOprTaskDetail: function(account, taskID) {
                return http.post(`${yun139Url}/orchestration/personalCloud/batchOprTask/v1.0/queryBatchOprTaskDetail`, {
                    queryBatchOprTaskDetailReq: {
                        taskID: taskID,
                        commonAccountInfo: {
                            account: account,
                            accountType: 1
                        }
                    }
                });
            },
            clickTask: function(id) {
                return http.get(`${caiyunUrl2}/market/signin/task/click?key=task&id=${id}`);
            },
            getTaskList: function(marketname = "sign_in_3") {
                return http.get(`${caiyunUrl2}/market/signin/task/taskList?marketname=${marketname}&clientVersion=`);
            },
            receive: function() {
                return http.get(`${caiyunUrl2}/market/signin/page/receive`);
            },
            shake: function() {
                return http.post(`${caiyunUrl2}/market/shake-server/shake/shakeIt?flag=1`);
            },
            beinviteHecheng1T: function() {
                return http.get(`${caiyunUrl2}/market/signin/hecheng1T/beinvite`);
            },
            finishHecheng1T: function() {
                return http.get(`${caiyunUrl2}/market/signin/hecheng1T/finish?flag=true`);
            },
            getOutLink: function(account, coIDLst, dedicatedName) {
                return http.post(`${yun139Url}/orchestration/personalCloud-rebuild/outlink/v1.0/getOutLink`, {
                    getOutLinkReq: {
                        subLinkType: 0,
                        encrypt: 1,
                        coIDLst: coIDLst,
                        caIDLst: [],
                        pubType: 1,
                        dedicatedName: dedicatedName,
                        period: 1,
                        periodUnit: 1,
                        viewerLst: [],
                        extInfo: {
                            isWatermark: 0,
                            shareChannel: "3001"
                        },
                        commonAccountInfo: {
                            account: account,
                            accountType: 1
                        }
                    }
                });
            },
            getBlindboxTask: function() {
                return http.post(`${caiyunUrl2}/market/task-service/task/api/blindBox/queryTaskInfo`, {
                    marketName: "National_BlindBox",
                    clientType: 1
                }, {
                    headers: {
                        accept: "application/json"
                    }
                });
            },
            registerBlindboxTask: function(taskId) {
                return http.post(`${caiyunUrl2}/market/task-service/task/api/blindBox/register`, {
                    marketName: "National_BlindBox",
                    taskId: taskId
                }, {
                    headers: {
                        accept: "application/json"
                    }
                });
            },
            blindboxUser: function() {
                return http.post(`${caiyunUrl2}/ycloud/blindbox/user/info`, {
                    from: "main"
                }, {
                    headers: {
                        accept: "application/json"
                    }
                });
            },
            openBlindbox: function() {
                return http.post(`${caiyunUrl2}/ycloud/blindbox/draw/openBox?from=main`, {
                    from: "main"
                }, {
                    headers: {
                        accept: "application/json",
                        "x-requested-with": "cn.cj.pe",
                        referer: "https://caiyun.feixin.10086.cn:7071/portal/caiyunOfficialAccount/index.html?path=blindBox&sourceid=1016",
                        origin: "https://caiyun.feixin.10086.cn",
                        "user-agent": "Mozilla/5.0 (Linux; Android 10; Redmi K20 Pro Build/QKQ1.190828.002; wv) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/83.0.4103.106 Mobile Safari/537.36(139PE_WebView_Android_10.2.2_mcloud139)"
                    }
                });
            },
            datacenter: function(base64) {
                return http.post("https://datacenter.mail.10086.cn/datacenter/", `data=${base64}&ext=${"crc=" + hashCode(base64)}`, {
                    headers: {
                        "content-type": "application/x-www-form-urlencoded",
                        platform: "h5"
                    }
                });
            },
            /**
       * 该接口用于测试
       */ // post(url:string,data:Record<string,any>|string,headers:Record<string,any>){
            //   return http.post(url,data,{
            //     headers
            //   })
            // },
            /**
       * 登记
       * @param other 其它参数，& 开头
       */ journaling: function(optkeyword, sourceid = 1010, other = "") {
                return http.post(`${caiyunUrl2}/portal/journaling`, `account=&module=uservisit&optkeyword=${optkeyword}&fromId=&flag=&fileId=&fileType=&fileExtname=&fileSize=&sourceid=${sourceid}&linkId=${other}`, {
                    headers: {
                        "content-type": "application/x-www-form-urlencoded"
                    }
                });
            },
            getCloudRecord: function(pn = 1, ps = 10, type = 1) {
                return http.get(`${caiyunUrl2}/market/signin/public/cloudRecord?type=${type}&pageNumber=${pn}&pageSize=${ps}`);
            },
            loginMail: function(token) {
                return http.post("https://mail.10086.cn/login/inlogin.action", `<?xml version="1.0" encoding="utf-8"?>
      <object>
       <string name="clientId">10804</string> 
       <string name="version">9</string>
       <string name="loginType">7</string> 
       <string name="token">${token}</string> 
      </object>`);
            },
            ...createSignInApi(http),
            ...createMsgPushApi(http),
            ...createBackupGiftApi(http),
            ...createAiRedPackApi(http),
            ...createMailChatApi(http)
        };
    }
    function getUploadXml({ phone, manualRename = 2, parentCatalogID, createTime: createTime2, digest, contentName, contentSize }) {
        return `<pcUploadFileRequest>
  <ownerMSISDN>${phone}</ownerMSISDN>
  <fileCount>1</fileCount>
  <totalSize>${contentSize}</totalSize>
  <uploadContentList length="1">
     <uploadContentInfo>
        <comlexFlag>0</comlexFlag>
        <contentDesc><![CDATA[]]></contentDesc>
        <contentName><![CDATA[${contentName}]]></contentName>
        <contentSize>${contentSize}</contentSize>
        <contentTAGList></contentTAGList>
        <digest>${digest}</digest>
        <exif>
           <createTime>${createTime2}</createTime>
        </exif>
        <fileEtag>0</fileEtag>
        <fileVersion>0</fileVersion>
        <updateContentID></updateContentID>
     </uploadContentInfo>
  </uploadContentList>
  <newCatalogName></newCatalogName>
  <parentCatalogID>${parentCatalogID}</parentCatalogID>
  <operation>0</operation>
  <path></path>
  <manualRename>${manualRename}</manualRename>
  <autoCreatePath length="0"/>
  <tagID></tagID>
  <tagType></tagType>
</pcUploadFileRequest>`;
    }
    // ../../core/caiyun/index.ts
    function getSsoTokenApi($, phone) {
        try {
            let specToken = $.api.querySpecToken(phone);
            if (!specToken.success) {
                // $.logger.error("\u83B7\u53D6 ssoToken \u5931\u8D25", specToken.message);
                return;
            }
            return specToken.data.token;
        } catch (error) {
            // $.logger.error("\u83B7\u53D6 ssoToken \u5F02\u5E38", error.message);
        }
    }
    function loginEmail($) {
        let ssoToken = getSsoTokenApi($, $.config.phone);
        if (ssoToken) try {
            let { code, summary, var: data } = $.api.loginMail(ssoToken);
            if (code !== "S_OK") {
                $.logger.error("\u83B7\u53D6 sid \u5931\u8D25", code, summary);
                return;
            }
            return data;
        } catch (error) {
            $.logger.error("\u83B7\u53D6 sid \u5F02\u5E38", error.message);
        }
    }
    function getJwtTokenApi($, ssoToken) {
        return request($, $.api.tyrzLogin, "\u83B7\u53D6 ssoToken ", ssoToken).token;
    }
    function signInApi($) {
        return request($, $.api.signInInfo, "\u7F51\u76D8\u7B7E\u5230");
    }
    function signInWxApi($) {
        return request($, $.api.signInfoInWx, "\u5FAE\u4FE1\u7B7E\u5230");
    }
    function getJwtToken($) {
        let ssoToken = getSsoTokenApi($, $.config.phone);
        if (ssoToken) return getJwtTokenApi($, ssoToken);
    }
    function refreshToken($) {
        try {
            let { token, phone } = $.config, tokenXml = $.api.authTokenRefresh(token, phone);
            return tokenXml ? getXmlElement(tokenXml, "token") : $.logger.error("authTokenRefresh \u5931\u8D25");
        } catch (error) {
            $.logger.error("\u5237\u65B0 token \u5931\u8D25", error.message);
        }
    }
    function signIn($) {
        let { todaySignIn, total, toReceive } = signInApi($) || {};
        if ($.logger.info(`\u5F53\u524D\u79EF\u5206${total}${toReceive ? `\uFF0C\u5F85\u9886\u53D6${toReceive}` : ""}`), todaySignIn === !0) {
            $.logger.info("\u7F51\u76D8\u4ECA\u65E5\u5DF2\u7B7E\u5230");
            return;
        }
        $.sleep(1e3);
        let info = signInApi($);
        if (info) {
            if (info.todaySignIn === !1) {
                $.logger.info("\u7F51\u76D8\u7B7E\u5230\u5931\u8D25");
                return;
            }
            $.logger.info("\u7F51\u76D8\u7B7E\u5230\u6210\u529F");
        }
    }
    function signInWx($) {
        let info = signInWxApi($);
        if (info) {
            if (info.todaySignIn === !1 && ($.logger.info("\u5FAE\u4FE1\u7B7E\u5230\u5931\u8D25"), info.isFollow === !1)) {
                // $.logger.info("\u5F53\u524D\u8D26\u53F7\u6CA1\u6709\u7ED1\u5B9A\u5FAE\u4FE1\u516C\u4F17\u53F7\u3010\u4E2D\u56FD\u79FB\u52A8\u4E91\u76D8\u3011");
                $.logger.info("\u5f53\u524d\u8d26\u53f7\u6ca1\u6709\u7ed1\u5b9a\u5fae\u4fe1\u516c\u4f17\u53f7");
                
                return;
            }
            $.logger.info("\u5FAE\u4FE1\u7B7E\u5230\u6210\u529F");
        }
    }
    function wxDraw($) {
        try {
            let drawInfo = $.api.getDrawInWx();
            if (drawInfo.code !== 0) {
                $.logger.error(`\u83B7\u53D6\u5FAE\u4FE1\u62BD\u5956\u4FE1\u606F\u5931\u8D25\uFF0C\u8DF3\u8FC7\u8FD0\u884C\uFF0C${JSON.stringify(drawInfo)}`);
                return;
            }
            if (drawInfo.result.surplusNumber < 50) {
                $.logger.info(`\u5269\u4F59\u5FAE\u4FE1\u62BD\u5956\u6B21\u6570${drawInfo.result.surplusNumber}\uFF0C\u8DF3\u8FC7\u6267\u884C`);
                return;
            }
            let draw = $.api.drawInWx();
            if (draw.code !== 0) {
                $.logger.error(`\u5FAE\u4FE1\u62BD\u5956\u5931\u8D25\uFF0C${JSON.stringify(draw)}`);
                return;
            }
            $.logger.info(`\u5FAE\u4FE1\u62BD\u5956\u6210\u529F\uFF0C\u83B7\u5F97\u3010${draw.result.prizeName}\u3011`);
        } catch (error) {
            $.logger.error("\u5FAE\u4FE1\u62BD\u5956\u5F02\u5E38", error.message);
        }
    }
    function receive($) {
        return request($, $.api.receive, "\u9886\u53D6\u4E91\u6735");
    }
    function clickTask($, task) {
        try {
            let { code, msg } = $.api.clickTask(task);
            if (code === 0) return !0;
            $.logger.info(`\u70B9\u51FB\u4EFB\u52A1${task}\u5931\u8D25`, msg);
        } catch (error) {
            $.logger.error(`\u70B9\u51FB\u4EFB\u52A1${task}\u5F02\u5E38`, error.message);
        }
        return !1;
    }
    function deleteFiles($, ids) {
        try {
            $.logger.debug(`\u5220\u9664\u6587\u4EF6${ids.join(",")}`);
            let { data: { createBatchOprTaskRes: { taskID } } } = $.api.createBatchOprTask($.config.phone, ids);
            $.api.queryBatchOprTaskDetail($.config.phone, taskID);
        } catch (error) {
            $.logger.error("\u5220\u9664\u6587\u4EF6\u5931\u8D25", error.message);
        }
    }
    function getNoteAuthToken($) {
        try {
            return $.api.getNoteAuthToken($.config.auth, $.config.phone);
        } catch (error) {
            $.logger.error("\u83B7\u53D6\u4E91\u7B14\u8BB0 Auth Token \u5F02\u5E38", error.message);
        }
    }
    function uploadFileDaily($) {
        let contentID = pcUploadFileRequest($, getParentCatalogID());
        contentID && setStoreArray($.store, "files", contentID);
    }
    function createNoteDaily($) {
        if (!$.config.auth) {
            $.logger.info("\u672A\u914D\u7F6E authToken\uFF0C\u8DF3\u8FC7\u4E91\u7B14\u8BB0\u4EFB\u52A1\u6267\u884C");
            return;
        }
        let headers = getNoteAuthToken($);
        if (!headers) {
            $.logger.info("\u83B7\u53D6\u9274\u6743\u4FE1\u606F\u5931\u8D25\uFF0C\u8DF3\u8FC7\u4E91\u7B14\u8BB0\u4EFB\u52A1\u6267\u884C");
            return;
        }
        try {
            let id = randomHex(32);
            $.api.createNote(id, `${randomHex(3)}`, $.config.phone, headers), $.sleep(2e3), $.api.deleteNote(id, headers);
        } catch (error) {
            $.logger.error("\u521B\u5EFA\u4E91\u7B14\u8BB0\u5F02\u5E38", error.message);
        }
    }
    function _clickTask($, id, currstep) {
        let idCurrstepMap = {
            434: 22
        };
        return idCurrstepMap[id] && currstep === idCurrstepMap[id] ? (clickTask($, id), !0) : currstep === 0 ? clickTask($, id) : !0;
    }
    function shareTime($) {
        try {
            let files = $.store.files;
            if (!files || !files[0]) {
                $.logger.info("\u672A\u83B7\u53D6\u5230\u6587\u4EF6\u5217\u8868\uFF0C\u8DF3\u8FC7\u5206\u4EAB\u4EFB\u52A1");
                return;
            }
            let { code, message } = $.api.getOutLink($.config.phone, [
                files[0]
            ], "");
            if (code === "0") return $.logger.info("\u5206\u4EAB\u94FE\u63A5\u6210\u529F"), !0;
            $.logger.info("\u5206\u4EAB\u94FE\u63A5\u5931\u8D25", code, message);
        } catch (error) {
            $.logger.error("\u5206\u4EAB\u94FE\u63A5\u5F02\u5E38", error.message);
        }
    }
    function getAppTaskList($, marketname = "sign_in_3") {
        let { month = [], day = [], time = [], new: new_ = [] } = request($, $.api.getTaskList, "\u83B7\u53D6\u4EFB\u52A1\u5217\u8868", marketname);
        return [
            ...month,
            ...day,
            ...time,
            ...new_
        ];
    }
    function getAllAppTaskList($) {
        let list1 = getAppTaskList($, "sign_in_3"), list2 = getAppTaskList($, "newsign_139mail");
        return list1.concat(list2);
    }
    function getTaskRunner($) {
        return {
            113: refreshToken,
            106: uploadFileDaily,
            107: createNoteDaily,
            434: shareTime,
            110: $.node && $.node.uploadTask
        };
    }
    function appTask($) {
        //  ------【任务列表】------
        // $.logger.info("------\u3010\u4EFB\u52A1\u5217\u8868\u3011------");
        let taskList = getAllAppTaskList($), taskRunner = getTaskRunner($), doingList = [];
        taskList.sort((a, b)=>a.id - b.id);
        for (let task of taskList)if (!(task.state === "FINISH" || task.enable !== 1)) if (TASK_LIST[task.id]) {
            var _taskRunner_task_id, _taskRunner_task_id1;
            if (task.marketname === "sign_in_3" && task.groupid === "month" && $.store.curMonthBackup === !1 && /* @__PURE__ */ new Date().getDate() < 20) {
                $.logger.debug("\u8DF3\u8FC7\u8FC7\u4EFB\u52A1\uFF08\u672A\u5F00\u542F\u5907\u4EFD\uFF09", task.name);
                continue;
            }
            _clickTask($, task.id, task.currstep) && (task.id === 110 ? (_taskRunner_task_id = taskRunner[task.id]) === null || _taskRunner_task_id === void 0 ? void 0 : _taskRunner_task_id.call(taskRunner, $, task.process) : (_taskRunner_task_id1 = taskRunner[task.id]) === null || _taskRunner_task_id1 === void 0 ? void 0 : _taskRunner_task_id1.call(taskRunner, $), doingList.push(task.id), $.sleep(500));
        } else SKIP_TASK_LIST.includes(task.id) || clickTask($, task.id);
        let skipCheck = [
            434
        ];
        if (doingList.length) for (let task of getAllAppTaskList($))doingList.includes(task.id) ? task.state === "FINISH" ? $.logger.info("\u6210\u529F", task.name) : !skipCheck.includes(task.id) && $.logger.info("\u5931\u8D25", task.name, "\u8BF7\u624B\u52A8\u5B8C\u6210") : (task.groupid === "month" || task.groupid === "day") && task.state !== "FINISH" && $.logger.info("\u672A\u5B8C\u6210", task.name, "\u8BF7\u624B\u52A8\u5B8C\u6210");
    }
    function shake($) {
        let { shakePrizeconfig, shakeRecommend } = request($, $.api.shake, "\u6447\u4E00\u6447");
        if (shakeRecommend) return $.logger.debug(shakeRecommend.explain || shakeRecommend.img);
        if (shakePrizeconfig) return $.logger.info(shakePrizeconfig.title + shakePrizeconfig.name);
    }
    function shakeTask($) {
        $.logger.info("------\u3010\u6447\u4E00\u6447\u3011------");
        let { delay, num } = $.config.shake;
        for(let index = 0; index < num; index++)shake($), index < num - 1 && $.sleep(delay * 1e3);
    }
    function shareFind($) {
        let phone = $.config.phone;
        try {
            let data = {
                traceId: Number(Math.random().toString().substring(10)),
                tackTime: Date.now(),
                distinctId: randomHex([
                    14,
                    15,
                    8,
                    7,
                    15
                ]),
                eventName: "discoverNewVersion.Page.Share.QQ",
                event: "$manual",
                flushTime: Date.now(),
                model: "",
                osVersion: "",
                appVersion: "",
                manufacture: "",
                screenHeight: 895,
                os: "Android",
                screenWidth: 393,
                lib: "js",
                libVersion: "1.17.2",
                networkType: "",
                resumeFromBackground: "",
                screenName: "",
                title: "\u3010\u7CBE\u9009\u3011\u4E00\u7AD9\u5F0F\u8D44\u6E90\u5B9D\u5E93",
                eventDuration: "",
                elementPosition: "",
                elementId: "",
                elementContent: "",
                elementType: "",
                downloadChannel: "",
                crashedReason: "",
                phoneNumber: phone,
                storageTime: "",
                channel: "",
                activityName: "",
                platform: "h5",
                sdkVersion: "1.0.1",
                elementSelector: "",
                referrer: "",
                scene: "",
                latestScene: "",
                source: "content-open",
                urlPath: "",
                IP: "",
                url: `https://h.139.com/content/discoverNewVersion?columnId=20&token=STuid00000${Date.now()}${randomHex(20)}&targetSourceId=001005`,
                elementName: "",
                browser: "Chrome WebView",
                elementTargetUrl: "",
                referrerHost: "",
                browerVersion: "122.0.6261.106",
                latitude: "",
                pageDuration: "",
                longtitude: "",
                urlQuery: "",
                shareDepth: "",
                arriveTimeStamp: "",
                spare: {
                    mobile: phone,
                    channel: ""
                },
                public: "",
                province: "",
                city: "",
                carrier: ""
            };
            $.api.datacenter(Buffer.from(JSON.stringify(data)).toString("base64"));
        } catch (error) {
            $.logger.error("\u5206\u4EAB\u6709\u5956\u5F02\u5E38", error.message);
        }
    }
    function getCloudRecord($) {
        return request($, $.api.getCloudRecord, "\u83B7\u53D6\u4E91\u6735\u8BB0\u5F55");
    }
    function getShareFindCount($) {
        if (!$.localStorage.shareFind) return 20;
        let { lastUpdate, count } = $.localStorage.shareFind;
        return /* @__PURE__ */ new Date().getMonth() === new Date(lastUpdate).getMonth() ? 20 - count : 20;
    }
    function shareFindTask($) {
        $.logger.info("------\u3010\u9080\u8BF7\u597D\u53CB\u770B\u7535\u5F71\u3011------"), $.logger.info("\u6D4B\u8BD5\u4E2D\u3002\u3002\u3002");
        let count = getShareFindCount($);
        if (count <= 0) {
            $.logger.info("\u672C\u6708\u5DF2\u5206\u4EAB");
            return;
        }
        let _count = 20 - --count;
        shareFind($), $.sleep(1e3), receive($), $.sleep(1e3);
        let { records } = getCloudRecord($), recordFirst = records === null || records === void 0 ? void 0 : records.find((record)=>record.mark === "fxnrplus5");
        if (recordFirst && /* @__PURE__ */ new Date().getTime() - new Date(recordFirst.updatetime).getTime() < 2e4) {
            for(; count > 0;)_count++, count--, $.logger.debug("\u9080\u8BF7\u597D\u53CB"), shareFind($), $.sleep(2e3);
            receive($);
            let { records: records2 } = getCloudRecord($);
            (records2 === null || records2 === void 0 ? void 0 : records2.filter((record)=>record.mark === "fxnrplus5").length) > 6 ? $.logger.info("\u5B8C\u6210") : $.logger.error("\u672A\u77E5\u60C5\u51B5\uFF0C\u65E0\u6CD5\u5B8C\u6210\uFF08\u6216\u5DF2\u5B8C\u6210\uFF09\uFF0C\u4ECA\u65E5\u8DF3\u8FC7");
        } else $.logger.error("\u672A\u77E5\u60C5\u51B5\uFF0C\u65E0\u6CD5\u5B8C\u6210\uFF08\u6216\u5DF2\u5B8C\u6210\uFF09\uFF0C\u672C\u6B21\u8DF3\u8FC7"), _count += 10;
        $.localStorage.shareFind = {
            lastUpdate: /* @__PURE__ */ new Date().getTime(),
            count: _count
        };
    }
    function openBlindbox($) {
        try {
            $.logger.debug("\u5F00\u76F2\u76D2");
            let { code, msg, result } = $.api.openBlindbox();
            switch(code){
                case 0:
                    return $.logger.info("\u83B7\u5F97", result.prizeName);
                case 200105:
                    return $.logger.debug("\u4EC0\u4E48\u90FD\u6CA1\u6709\u54E6");
                case 200106:
                    return $.logger.error("\u5F02\u5E38", code, msg);
                default:
                    return $.logger.warn("\u672A\u77E5\u539F\u56E0\u5931\u8D25", code, msg);
            }
        } catch (error) {
            $.logger.error("openBlindbox \u5F02\u5E38", error.message);
        }
    }
    function openBlindboxAfterGetCount($) {
        try {
            let { chanceNum } = request($, $.api.blindboxUser, "\u83B7\u53D6\u76F2\u76D2\u4EFB\u52A1");
            if (!chanceNum) return;
            $.sleep(666);
            for(let i = 0; i < chanceNum; i++)openBlindbox($), $.sleep(666);
        } catch (error) {
            $.logger.error("\u5F00\u76D2\u5F02\u5E38", error.message);
        }
    }
    function registerBlindboxTask($, taskId) {
        request($, $.api.registerBlindboxTask, "\u6CE8\u518C\u76F2\u76D2", taskId);
    }
    function openMoreBlindbox($) {
        try {
            let taskList = request($, $.api.getBlindboxTask, "\u83B7\u53D6\u76F2\u76D2\u4EFB\u52A1");
            if (!Array.isArray(taskList)) return;
            let tasks = taskList.filter((task)=>task.memo && !task.memo.includes("isLimit") && task.status === 0);
            if (tasks.length <= 0) return;
            for (let { taskName, taskId } of tasks)$.logger.debug("\u6CE8\u518C\u76F2\u76D2\u4EFB\u52A1", taskName), registerBlindboxTask($, taskId), $.sleep(666), openBlindboxAfterGetCount($);
        } catch (error) {
            $.logger.error(error);
        }
    }
    function blindboxJournaling2({ api, sleep }) {
        api.journaling("National_BlindBox_userLogin"), sleep(200), api.journaling("National_BlindBox_login"), sleep(200), api.journaling("National_BlindBox_loginAppOuterEnd"), sleep(200);
    }
    function blindboxTask($) {
        $.logger.info("------\u3010\u5F00\u76F2\u76D2\u3011------");
        try {
            blindboxJournaling2($);
            let r1 = request($, $.api.blindboxUser, "\u83B7\u53D6\u76F2\u76D2\u7528\u6237\u4FE1\u606F");
            if (typeof r1.chanceNum != "number") return openBlindbox($);
            if (r1.chanceNum === 0 && r1.taskNum >= 2) {
                $.logger.info("\u4ECA\u65E5\u5DF2\u5B8C\u6210");
                return;
            }
            r1.firstTime && $.logger.info("\u4ECA\u65E5\u9996\u6B21\u767B\u5F55\uFF0C\u83B7\u53D6\u6B21\u6570 +1");
            for(let i = 0; i < r1.chanceNum; i++)openBlindbox($), $.sleep(666);
            openMoreBlindbox($);
        } catch (error) {
            $.logger.error("\u5F00\u76F2\u76D2\u4EFB\u52A1\u5F02\u5E38", error.message);
        }
    }
    function checkHc1T({ localStorage }) {
        if (localStorage.hc1T) {
            let { lastUpdate } = localStorage.hc1T;
            if (/* @__PURE__ */ new Date().getMonth() <= new Date(lastUpdate).getMonth()) return !0;
        }
    }
    function hc1Task($) {
        if ($.logger.info("------\u3010\u5408\u6210\u829D\u9EBB\u3011------"), checkHc1T($)) {
            $.logger.info("\u672C\u6708\u5DF2\u9886\u53D6");
            return;
        }
        try {
            request($, $.api.beinviteHecheng1T, "\u5408\u6210\u829D\u9EBB"), $.sleep(5e3), request($, $.api.finishHecheng1T, "\u5408\u6210\u829D\u9EBB"), $.logger.info("\u5B8C\u6210\u5408\u6210\u829D\u9EBB"), $.localStorage.hc1T = {
                lastUpdate: /* @__PURE__ */ new Date().getTime()
            };
        } catch (error) {
            $.logger.error("\u5408\u6210\u829D\u9EBB\u5931\u8D25", error.message);
        }
    }
    function afterTask($) {
        $.logger.info("------\u3010\u643D\u5C41\u80A1\u3011------");
        try {
            $.store && $.store.files && deleteFiles($, $.store.files);
        } catch (error) {
            $.logger.error("afterTask \u5F02\u5E38", error.message);
        }
    }
    function run($) {
        let { config } = $, taskList = [
            signIn,
            taskExpansionTask,
            signInWx,
            wxDraw,
            appTask,
            shareFindTask,
            blindboxTask,
            hc1Task,
            aiRedPackTask,
            shakeTask,
            receive,
            msgPushOnTask,
            backupGiftTask
        ];
        config && config.garden && config.garden.enable && taskList.push(gardenTask);
        for (let task of taskList)task($), $.sleep(1e3);
        afterTask($);
    }
    function getTokenExpireTime(token) {
        return Number(token.split("|")[3]);
    }
    function isNeedRefresh(expireTime) {
        return expireTime - Date.now() < 432e6;
    }
    function createNewAuth($) {
        let config = $.config, expireTime = getTokenExpireTime(config.token);
        if ($.logger.debug("------\u3010\u68C0\u6D4B\u8D26\u53F7\u6709\u6548\u671F\u3011------"), $.logger.debug(`token \u6709\u6548\u671F ${Math.floor((expireTime - Date.now()) / 864e5)} \u5929`), !isNeedRefresh(expireTime)) return;
        $.logger.info("\u5C1D\u8BD5\u751F\u6210\u65B0\u7684 auth");
        let token = refreshToken($);
        if (token) return Buffer.from(// @ts-ignore
        `${config.platform}:${config.phone}:${token}`).toString("base64");
        $.logger.error("\u751F\u6210\u65B0 auth \u5931\u8D25");
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
//     function serverChan(apiOption, { token, ...option }, title, text) {
//         return _send(apiOption, "Server\u9171", {
//             url: `https://sctapi.ftqq.com/${token}.send`,
//             headers: {
//                 "content-type": "application/x-www-form-urlencoded"
//             },
//             data: {
//                 text: title,
//                 desp: text.replace(/\n/g, `

// `),
//                 ...option
//             }
//         });
//     }
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
        let request3 = (options2)=>{
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
            request: request3,
            get: (url, options2 = {})=>request3({
                    ...options2,
                    method: "GET",
                    url: url
                }),
            post: (url, body, options2 = {})=>request3({
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
    function main(index, config, option) {
        if (config = {
            ...config,
            ...getAuthInfo(config.auth)
        }, config.phone.length !== 11 || !config.phone.startsWith("1")) {
            console.info("auth \u683C\u5F0F\u89E3\u6790\u9519\u8BEF\uFF0C\u8BF7\u67E5\u770B\u662F\u5426\u586B\u5199\u6B63\u786E\u7684 auth");
            return;
        }
        let cookieJar = createCookieJar(), logger = createLogger({
            pushData: option && option.pushData
        }), DATA = {
            baseUA: "Mozilla/5.0 (Linux; Android 13; 22041216C Build/TP1A.220624.014; wv) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/121.0.6167.178 Mobile Safari/537.36",
            mailUaEnd: "(139PE_WebView_Android_10.2.2_mcloud139)",
            mailRequested: "cn.cj.pe",
            mcloudRequested: "com.chinamobile.mcloud"
        };
        // logger.info("--------------"), 
        logger.info(`\u4F60\u597D\uFF1A${config.phone}`);
        let jwtToken, headers = {
            "user-agent": DATA.baseUA,
            "x-requested-with": DATA.mcloudRequested,
            charset: "utf-8",
            "content-type": "application/json",
            accept: "application/json"
        };
        function getHeaders(url) {
            return getHostname(url) === "caiyun.feixin.10086.cn" && jwtToken ? {
                ...headers,
                cookie: cookieJar.getCookieString(),
                jwttoken: jwtToken
            } : {
                ...headers,
                authorization: config.auth
            };
        }
        let http = createRequest({
            hooks: {
                beforeRequest: function(options) {
                    return options.headers = getHeaders(options.url), options;
                }
            }
        }), $ = {
            api: createApi(http),
            logger: logger,
            DATA: DATA,
            sleep: Time.sleep,
            config: {
                shake: {
                    num: 15,
                    delay: 2
                },
                ...config
            },
            gardenApi: createGardenApi(http),
            store: {},
            localStorage: {},
            md5: md5
        };
        if (jwtToken = getJwtToken($), !!jwtToken) return run($), createNewAuth($);
    }
    // var sheet = Application.Sheets.Item("\u79FB\u52A8\u4E91\u76D8") || Application.Sheets.Item("caiyun") || ActiveSheet, usedRange = sheet.UsedRange, AColumn = sheet.Columns("A"), len = usedRange.Row + usedRange.Rows.Count - 1, BColumn = sheet.Columns("B"), pushData = [];
    // for(let i = 1; i <= len; i++){
    //     let cell = AColumn.Rows(i);
    //     cell.Text && (console.log(`\u6267\u884C\u7B2C ${i} \u884C`), runMain(i, cell), console.log(`\u7B2C ${i} \u884C\u6267\u884C\u7ED3\u675F`));
    // }
    // sendWpsNotify(pushData, getPushConfig());
    function runMain(i, cell) {
        try {
            let newAuth = main(i, {
                // auth: cell.Text.length === 11 ? BColumn.Rows(i).Text : cell.Text
                auth: cell
            }, {
                pushData: pushData
            });
            // newAuth && (console.log("🍳 \u66F4\u65B0 auth \u6210\u529F"), cell.Text.length === 11 ? BColumn.Rows(i).Value = newAuth : AColumn.Rows(i).Value = newAuth);
            newAuth
        } catch (error) {
            // console.log(error.message);
        }
    }
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
  messageHeader[posLabel] = "👨‍🚀" + messageName
  // try {

    // var sheet = Application.Sheets.Item("\u79FB\u52A8\u4E91\u76D8") || Application.Sheets.Item("caiyun") || ActiveSheet, usedRange = sheet.UsedRange, AColumn = sheet.Columns("A"), len = usedRange.Row + usedRange.Rows.Count - 1, BColumn = sheet.Columns("B"), pushData = [];
    // for(let i = 1; i <= len; i++){
    //     let cell = AColumn.Rows(i);
    //     cell.Text && (console.log(`\u6267\u884C\u7B2C ${i} \u884C`), runMain(i, cell), console.log(`\u7B2C ${i} \u884C\u6267\u884C\u7ED3\u675F`));
    // }
    // sendWpsNotify(pushData, getPushConfig());

    cell = cookie
    i = 0

    runMain(i, cell)
    // console.log(pushData)
    msg = ""
    errormsg = ""
    countInfo = 0 // 计算info次数，如果为一次或0次说明失败
    countError = 0  // 计算error次数
    for(i = 0; i<pushData.length; i++)
    {
      // console.log(pushData[i]["msg"])
      try{
        type = pushData[i]["type"]
        if(type == "info")
        {
          if(countInfo == 0){
            msg += "" + pushData[i]["msg"] + ""
          }else{
            msg += "\n🎉 " + pushData[i]["msg"] + ""
          }
          countInfo += 1
        }else
        {
          errormsg += "\n📢 " + pushData[i]["msg"] + ""
          // console.log(pushData[i]["msg"])
        }
        
      }catch{

      }
      
    }
    // console.log(msg)
    if(countInfo <= 1){
      messageSuccess += ""
      if(countInfo == 1){ // 有一个说明是用户的电话
        messageFail += "❌ " +  msg + "失败"
      }
    }else{
      messageSuccess += "" + msg + ""
      messageFail += "" + errormsg + ""
    }
    
    
    
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

