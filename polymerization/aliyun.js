// 阿里云盘(自动更新token版)、已移除自动领奖功能
// 20240706
// 文中引用代码改编自https://www.52pojie.cn/thread-1869673-43-1.html

let sheetNameSubConfig = "aliyun"; // 分配置表名称
let pushHeader = "【阿里云盘】";
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
function sleep(d) {
    for (var t = Date.now(); Date.now() - t <= d;); // 使程序暂停执行一段时间
}
 
function log(message) {
    console.log(message); // 打印消息到控制台
    // TODO: 将日志写入文件
}
 

function doTask(row){
    var myDate = new Date(); // 创建一个表示当前时间的 Date 对象
    var data_time = myDate.toLocaleDateString(); // 获取当前日期的字符串表示
    var date = `${new Date().getMonth() + 1}-${new Date().getDate().toString().padStart(2, '0')}`
    
    //需要修改的地方
    var dengdai = 2
    //意思是 等待多少分钟去登录 等待多少分钟去签到  默认1-15分钟之间 改了15就是1-X分钟之间 取随机的
    var renwu = 2   //建议定时2个任务 就写1  定时3个任务就写2  定时任务-1或者-2都可以  他是从1到renwu之间 取一个随机数 
    var zong = 10   //总任务  就是你有几个号你就写几  就行 默认10 也可以 如果多10行  就在10行以内 右键 插入行 

    var tokenColumn = "A"; // 设置列号变量为 "A"
    // var qiandaoColumn = "B"; // 设置列号变量为 "B"
    // var serverchanColumn = "C"; // 设置列号变量为 "C"
    // var pushplusColumn = "D"; // 设置列号变量为 "D"
    // var pushColumn = "E"; // 设置列号变量为 "E"
    // var logindateColumn = "F"; // 设置列号变量为 "F"
    // var signinresult = "G"//签到的结果  设置列号变量为 "G"
    var qiandaoColumn = "B"; // 设置列号变量为 "B",签到且领取奖励
    var backupsColumn = "D";  // 备份奖励
    var logindateColumn = "E"; // 设置列号变量为 "F",登录时间
    var signinresult = "F"//签到的结果  设置列号变量为 "G"

    messageSuccess = ""
    messageFail = ""
    content = ""

// for (let row = 1; row <= zong; row++) { // 循环遍历从第 1 行到第 10 行的数据


    var refresh_token = Application.Range(tokenColumn + row).Text; // 获取指定单元格的值
    var qiandao = Application.Range(qiandaoColumn + row).Text; // 获取指定单元格的值
    // var servertoken = Application.Range(serverchanColumn + row).Text; // 获取指定单元格的值
    // var pushtoken = Application.Range(pushplusColumn + row).Text; // 获取指定单元格的值
    // var push = Application.Range(pushColumn + row).Text; // 获取指定单元格的值
    var ldate = Application.Range(logindateColumn + row).Text; // 获取指定单元格的值
    var signin = Application.Range(signinresult + row).Text; // 获取签到结果
    var backups = Application.Range(backupsColumn + row).Text;  //是否领取备份奖励  是写true  不领取写false
    if(backups == "是")
    {
      backup = true
    }else
    {
      backup = false
    }
    if (refresh_token != "") { // 如果刷新令牌不为空
        if (qiandao == "是") {//签到&领奖开关
            if (signin !== date + '已签到') {
                // var randomInt = Math.floor(Math.random() * renwu) + 1
                randomInt = 1 ; // 强制签到
                //randomInt  中的3 可以修改  写3就是 从1到3取一个随机整数 比如2  当他 = 1 的时候 签到才会执行
                if (randomInt === 1) {//等于1 就开始签到  不等于1 都不签到
 
                    //获取Bearer-token
                    var mtid = parseInt(Math.floor(Math.random() * 60000 * dengdai)) + 6000
                    var loginresult = "登录延迟" + parseFloat((mtid / 120000).toFixed(2)) + "分，即" + + parseFloat((mtid / 2000).toFixed(2)) + "秒"
                    console.log("🍳 登录延迟" + parseFloat((mtid / 120000).toFixed(2)) + "分，即" + + parseFloat((mtid / 2000).toFixed(2)) + "秒")
                    // Time.sleep(mtid / 2)  // 进行延迟
                    let data = HTTP.post("https://auth.aliyundrive.com/v2/account/token",
                        JSON.stringify({
                            "grant_type": "refresh_token",
                            "refresh_token": refresh_token
                        })
                    )
                    data = data.json()
                    var access_token = data['access_token']
                    var phone = data["user_name"]
                    if (access_token == undefined) { // 如果报错
                        console.log("🍳 单元格【" + tokenColumn + row + "】token执行出错,请检查token或者API接口");
                        // continue; // 跳过当前行的后续操作()
                    }
                    var mtid = parseInt(Math.floor(Math.random() * 60000 * dengdai / 2)) + 6000
                    var signresult = "签到延迟" + parseFloat((mtid / 60000).toFixed(2)) + "分，即" + + parseFloat((mtid / 1000).toFixed(2)) + "秒"
                    console.log("🍳 签到延迟" + parseFloat((mtid / 60000).toFixed(2)) + "分，即" + + parseFloat((mtid / 1000).toFixed(2)) + "秒")
                    // Time.sleep(mtid)  // // 进行延迟
                    try {
                        // 签到
                        var access_token2 = 'Bearer ' + access_token; // 构建包含访问令牌的请求头
                        let data2 = HTTP.post("https://member.aliyundrive.com/v1/activity/sign_in_list",
                            JSON.stringify({ "_rx-s": "mobile" }),
                            { headers: { "Authorization": access_token2 } }
                        );
                        data2 = data2.json(); // 将响应数据解析为 JSON 格式
                        var signin_count = data2['result']['signInCount']; // 获取签到次数
                        var result1 = "账号：" + phone + " - 签到成功";
                        var result2 = "本月累计签到 " + signin_count + " 天";
                        // console.log(result1)
                        content = result1 + " " + result2
                        messageSuccess += content
                        console.log(content)
                    } catch (error) {
                        console.log("🍳 单元格【" + tokenColumn + row + "】签到出错,请检查API接口")
                        content = "签到出错,请检查API接口 "
                        messageFail += content
                        // console.log(content)
                        // continue; // 跳过当前行的后续操作()
                    }
                    Time.sleep(3000)
                    if(0){
                    try {
                        // 领取奖励
                        let data3 = HTTP.post(
                            "https://member.aliyundrive.com/v1/activity/sign_in_reward?_rx-s=mobile",
                            JSON.stringify({ "signInDay": signin_count }),
                            { headers: { "Authorization": access_token2 } }
                        );
                        data3 = data3.json(); // 将响应数据解析为 JSON 格式
                        console.log(data3)
                        var result3 = data3["result"]["name"]; // 获取奖励名称
                        var result4 = data3["result"]["notice"]; // 获取奖励描述
                        Application.Range(signinresult + row).Value = date + '已签到'
                        //把签到结果 写入文档内
                        // console.log(result4)
                        content = " " + result4
                        messageSuccess += content
                        console.log(content)
                    } catch (error) {
                        // console.log("🍳 单元格【" + tokenColumn + row + "】领奖出错，请手动确认");
                        // continue; // 跳过当前行的后续操作()
                        console.log("🍳 单元格【" + tokenColumn + row + "】领奖出错，请手动确认")
                        content = "领奖出错，请手动确认 "
                        messageFail += content
                        // console.log(content)
                    }
                    }
                    if (backups === true) {
                        // try {
                        // 备份的奖励  
                        var access_token2 = 'Bearer ' + access_token; // 构建包含访问令牌的请求头
                        let data5 = HTTP.post("https://member.aliyundrive.com/v2/activity/sign_in_task_reward",
                            JSON.stringify({ "signInDay": signin_count }),
                            { headers: { "Authorization": access_token2 } }
                        );
                        data5 = data5.json(); // 将响应数据解析为 JSON 格式
                        // console.log('备份奖励', data5)
                        var success = data5['success']
                        if (success == true) {
                            var result5 = data5["result"]["notice"]; // 获取奖励名称
                            content = " " + result5
                            messageSuccess += content
                            console.log(content)
                        } else {
                            var result5 = data5["message"] // 失败原因
                            content = " " + result5
                            messageFail += content
                            console.log(content)
                        }
 
                        // console.log(result5)
                        

                        // } catch (error) {
                        //     console.log("🍳 单元格【" + tokenColumn + row + "】领取备份出错,请检查API接口");
                        //     continue; // 跳过当前行的后续操作()
                        // }
                    } else {
                        // console.log('不领取备份奖励')
                        content = " " + '不领取备份奖励'
                        messageSuccess += content
                        console.log(content)
                    }
 
 
                    var loginnotice = "" //20天后自动刷新refresh_token
                    var ldate = Application.Range(logindateColumn + row).Text; // 获取指定单元格的值
                    if (ldate !== '') {
                        Application.Range(logindateColumn + row).NumberFormat = 'yyyy-mm-dd;@'
                        var formatlogindate = Application.Range(logindateColumn + row).Text
                        console.log(formatlogindate)
                        function formatDateTime(date) {
                            const year = date.getFullYear();
                            const month = date.getMonth() + 1;
                            const day = date.getDate();
                            return `${year}-${pad(month)}-${pad(day)} `;
                        }
                        function pad(num) {
                            return num.toString().padStart(2, '0');
                        }
                        const currentDate = new Date();
                        var formacurrentdate = formatDateTime(currentDate);
                        console.log(formacurrentdate)
                        function getDate(strDate) {
                            if (strDate == null || strDate === undefined) return null;
                            var date = new Date();
                            try {
                                if (strDate == undefined) {
                                    date = null;
                                } else if (typeof strDate == 'string') {
                                    strDate = strDate.replace(/:/g, '-');
                                    strDate = strDate.replace(/ /g, '-');
                                    var dtArr = strDate.split("-");
                                    if (dtArr.length >= 3 && dtArr.length < 6) {
                                        date = new Date(dtArr[0], dtArr[1], dtArr[2]);
                                    } else if (date.length > 8) {
                                        date = new Date(Date.UTC(dtArr[0], dtArr[1] - 1, dtArr[2], dtArr[3] - 8, dtArr[4], dtArr[5]));
                                    }
                                } else {
                                    date = null;
                                }
                                return date;
                            } catch (e) {
                                alert('格式化日期出现异常：' + e.message);
                            }
                        }
                        var timeslong = getDate(formacurrentdate).getTime() - getDate(formatlogindate).getTime();
                        console.log(timeslong)
                        if (timeslong > 1728000000) {//时间差单位毫秒
                            var loginnotice = "登录已超20天自动刷新refresh_token";
 
                            let my_token = HTTP.post("https://auth.aliyundrive.com/v2/account/token",
                                JSON.stringify({
                                    "grant_type": "refresh_token",
                                    "refresh_token": refresh_token
                                }));
                            my_token = my_token.json()["refresh_token"]
                            if (my_token) {
                                console.log("🍳 当前账号refresh_token刷新为", my_token);
                                Application.Range(tokenColumn + row).Value = my_token;
                                console.log("🍳 当前账号登录日期刷新为", formacurrentdate);
                                Application.Range(logindateColumn + row).Value = formacurrentdate
                            }
                        }
                    }
 
 
                } else {//randomInt 不等于1  不签到
                    console.log('取得随机值不是1,不签到')
                }
 
            } else {//签到了
                // console.log(date + '已签到')
                content = " " + date + '已签到'
                messageSuccess += content
                console.log(content)
            }
        }//不需要签到
    }//令牌为空
    //无有效token

    msg = [messageSuccess, messageFail]
    return msg
}
// 引用结束



// 具体的执行函数
function execHandle(cookie, pos) {
  let messageSuccess = "";
  let messageFail = "";
  let messageName = ""; // 存放推送位置标识，如昵称或单元格（昵称为空时）

  // // 推送昵称还是单元格
  // if (messageNickname == 1) {
  //   messageName = Application.Range("C" + pos).Text;
  // } else {
  //   messageName = "单元格A" + pos + "";
  // }

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
    
    // console.log(msg)
    msg = doTask(pos)
    messageSuccess += "🎉 " + msg[0]
    messageFail += msg[1]

    
    
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
  // if (messageOnlyError == 1) {
  //   message += messageFail;
  // } else {
  //   message += messageFail + " " + messageSuccess;
  // }

  if (messageOnlyError == 1) {
    messageArray[posLabel] = messageFail;
  } else {
    messageArray[posLabel] = messageFail + " " + messageSuccess;
  }

  if(messageArray[posLabel] != "")
  {
    console.log(messageArray[posLabel]);
  }
}

