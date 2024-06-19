// 中兴商城自动签到、做任务
// 20240602

let sheetNameSubConfig = "ztemall"; // 分配置表名称
let pushHeader = "【中兴商城】";
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

  message = messageMerge()// 将消息数组融合为一条总消息
  push(message); // 推送消息
}

// 将消息数组融合为一条总消息
function messageMerge(){
  for(i=0; i<messageArray.length; i++){
    if(messageArray[i] != "" && messageArray[i] != null)
    {
      message += messageHeader[i] + messageArray[i] + " "; // 加上推送头
    }
  }
  if(message != "")
  {
    console.log(message)  // 打印总消息
  }
  return message
}

// 总推送
function push(message) {
  if (message != "") {
    message = messagePushHeader + message // 消息头最前方默认存放：【xxxx】
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
  messageHeader[posLabel] = messageName
  // try {
    // var url1 = "https://api-bbs.ztedevices.com/points/home/pointsRegister"; // 社区签到
    var url2 = "https://www.ztemall.com/index.php/topapi" // 商城签到
    // ztemallcookie = Application.Range("D" + pos).Text;  // 商城的cookie
    
    // // 社区签到
    // // {"status":200,"msg":"success","data":{"list":[{"day":"4.30","energy":0,"tab":1},{"day":"5.1","energy":0,"tab":1},{"day":"5.2","energy":0,"tab":1},{"day":"5.3","energy":1,"tab":2},{"day":"5.4","energy":2,"tab":3},{"day":"5.5","energy":3,"tab":3},{"day":"5.6","energy":4,"tab":3}],"msg":"签到成功！经验 +10 流星 +1"","continueDays":1}}
    // // {"status":200,"msg":"success","data":{"list":[{"day":"4.30","energy":0,"tab":1},{"day":"5.1","energy":0,"tab":1},{"day":"5.2","energy":0,"tab":1},{"day":"5.3","energy":1,"tab":2},{"day":"5.4","energy":2,"tab":3},{"day":"5.5","energy":3,"tab":3},{"day":"5.6","energy":4,"tab":3}],"msg":"签到成功","continueDays":1}}
    // // {"status":400,"msg":"请求数据异常！","data":[]}
    // data = {
    //     "v" : ""
    // }

    // headers={
    //   "Accesstoken": cookie, 
    //   "Content-Type": "application/x-www-form-urlencoded",
    //   // "Host":"api-bbs.ztedevices.com",
    //   // "Cookie":""
    // }

    // resp = HTTP.post(
    //   url1,
    //   JSON.stringify(data),
    //   { headers: headers }
    // );

    // if (resp.status == 200) {
    //   resp = resp.json();
    //   console.log(resp)
    //   status = resp["status"]
      
    //   if(status == 200)
    //   {
    //     msg = resp["data"]["msg"]
    //     continueDays = resp["data"]["continueDays"]

    //     content = msg + " 已签" + continueDays + "天 "
    //     messageSuccess += content;
    //     console.log(content)
    //   }else
    //   {
    //     content = "签到失败 "
    //     messageFail += content;
    //     console.log(content);
    //   }
    // } else {
    //   content = "签到失败 "
    //   messageFail += content;
    //   console.log(content);
    // }
    
    try{
      cookie_json = cookie_to_json(cookie)
      accessToken = cookie_json["auth_token_pc"]
      if(accessToken == undefined)
      {
        accessToken = cookie
        console.log("cookie中无auth_token_pc，将cookie作为accessToken")
      }
    }catch{
      accessToken = cookie
    }

    console.log("已读取到accessToken:" + accessToken)

    // 商城签到
    params = "?method=member.checkIn.add&format=json&v=v1&sign=&accessToken=" + accessToken
    headers={
      "Host": "www.ztemall.com",
      // "Cookie": ztemallcookie, 
    }
    resp = HTTP.fetch(url2 + params, {
      method: "get",
      headers: headers,
      // data: data
    });

    // {"errorcode":0,"msg":"","data":{"checkin_days":1,"currentCheckInPoint":"10","point":1010,"status":"success"}}
    // {"errorcode":"10000","msg":"会员签到记录保存失败","data":{}}
    // {"errorcode":20001,"msg":"invalid token","data":{}}
    if (resp.status == 200) {
      resp = resp.json();
      console.log(resp)
      errorcode = resp["errorcode"]
      
      if(errorcode == 0)
      {
        msg = resp["msg"]
        checkin_days = resp["data"]["checkin_days"] // 签到天数
        currentCheckInPoint = resp["data"]["currentCheckInPoint"] // 获得积分
        point = resp["data"]["point"] // 总积分

        content = msg + " 连签" + checkin_days + "天,获得" + currentCheckInPoint +"积分 " //，当前共有" + point + "积分 "
        messageSuccess += content;
        console.log(content)
      }else if(errorcode == 10000){
        content =" 可能已经签过了 "
        messageSuccess += content;
        console.log(content)
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

    // 商城做任务
    var url3 = "https://www.ztemall.com/index.php/topapi?platform=h5&method=task.list&format=json&v=v1&sign=&accessToken=" // 商城任务列表
    var url4 = "https://www.ztemall.com/index.php/topapi" // 商城任务开始、任务完成、任务领取(get)

    // 取任务id
    console.log("获取任务id列表")
    resp = HTTP.fetch(url3 + params, {
      method: "get",
      headers: headers,
      // data: data
    });
    task_id_list = []  // 存放任务列表
    resp = resp.json()
    console.log(resp)
    tasks = resp["data"]["tasks"]
    // console.log(tasks.length)
    for (key = 0; key < tasks.length; key ++) {

      // 类型 去参与、立即领取、已完成、去拼团、去购买。只有去参与才是浏览10s,可以做任务
      desc = tasks[key]["btn"]["desc"]
      if(desc == "去参与" || desc == "立即领取")
      {
        dict = {
          "title" : tasks[key]["title"],
          "task_id" : tasks[key]["task_id"],
          "page_ids" : tasks[key]["task_data"]["page_ids"],
        }
        // console.log(dict)
        task_id_list.push(dict)
      }
      // console.log(tasks[key]["title"])
      // status = tasks[i]["btn"]["status"]  // 任务状态 finish、todo
    }
    console.log(task_id_list)
    
    data = {
      "page_id":0,
      "task_id":"",
      "method":"task.start",  // task.start task.finish
      "format":"json",
      "v":"v1",
      "sign":"",
      "accessToken":accessToken,
    }

    headers={
      "Host": "www.ztemall.com",
      "Content-Type": "application/x-www-form-urlencoded",
    }

    // headers["Content-Type"] = "application/x-www-form-urlencoded"

    for(i=0; i<task_id_list.length; i++){
      // 任务开始
      // console.log(task_id_list[i])
      title = task_id_list[i]["title"]
      task_id = task_id_list[i]["task_id"]
      data["task_id"] = task_id
      try{
        page_id = task_id_list[i]["page_ids"][0]
        // console.log(page_id)
        if(page_id != undefined && page_id != "" && page_id != "undefined"){
          data["page_id"] = page_id
        }else
        {
          data["page_id"] = 0
        }
        // console.log("page_id:" + data["page_id"])
      }catch{
        data["page_id"] = 0
      }

      data["method"] = "task.start"
      console.log("开始做任务：" + title) // + " 任务id:" + task_id)
      resp = HTTP.post(
        url4,
        // JSON.stringify(data),
        data,
        { headers: headers }
      );
      
      // {"errorcode":10001,"msg":"系统参数：版本号必填","data":{}}
      // {"from":"lua","msg":"无效请求","errorcode":20001,"data":{}}
      // {"errorcode":"10000","msg":"找不到API:task.start","data":{}}
      // {"from":"lua","msg":"","errorcode":0,"data":{"res_sec":10}}
      
      // {"from":"lua","msg":"请从任务中心开始任务","errorcode":10000,"data":{}}
      // {"from":"lua","msg":"该类任务不支持此操作","errorcode":10000,"data":{}}
      resp = resp.json()
      console.log(resp)
      errorcode = resp["errorcode"]
      if(errorcode != 0)  // 这个任务没做成功，直接进入下一个任务
      {
        content = title + "无法开始，直接进入下一个任务 "
        messageFail += content
        console.log(content)
        sleep(2000);
        continue
      }

      sleep(11000);
      // 任务完成
      data["method"] = "task.finish"
      resp = HTTP.post(
        url4,
        // JSON.stringify(data),
        data,
        { headers: headers }
      );
      resp = resp.json()
      console.log(resp)
      errorcode = resp["errorcode"]
      // {"from":"lua","msg":"","errorcode":0,"data":{"status":"succ"}}
      if(errorcode != 0)  // 这个任务没做成功，直接进入下一个任务
      {
        content = title + "无法完成，直接进入下一个任务 "
        messageFail += content
        console.log(content)
        sleep(2000);
        continue
      }

      sleep(2000);
      // 领取奖励
      params = "?task_id=" + task_id + "&method=task.check&format=json&v=v1&sign=&accessToken=" + accessToken // 商城任务领取
      resp = HTTP.fetch(url4 + params, {
        method: "get",
        headers: headers,
      });
      resp = resp.json()
      console.log(resp)
      // {"errorcode":20001,"msg":"invalid token","data":{}}
      // {"errorcode":0,"msg":"","data":{"status":"succ"}}
      if(errorcode != 0)  // 这个任务没做成功，直接进入下一个任务
      {
        content = title + "无法领取奖励，请手动领取奖励 "
        messageFail += content
        console.log(content)
        sleep(2000);
        continue
      }else{
        content = title + "已完成 "
        messageFail += content
        console.log(content)
      }

      sleep(2000);
    }

    // 查询当前积分
    url5 = "https://www.ztemall.com/index.php/topapi?page_no=1&page_size=1&method=member.point.detail&format=json&v=v1&sign=&accessToken=" + accessToken
    // headers={
    //   "Host": "www.ztemall.com",
    // }

    resp = HTTP.fetch(url5, {
      method: "get",
      headers: headers,
    });
    resp = resp.json()
    console.log(resp)
    errorcode = resp["errorcode"]
    if(errorcode == 0)
    {
      point_count = resp["data"]["point_total"]["point_count"]  // 积分总量
      content = "当前积分:" + point_count + " "
      messageSuccess += content
      console.log(content)
      sleep(2000);
    }else{
      content = "无法查询到当前积分 "
      // messageFail += content
      console.log(content)
    }

  // } catch {
  //   messageFail += messageName + "失败";
  // }

  sleep(2000);
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
