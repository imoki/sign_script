// 叮咚买菜-叮咚鱼塘自动签到
// 20240613

let sheetNameSubConfig = "ddmc"; // 分配置表名称
let sheetNameSubConfig2 = "ddmc_ddyt";
let pushHeader = "【叮咚买菜-叮咚鱼塘】";
let sheetNameConfig = "CONFIG"; // 总配置表
let sheetNamePush = "PUSH"; // 推送表名称
let sheetNameEmail = "EMAIL"; // 邮箱表
let flagSubConfig = 0; // 激活分配置工作表标志
let flagConfig = 0; // 激活主配置工作表标志
let flagPush = 0; // 激活推送工作表标志
let line = 21; // 指定读取从第2行到第line行的内容
var message = ""; // 待发送的消息
var messageOnlyError = 0; // 0为只推送失败消息，1则为推送成功消息。
var messageNickname = 0; // 1为用昵称替代单元格，0为不替代
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
    if (name == sheetNameSubConfig2) {
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

  push(message); // 推送消息
}

// 总推送
function push(message) {
  if (message != "") {
    message = pushHeader + message; // 加上推送头
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
    url = "http://www.pushplus.plus/send?token=" + key + "&content=" + message;
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

// 具体的执行函数
function execHandle(cookie, pos) {
  let messageSuccess = "";
  let messageFail = "";
  let messageName = "";
  if (messageNickname == 1) {
    messageName = Application.Range("C" + pos).Text;
  } else {
    messageName = "单元格A" + pos + "";
  }
  try {
    let seedId = Application.Range("F" + pos).Text;
    let propsId = Application.Range("G" + pos).Text;

    // 获取任务taskCode
    let taskCode = []
    // 领取任务奖励
    let userTaskLogId = []

    let url = [
      'https://sunquan.api.ddxq.mobi/api/v2/user/signin/',//  签到积分
      'https://farm.api.ddxq.mobi/api/v2/task/achieve?api_version=9.1.0&app_client_id=1&station_id=&stationId=&native_version=&app_version=10.15.0&OSVersion=15&CityId=0201&uid=&latitude=40.123389&longitude=116.345477&lat=40.123389&lng=116.345477&device_token=&gameId=1&taskCode=DAILY_SIGN',  // 签到
      'https://farm.api.ddxq.mobi/api/v2/task/achieve?api_version=9.1.0&app_client_id=1&station_id=&stationId=&native_version=&app_version=10.1.2&OSVersion=15&CityId=0201&uid=&latitude=40.123389&longitude=116.345477&lat=40.123389&lng=116.345477&device_token=&gameId=1&taskCode=CONTINUOUS_SIGN',  // 签到2
      'https://farm.api.ddxq.mobi/api/v2/props/feed?api_version=9.1.0&app_client_id=1&station_id=&stationId=&native_version&app_version=10.0.1&OSVersion=15&CityId=0201&uid=&latitude=40.123389&longitude=116.345477&lat=40.123389&lng=116.345477&device_token=&gameId=1&propsId=' + propsId + '&seedId=' + seedId + '&cityCode=0201&feedPro=0&triggerMultiFeed=1',// 喂饲料
      'https://farm.api.ddxq.mobi/api/v2/task/list?latitude=40.123389&longitude=116.345477&env=PE&station_id=&city_number=0201&api_version=9.44.0&app_client_id=3&native_version=10.15.0&h5_source=&page_type=2&gameId=1',  // 获取任务taskCode
      'https://farm.api.ddxq.mobi/api/v2/task/achieve?api_version=9.1.0&app_client_id=1&station_id=&stationId=&native_version=&app_version=10.15.0&OSVersion=15&CityId=0201&uid=&latitude=40.123389&longitude=116.345477&lat=40.123389&lng=116.345477&device_token=&gameId=1&taskCode=',  // 完成任务
      'https://farm.api.ddxq.mobi/api/v2/task/reward?api_version=9.1.0&app_client_id=1&station_id=&stationId=&native_version=&app_version=10.15.1&OSVersion=15&CityId=0201&uid=&latitude=40.123389&longitude=116.345477&lat=40.123389&lng=116.345477&device_token=&userTaskLogId=',// 领取任务奖励
    ]

    headers = {
      'Host': 'farm.api.ddxq.mobi',
      'Origin': 'https://game.m.ddxq.mobi',
      'Cookie': cookie,
      'Accept': '*/*',
      // 'User-Agent': 'Mozilla/5.0 (iPhone; CPU iPhone OS 15_7 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Mobile/15E148 xzone/10.0.1 station_id/' + station_id + ' device_id/' + device_id,
      'Accept-Language': 'zh-CN,zh-Hans;q=0.9',
      'Referer': 'https://game.m.ddxq.mobi/',
    };

    headerintegral =
    { // 积分
        'Host': 'sunquan.api.ddxq.mobi',
        'Cookie': cookie,
        'Referer': 'https://activity.m.ddxq.mobi/',
        'ddmc-city-number': '0201',
        'ddmc-api-version': '9.7.3',
        'Origin': 'https://activity.m.ddxq.mobi',
        'ddmc-build-version': '10.15.0',
        'ddmc-longitude': 114.345477,
        'ddmc-latitude': 40.123389,
        'ddmc-app-client-id': 3,
        'Accept-Language': 'zh-CN,zh-Hans;q=0.9',
        'ddmc-channel': ' ',
        'Accept': '*/*',
        'Content-Type': 'application/x-www-form-urlencoded',
        'ddmc-station-id': '',
        'ddmc-ip': '',
    },

    dataintegral = {
      'api_version':'9.7.3',
      'app_client_id':3,
      'app_version':'2.14.5',
      'app_client_name':'activity',
      'station_id':'',
      'native_version':'10.15.0',
      'city_number':'0201',
      'device_token':'',
      'device_id':'',
      'latitude':'40.123389',
      'longitude':'116.345477',
    }
    // 积分签到
    resp = HTTP.fetch(url[0], {
      method: "post",
      headers: headerintegral,
      data : dataintegral
    });

    if (resp.status == 200) {
      resp = resp.json();
      console.log(resp);
      code = resp["code"];
      msg = resp["msg"];
      if(code == 0){
        // content = "帐号：" + messageName + "积分签到成功 "
        content = "积分签到成功 "
        messageSuccess += content
        console.log(content);
      }else{
        // {"msg":"出了点问题哦，请稍后再试吧","code":119000001,"timestamp":"2023-08-10 21:06:53","success":false,"exec_time":{}}
        // content += "帐号：" + messageName + msg + " ";
        content += msg + " ";
        messageFail += content;
        console.log(content);
      }
    } else {
      console.log(resp.text());
      // content = "帐号：" + messageName + "积分签到失败 "
      content = "积分签到失败 "
      messageFail += content;
      console.log(content);
    }

    // 签到领饲料
    let flagSign = 0; // 标识是否签到领取饲料
    let tempmessageFail = "";  // 记录临时失败的消息
    resp = HTTP.fetch(url[1], {
      method: "get",
      headers: headers,
    });

    if (resp.status == 200) {
      resp = resp.json();
      console.log(resp);
      code = resp["code"];
      msg = resp["msg"];
      if(code == 0){
        // messageSuccess += "帐号：" + messageName + "鱼塘签到成功 "
        flagSign = 1;
        console.log("帐号：" + messageName + "鱼塘签到成功 ");
      }else{
        // {"msg":"今日已完成任务，明日再来吧！","code":601,"timestamp":"2023-08-10 21:23:49","success":false,"exec_time":{}}
        // {"msg":"出了点问题哦，请稍后再试吧","code":119000001,"timestamp":"2023-08-10 21:06:53","success":false,"exec_time":{}}
        // messageFail += "帐号：" + messageName + msg + " ";
        console.log("帐号：" + messageName + msg + " ");
      }
    } else {
      console.log(resp.text());
      // messageFail += "帐号：" + messageName + "签到失败 ";
      console.log("帐号：" + messageName + "签到失败 ");
    }

    resp = HTTP.fetch(url[2], {
      method: "get",
      headers: headers,
    });

    if (resp.status == 200) {
      resp = resp.json();
      console.log(resp);
      code = resp["code"];
      msg = resp["msg"];
      if(code == 0 ){
        flagSign = 1;
        console.log("帐号：" + messageName + "鱼塘签到成功 ");
      }else{
        if(code == 601){
          // 此不为错误消息
          // {"msg":"今日已完成任务，明日再来吧！","code":601,"timestamp":"2024-06-13 20:30:28","success":false}
          flagSign = 1;
          console.log("帐号：" + messageName + msg + " ");
        }else{
          // content = "帐号：" + messageName + msg + " ";
          content = msg + " ";
          tempmessageFail = content;
          console.log(content);

        }
        
      }
    } else {
      console.log(resp.text());
      // content = "帐号：" + messageName + "签到失败 ";
      content = "签到失败 ";
      tempmessageFail = content;
      console.log(content);
    }

    if(flagSign == 1){
      // content = "帐号：" + messageName + "鱼塘签到成功 "
      content = "鱼塘签到成功 "
      messageSuccess += content;
    }else{
      messageFail += tempmessageFail;
    }

    // 获取任务列表
    resp = HTTP.fetch(url[4], {
      method: "get",
      headers: headers,
    });

    if (resp.status == 200) {
      resp = resp.json();
      // console.log(resp);
      code = resp["code"];
      if(code == 0){
        console.log("正在获取taskCode ");
        userTasks = resp["data"]["userTasks"];
        for (let j = 0; j < userTasks.length; j++) {
          taskCode[j] = userTasks[j]["taskCode"]
        }
        console.log(taskCode)
      }else{
        console.log("获取taskCode失败 ");
      }
    } else {
      console.log(resp.text());
      console.log("获取taskCode失败 ");
    }

    // taskCode = ["ANY_ORDER","BROWSE_GOODS","BUY_GOODS","CONTINUOUS_SIGN","DAILY_SIGN","FIRST_ORDER","HARD_BOX","INVITATION","LOTTERY","LUCK_DRAW","MULTI_ORDER","STEAL_FEED"]
    // 完成任务
    if(taskCode.length > 0){
      console.log("尝试完成任务...")
      for (let j = 0; j < taskCode.length; j++) {
          urlTask = url[5] + taskCode[j]
          // console.log(urlTask)
          try{
            resp = HTTP.fetch(urlTask, {
              method: "get",
              headers: headers,
            });
            // console.log(resp.text())
            sleep(2000)
          }catch{
            console.log("忽略任务：" + taskCode[j])
          }
      }
    }

    // 获取奖励id
    resp = HTTP.fetch(url[4], {
      method: "get",
      headers: headers,
    });

    if (resp.status == 200) {
      resp = resp.json();
      // console.log(resp);
      code = resp["code"];
      if(code == 0){
        console.log("正在获取userTaskLogId ");
        userTasks = resp["data"]["userTasks"];
        let temp
        let num = 0
        for (let j = 0; j < userTasks.length; j++) {
          temp = userTasks[j]["userTaskLogId"]
          // console.log(typeof(temp))
          // console.log(temp.length) // 长度为18才是id
          if(typeof(temp) != "object"){  // || temp != "{}" || temp != "null" || temp != "" || temp != null
            userTaskLogId[num] = temp
            num += 1
          }
        }
        console.log(userTaskLogId)
      }else{
        console.log("获取userTaskLogId失败 ");
      }
    } else {
      console.log(resp.text());
      console.log("获取userTaskLogId失败 ");
    }

    // 领取任务奖励
    if(userTaskLogId.length > 0){
      console.log("尝试领取任务奖励...")
      for (let j = 0; j < userTaskLogId.length; j++) {
          urlTask = url[6] + userTaskLogId[j]
          // console.log(urlTask)
          try{
            resp = HTTP.fetch(urlTask, {
              method: "get",
              headers: headers,
            });
            // console.log(resp.text())
            sleep(2000)
          }catch{
            console.log("忽略任务：" + userTaskLogId[j])
          }
      }
    }else{
      console.log("没有可领取的奖励")
    }

    // 喂饲料
    let amount = 10; // 记录剩余数目
    let amoutCount = 0; // 已喂饲料次数
    let flagAmount = 0;  // 标志，1为饲料
    let countSeedId = 0; // 计算是不是每次浇花的剩余水量都一样，如果三次都一样，则认为seedid过期
    let lastamount = 0; // 记录上一次剩余水量
    while(amount >= 10){
      resp = HTTP.fetch(url[3], {
        method: "get",
        headers: headers,
      });

      if (resp.status == 200) {
        resp = resp.json();
        // console.log(resp);
        code = resp["code"];
        msg = resp["msg"];
        if(code == 0){
          amount = resp["data"]["props"]["amount"];

          // 用于判断seedId是否过期，也即浇水是否失败
          if(lastamount == amount){ // 和上次剩余水量一样，可能没浇水成功
            countSeedId += 1; // 记录相同次数
          }else{
            countSeedId = 0;  // 水量不同，浇水成功，置零
          }
          lastamount = amount; // 记录水量，以便下一次循环使用
          if(countSeedId >=3){  // 浇了三次剩余水量都相同，则认为浇水失败，不再浇水，并提醒用户更换新的seedId值
            msg = "[❗❗❗提醒]seedId值可能过期，请抓包获取最新的值"
            messageFail += "[❗❗❗提醒]seedId值可能过期，请抓包获取最新的值"
            console.log("提前退出浇水，错误消息为：" + msg)
            amoutCount -= 3;  // 减去浇水失败的次数
            break;  
          }

          flagAmount = 1;
          amoutCount += 1;
          console.log("喂饲料中... ,剩余饲料：" + amount)
        }else{
          console.log(resp);
          console.log("提前退出喂饲料，错误消息为：" + msg)
          amount = 0; // 直接置水为0 退出投喂
        }
      } else {
        console.log(resp.text());
        console.log("提前退出喂饲料")
        amount = 0; // 直接置水为0 退出投喂
      }

      sleep(3000)
    }

    if(flagAmount ==  1){
      messageSuccess += "成功喂饲料" +  amoutCount + "次 "
      console.log("成功喂饲料" +  amoutCount + "次 ");
    }else{
      // messageFail += "喂饲料日志：" + msg + " ";   // 此错误消息无需推送
      console.log( "喂饲料日志：" + msg + " ");
    }


  } catch {
    messageFail += messageName + "失败";
  }

  sleep(2000);
  if (messageOnlyError == 1) {
    message += messageFail;
  } else {
    message += messageName + " "
    message += messageFail + " " + messageSuccess;
  }
  console.log(message);
}