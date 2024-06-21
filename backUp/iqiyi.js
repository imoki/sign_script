// 爱奇艺自动签到、白金抽奖、做任务获取成长值
// 20240512

let sheetNameSubConfig = "iqiyi"; // 分配置表名称
let pushHeader = "【爱奇艺】";
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

function give_times(P00001){
    // {"code":"A00000","msg":"成功","data":{}}
    url = "https://pcell.iqiyi.com/lotto/giveTimes"
    times_code_list = ["browseWeb", "browseWeb", "bookingMovie"]
    for(times_code in times_code_list){
      times_code = times_code_list[times_code]
      // console.log(times_code)
      params = "?actCode=bcf9d354bc9f677c&timesCode=" + times_code + "&P00001=" + P00001 
      resp = HTTP.fetch(url + params , {
        method: "get",
      });
      console.log(resp.json())
    }
}

// // 摇一摇抽奖
// function lottery(P00001, award_list=[]){
//   url = "https://act.vip.iqiyi.com/shake-api/lottery"
//   params = {
//       "P00001": p00001,
//       "lotteryType": "0",
//       "actCode": "0k9GkUcjqqj4tne8",
//   }
//   params = {
//       "P00001": p00001,
//       "deviceID": str(uuid4()),
//       "version": "15.3.0",
//       "platform": str(uuid4())[:16],
//       "lotteryType": "0",
//       "actCode": "0k9GkUcjqqj4tne8",
//       "extendParams": json.dumps(
//           {
//               "appIds": "iqiyi_pt_vip_iphone_video_autorenew_12m_348yuan_v2",
//               "supportSk2Identity": True,
//               "testMode": "0",
//               "iosSystemVersion": "17.4",
//               "bundleId": "com.qiyi.iphone",
//           }
//       ),
//   }
  
//   res = requests.get(url, params=params).json()
//   msgs = []
//   if res.get("code") == "A00000":
//       award_info = res.get("data", {}).get("title")
//       award_list.append(award_info)
//       time.sleep(3)
//       return self.lottery(p00001=p00001, award_list=award_list)
//   elif res.get("msg") == "抽奖次数用完":
//       if award_list:
//           msgs = [{"name": "每天摇一摇", "value": "、".join(award_list)}]
//       else:
//           msgs = [{"name": "每天摇一摇", "value": res.get("msg")}]
//   else:
//       msgs = [{"name": "每天摇一摇", "value": res.get("msg")}]
//   return msgs
// }

// function draw(draw_type, p00001, p00003):
//         """
//         查询抽奖次数(必),抽奖
//         :param draw_type: 类型。0 查询次数；1 抽奖
//         :param p00001: 关键参数
//         :param p00003: 关键参数
//         :return: {status, msg, chance}
//         """
//         url = "https://iface2.iqiyi.com/aggregate/3.0/lottery_activity"
//         params = {
//             "lottery_chance": 1,
//             "app_k": "b398b8ccbaeacca840073a7ee9b7e7e6",
//             "app_v": "11.6.5",
//             "platform_id": 10,
//             "dev_os": "8.0.0",
//             "dev_ua": "FRD-AL10",
//             "net_sts": 1,
//             "qyid": "2655b332a116d2247fac3dd66a5285011102",
//             "psp_uid": p00003,
//             "psp_cki": p00001,
//             "psp_status": 3,
//             "secure_v": 1,
//             "secure_p": "GPhone",
//             "req_sn": round(time.time() * 1000),
//         }
//         if draw_type == 1:
//             del params["lottery_chance"]
//         res = requests.get(url=url, params=params).json()
//         if not res.get("code"):
//             chance = int(res.get("daysurpluschance"))
//             msg = res.get("awardName")
//             return {"status": True, "msg": msg, "chance": chance}
//         else:
//             try:
//                 msg = res.get("kv", {}).get("msg")
//             except Exception as e:
//                 print(e)
//                 msg = res["errorReason"]
//         return {"status": False, "msg": msg, "chance": 0}

  // chance = self.draw(draw_type=0, p00001=p00001, p00003=p00003)["chance"]
  // lottery_msgs = self.lottery(p00001=p00001, award_list=[])
  // if chance:
  //     draw_msg = ""
  //     for _ in range(chance):
  //         ret = self.draw(draw_type=1, p00001=p00001, p00003=p00003)
  //         draw_msg += ret["msg"] + ";" if ret["status"] else ""
  // else:
  //     draw_msg = "抽奖机会不足"



// 白金抽奖
function lotto_lottery(P00001){
  // {"code":"Q00397","msg":"次数类型不存在","data":{}}
  // {"code":"A00000","msg":"成功","data":{"giftName":"未中奖","orderCode":"20240505000000000","giftType":29,"giftId":17000,"times":3,"giftLevel":5,"giftCode":"G_9aaaaaaaaaa","ticket":{},"sendType":1,"fillPhone":0,"imageUrl":{},"giftExtendConfig":{},"giftInfos":{}}}
  // {"code":"Q00702","msg":"抽奖次数用完","data":{"giftName":{},"orderCode":{},"giftType":{},"giftId":{},"times":0,"giftLevel":{},"giftCode":{},"ticket":{},"sendType":{},"fillPhone":{},"imageUrl":{},"giftExtendConfig":{},"giftInfos":{}}}
  messageSuccess = "抽奖:"
  messageFail = ""

  give_times(P00001)
  gift_list = []
  for(i=0; i<5; i++)
  {
    url = "https://pcell.iqiyi.com/lotto/lottery"
    params ="?actCode=bcf9d354bc9f677c&P00001=" + P00001
    resp = HTTP.fetch(url + params , {
      method: "get",
    });
    resp = resp.json()
    console.log(resp)
    msg = resp["msg"]
    gift_name = resp["data"]["giftName"]
    content = ""
    if(gift_name == "[object Object]")
    {
      content = "第" + (i+1) + "次" +msg + " "
    }else
    {
      content = "第" + i + "次" +gift_name + " "
    }
    // console.log(gift_name)
    messageSuccess += content
    sleep(2000)
  
  }  

  msg = [messageSuccess, messageFail]
  return msg
}

// 获取 VIP 日常任务 和 taskCode(任务状态)
function query_user_task(P00001){
    // 获取 VIP 日常任务 和 taskCode(任务状态)
    tasklist = []
    url = "https://tc.vip.iqiyi.com/taskCenter/task/queryUserTask"
    // params = {"P00001": P00001}
    params = "?P00001=" + P00001
    task_list = []
    // res = requests.get(url=url, params=params).json()
    res = HTTP.fetch(url + params , {
      method: "post",
    });
    res = res.json()
    // console.log(res)
    // Application.Range("A8").Value = res.text()
    if(res["code"] == "A00000"){
        // 判断是否是vip
        flagvip = res["data"]["userInfo"]["vip"]
        if(flagvip == true) // 是vip
        {
            for(i in res["data"]["tasks"]["daily"]){
              item = res["data"]["tasks"]["daily"][i]
              // console.log(item["taskTitle"])
              task_list[i] = {
                "taskTitle" : item["taskTitle"],
                "taskCode" : item["taskCode"],
                "status" : item["status"],
                "taskReward" : item["taskReward"]["task_reward_growth"],
                }
              
            }
          console.log(task_list)
        }
    }
    return task_list
}

// 完成任务
function join_task(P00001, task_list){
    // 遍历完成任务
    url = "https://tc.vip.iqiyi.com/taskCenter/task/joinTask"
    // params = {
    //     "P00001": P00001,
    //     "taskCode": "",
    //     "platform": "bb136ff4276771f3",
    //     "lang": "zh_CN",
    // }
    params = "?P00001=" + P00001 + "&platform=bb136ff4276771f3&lang=zh_CN" 
    // console.log(task_list)
    for(i in task_list)
    {
      item = task_list[i]
      // console.log(item)
      // if(item["status"] == 2){
        params= params + "&taskCode=" + item["taskCode"]
        // console.log(params)
        res = HTTP.fetch(url + params , {
          method: "get",
        });
        // {"code":"A00000","msg":"成功"}
        // {"code":"Q00401","msg":"任务无效"}
        console.log(res.json())
      // }
    } 
}

function get_task_rewards(P00001, task_list){
  // 获取任务奖励
  messageSuccess = ""
  messageFail = ""

  url = "https://tc.vip.iqiyi.com/taskCenter/task/getTaskRewards"
  // params = {
  //     "P00001": P00001,
  //     "taskCode": "",
  //     "platform": "bb136ff4276771f3",
  //     "lang": "zh_CN",
  // }
  params = "?P00001=" + P00001 + "&platform=bb136ff4276771f3&lang=zh_CN" 
  growth_task = 0
  for(i in task_list){
    item = task_list[i]
    if(item["status"] == 0){
        // params["taskCode"] = item.get("taskCode")
        params= params + "&taskCode=" + item["taskCode"]
        res = HTTP.fetch(url + params , {
          method: "get",
        });
    }else if(item["status"] == 4){
        params["taskCode"] = item.get("taskCode")
        // requests.get(
        //     url="https://tc.vip.iqiyi.com/taskCenter/task/notify", params=params
        // )
        res = HTTP.fetch("https://tc.vip.iqiyi.com/taskCenter/task/notify" + params , {
          method: "get",
        });
    }else if(item["status"] == 1){
        taskTitle = item["taskTitle"]
        growth_task += item["taskReward"]
        content = taskTitle + "完成奖励" + growth_task + "成长值 "
        messageSuccess += content
    }
  }

  msg = [messageSuccess, messageFail]
  return msg
}

// 签到
function signin(url, headers, data){
  messageSuccess = ""
  messageFail = ""

  resp = HTTP.post(
    // url2 + params,
    url,
    JSON.stringify(data),
    { headers: headers }
  );

  if (resp.status == 200) {
    resp = resp.json();
    console.log(resp)
    code = resp["code"]
    
    if(code == "A00000")
    {
      respmsg = resp["data"]["msg"]
      // console.log(msg)
      if(respmsg == "[object Object]")
      {
        respmsg = ""
      }
      signDays = ""
      try{
          signDays = resp["data"]["data"]["signDays"] // 签到天数
      }catch{
          console.log("无法获取到签到天数")
      }

      if(signDays == undefined)
      {
        signDays = ""
        content = respmsg + " "
      }else
      {
        content = "已签到" + signDays + "天 " + respmsg
      }
      
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

  msg = [messageSuccess, messageFail]
  return msg
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
    var url1 = "http://serv.vip.iqiyi.com/vipgrowth/query.action"; // 账号信息查询
    var url2 = "https://community.iqiyi.com/openApi/task/execute"; // VIP 签到

    // 解析cookie
    cookie_json = cookie_to_json(cookie);
    P00001 = cookie_json["P00001"];
    P00002 = cookie_json["P00002"];
    P00003 = cookie_json["P00003"];
    dfp = cookie_json["__dfp"];
    // console.log(P00001)
    // console.log(dfp)
    // console.log(P00003)
    // headers = {"P00001": P00001}

    // headers = {
    //   "Cookie": cookie,
    //   "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/46.0.2486.0 Safari/537.36 Edge/13.10586"

    // };

    // 查询信息
    // params = "?P00001="+P00001
    // resp = HTTP.fetch(url1 + params, {
    //   method: "get",
    //   // headers: headers,
    //   // data: data
    // });
    // // {"messageId":"xxxx","code":"A00000","msg":"成功","data":{"level":"1","growthvalue":"0","todayGrowthValue":0,"speed":0,"distance":300,"days":-1,"exceed":"0.00%"}}

    // var digits = getUUIDDigits(16);
    // console.log(digits);



    // 签到
    // {'code': 'A00000', 'message': '成功执行.', 'validateResult': True, 'data': {'msg': '任务今日完成次数已经到达上限', 'duration': 0, 'code': 'A0014', 'data': None, 'success': False, 'timestamp': 1714400000000}, 'abtest': {}}
    // {"code":"A00000","message":"成功执行.","validateResult":true,"data":{"msg":"qyid和iqid不能都为空","duration":0,"code":"A0005","data":{},"success":false,"timestamp":1714400000000},"abtest":{}}
    // {"code":"A00101","message":"接口参数校验失败."}
    time_stamp = getts13()
    qyid = getUUIDDigits(16)
    data = "agentType=1|agentversion=1|appKey=basic_pcw|authCookie=" + P00001 + "|qyid=" + qyid + "|task_code=natural_month_sign|timestamp=" + time_stamp + "|typeCode=point|userId=" + P00003 + "|UKobMjDMsDoScuWOfp6F"
    sign = getsign(data) // 注意，需要小写sign
    params = "?agentType=1&agentversion=1&appKey=basic_pcw&authCookie=" + P00001 + "&qyid=" + qyid + "&sign=" + sign + "&task_code=natural_month_sign&timestamp=" + time_stamp + "&typeCode=point&userId=" + P00003
    data = {
        "natural_month_sign": {
            "taskCode": "iQIYI_mofhr",
            "agentType": 1,
            "agentversion": 1,
            "authCookie": String(P00001),
            "qyid": qyid,
            "verticalCode": "iQIYI",
        }
    }
    headers={
      "Cookie": "P00001=" + P00001, 
      "Content-Type": "application/json"
    }

    msg = signin(url2+params, headers, data)
    messageSuccess += msg[0]
    messageFail += msg[1]

    // resp = HTTP.fetch(url2 + params , {
    //   method: "post",
    //   headers: headers,
    //   data: data
    // });

    // 抽奖
    msg = lotto_lottery(P00001)
    messageSuccess += msg[0]
    messageFail += msg[1]

    // 做任务,VIP才会做任务
    task_list = query_user_task(P00001)
    join_task(P00001, task_list)
    msg = get_task_rewards(P00001, task_list)
    messageSuccess += msg[0]
    messageFail += msg[1]
    
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
