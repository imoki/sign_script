// WPS自动签到(客户端版)（多用户版，支持bark推送）
// 需配合“金山文档”中的表格内容

// 推送bark消息
function bark(message){
  let bark_push = Application.Range("E"+2).Text
  if(bark_push == "是"){
    let bark_id = Application.Range("D"+2).Text
    let BARK_PUSH = 'https://api.day.app/' + bark_id + "/" + message;
    let barkdata = HTTP.get(BARK_PUSH,
      {headers:{'Content-Type': 'application/x-www-form-urlencoded'}}
    )
    barkdata = barkdata.json()
  }
}

// 推送pushplus消息
function pushplus(message){
  let push = Application.Range("G"+2).Text
  let token = Application.Range("F"+2).Text
  if(push == "是" && token != ""){
    url = 'http://www.pushplus.plus/send?token=' + token + '&content=' + message
    let resp = HTTP.fetch(url, {
      method: "get"
    })
    sleep(5000)
  }
}

// 推送serverchan消息
function serverchan(message){
  let push = Application.Range("I"+2).Text
  let key = Application.Range("H"+2).Text
  if(push == "是" && key != ""){
    url = "https://sctapi.ftqq.com/" + key + ".send"  + "?title=消息推送"  + "&desp=" + message
    let resp = HTTP.fetch(url, {
      method: "get"
    })
    sleep(5000)
  }
}

function sleep(d){
  for(var t = Date.now();Date.now() - t <= d;);
}

var message= "【wps客户端版】";
var line = 21;  // 默认支持20个账户
for (let i = 2; i <= line; i++){
  var cookie = Application.Range("A"+i).Text
  var exec = Application.Range("B"+i).Text
  if(cookie != "" && exec == "是"){
        url = "https://vipapi.wps.cn/wps_clock/v2"
        headers = {
            "Cookie": "wps_sid=" + cookie,
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/46.0.2486.0 Safari/537.36 Edge/13.10586"
        }
        data = {
          'double': 0
        }

        let resp = HTTP.fetch(url, {
          method: "post",
          headers: headers,
          data: data
        })

        try{
          resp = resp.json()
          var result = resp["result"]
          var msg = resp["msg"]
          if(result == "ok")
          {
            message = message + "单元格A" + i + "签到成功，请到客户端手动兑换时长 "
          }else{
            if(msg == "ClockAgent")
            {
              message = message + "单元格A" + i + "已经签到过了 "
            }else{
              message = message + "单元格A" + i + "签到失败 "
            }
          }
        }catch{
          message = message + "单元格A" + i + "签到失败 "
        }
        console.log(resp)
      sleep(2000)
  }
}

bark(message);
pushplus(message);
serverchan(message);
