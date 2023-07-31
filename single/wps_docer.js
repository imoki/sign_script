// WPS自动签到(稻壳版)
// 需配合“金山文档”中的表格内容

// 推送bark消息
function bark(message){
  let push = Application.Range("E"+2).Text
  let bark_id = Application.Range("D"+2).Text
  if(push == "是" && bark_id != ""){
    let url = 'https://api.day.app/' + bark_id + "/" + message;
    // 若需要修改推送的分组，则将上面一行改为如下的形式
    // let url = 'https://api.day.app/' + bark_id + "/" + message + "?group=分组名";
    let resp = HTTP.get(url,
      {headers:{'Content-Type': 'application/x-www-form-urlencoded'}}
    )
    sleep(5000)
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

var message= "【wps稻壳】";
var line = 21;  // 默认支持20个账户
for (let i = 2; i <= line; i++){
  var cookie = Application.Range("A"+i).Text
  var exec = Application.Range("B"+i).Text
    if(cookie != "" && exec == "是"){
        let flagSign = 0; // 签到标志
        url = "https://welfare.docer.wps.cn/sign_in/v1/user_sign_in"  // 签到
        url2 = "https://welfare.docer.wps.cn/common/v1/sign_base_info"  // 获取pptid
        url3 = "https://welfare.docer.wps.cn/sign_in/v1/sign_in_mb/receive_mb_today"  // 领取ppt
        headers = {
            "Cookie": "wps_sid=" + cookie,
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/46.0.2486.0 Safari/537.36 Edge/13.10586"
        }
        data = {
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
            message = message + "单元格A" + i + "签到成功 "
            flagSign = 1;
          }else{
            if(msg == "had sign in")
            {
              message = message + "单元格A" + i + "已经签到过了 "
              flagSign = 1;
            }else{
              message = message + "单元格A" + i + "签到失败 "
            }
          }

          
          console.log(resp)

          // 领取每日ppt
          if(flagSign == 1){
            // 请求获取ppt_id
            resp = HTTP.fetch(url2, {
              method: "get",
              headers: headers
            })

            try{
              resp = resp.json()
              console.log(resp)
              let ppt_id = resp["data"]["today_ppt"]["ppt_id"]
              let ppt_title = resp["data"]["today_ppt"]["title"]
              console.log(ppt_title)
              data = {"ppt_id":ppt_id}
              // 请求领取ppt
              resp = HTTP.fetch(url3 +"?ppt_id=" + ppt_id, {
                method: "post",
                headers: headers,
                data: data
              })
              try{
                resp = resp.json()
                let result = resp["result"]
                let msg = resp["msg"]
                if(result == "ok"){
                  console.log("成功领取每日ppt")
                  message += "成功领取 " + ppt_title
                }else{
                  console.log(msg)
                  message += msg
                }
                console.log(resp)
              }catch{
                console.log(resp.text())
                console.log("领取ppt失败")
                message += "领取ppt失败 "
              }
            }catch{
              console.log(resp.text())
              console.log("获取ppt_id失败")
              message += "无法获取ppt_id,领取ppt失败 "
            }
          }

        }catch{
          message = message + "单元格A" + i + "签到失败 "
        }
      sleep(2000)
  }
}

bark(message);
pushplus(message);
serverchan(message);
