// 阿里云盘自动签到领取奖励（多用户版，支持bark推送）
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

function sleep(d){
  for(var t = Date.now();Date.now() - t <= d;);
}

var message= ""
var line = 21;  // 指定读取从第2行到第line行的内容
for (let i = 2; i <= line; i++){
  var refresh_token_pos = "A" + i
  var refresh_token = Application.Range(refresh_token_pos).Text
  if(refresh_token != ""){
    // 发起网络请求-获取token
    let res = HTTP.post("https://auth.aliyundrive.com/v2/account/token",
          JSON.stringify({
          "grant_type": "refresh_token",
          "refresh_token":refresh_token
          })
    )
    res = res.json()
    var access_token = res['access_token']

    if  (access_token == undefined){
      var message ="单元格【" + refresh_token_pos + "】的refresh_token值错误，请填写正确的refresh_token值 "
    }else{
      try{
        // 签到
        let res2 = HTTP.post("https://member.aliyundrive.com/v1/activity/sign_in_list",
            JSON.stringify({"_rx-s": "mobile"}),
            {headers:{"Authorization" : 'Bearer '+ access_token}}
        )
        res2=res2.json()
        var signInCount = res2['result']['signInCount']
        var message = message + "账号："+ res["user_name"] + "-签到成功, 本月累计签到" + signInCount + "天 "

      }catch{
        var message ="单元格【" + refresh_token_pos + "】的refresh_token签到失败 "
        return
      }
      sleep(1000)

      // 领取奖励
      var reward = Application.Range("B"+i).Text
      if (reward == "是"){
        try{
          let res3 = HTTP.post("https://member.aliyundrive.com/v1/activity/sign_in_reward?_rx-s=mobile",
            JSON.stringify({
              "signInDay": signInCount
              }),
            {headers:{"Authorization" : 'Bearer '+ access_token}}
          )
          res3 = res3.json()
          var message = message + "签到获得" + res3["result"]["name"] + "," + res3["result"]["description"] + " "
        }catch{
              var message = message + "账号：" + res["user_name"] + "-领取奖励失败 "
        }
      }else{
        message = message +"  奖励待领取 "
      }

      console.log(message)
      // 发送推送消息，若不需要推送消息，则注释掉下面这一行
      bark(message);

    }
  }
}
