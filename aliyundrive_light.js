// 阿里云盘自动签到领取奖励（极简版）
// 无推送功能。若需要推送功能，请使用阿里云盘自动签到（多用户版）

// 将如下的xxxxxxxx替换成自己的refresh_token值
var refresh_token = "xxxxxxxx"

if(refresh_token != ""){
  let res = HTTP.post("https://auth.aliyundrive.com/v2/account/token",
      JSON.stringify({
      "grant_type": "refresh_token",
      "refresh_token":refresh_token
      })
  )
  res = res.json()
  var access_token = res['access_token']

  if(access_token == undefined){
    console.log("refresh_token错误,请重新填写refresh_token")
  }else{
    try{
      let res2 = HTTP.post("https://member.aliyundrive.com/v1/activity/sign_in_list",
            JSON.stringify({"_rx-s": "mobile"}),
            {headers:{"Authorization":'Bearer '+access_token}}
      )
      res2=res2.json()
      var signInCount = res2['result']['signInCount']
      console.log("账号：" + res["user_name"] + "-签到成功, 本月累计签到" + signInCount + "天")
    }catch{
      console.log("refresh_token签到失败")
    }
    sleep(1000)

    // 领取奖励
    try{
      let res3 = HTTP.post("https://member.aliyundrive.com/v1/activity/sign_in_reward?_rx-s=mobile",
        JSON.stringify({
          "signInDay": signInCount
        }),
        {headers:{"Authorization":'Bearer '+access_token}}
      )
      res3=res3.json()
      console.log("签到获得" + res3["result"]["name"] + "," + res3["result"]["description"])
    }catch{
      console.log("领取奖励失败")
    }
  }
}

function sleep(d){
  for(var t = Date.now(); Date.now() - t <= d; );
}
