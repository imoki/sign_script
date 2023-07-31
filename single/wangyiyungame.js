// 网易云游戏自动签到
// 需配合“金山文档”中的表格内容
// 独立脚本

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

var message= "【网易云游戏】";
var line = 21;  // 默认支持20个账户
for (let i = 2; i <= line; i++){
  var cookie = Application.Range("A"+i).Text
  var exec = Application.Range("B"+i).Text
  if(cookie != "" && exec == "是"){
        try{
          url = 'https://n.cg.163.com/api/v2/sign-today'
          headers = {
              'Accept': 'application/json, text/plain, */*',
              'Accept-Encoding': 'gzip, deflate, br',
              'Accept-Language': 'zh-CN,zh;q=0.9,en-US;q=0.8,en;q=0.7,ja-JP;q=0.6,ja;q=0.5',
              'Authorization': cookie,
              'Connection': 'keep-alive',
              'Content-Length': '0',
              'Host': 'n.cg.163.com',
              'Origin': 'https://cg.163.com',
              'Referer': 'https://cg.163.com/',
              'Sec-Fetch-Dest': 'empty',
              'Sec-Fetch-Mode': 'cors',
              'Sec-Fetch-Site': 'same-site',
              'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.198 Safari/537.36',
              'X-Platform': '0'
          }

          let resp = HTTP.fetch(url,{
            method: "post",
            headers: headers
          })
          
          if(resp.status==200){
            message =  message + "单元格A" + i + "签到成功 "
            // 如果需要推送昵称，则改为下面这一行
            // message += Application.Range("C"+i).Text + "签到成功 "
            console.log("单元格A" + i + "签到成功 ")
          }else{
            message =  message + "单元格A" + i + "签到失败或已签到 "
            // 如果需要推送昵称，则改为下面这一行
            // message += Application.Range("C"+i).Text + "签到失败或已签到 "
            console.log("单元格A" + i + "签到失败 ")
          }
        }catch{
          console.log("单元格A" + i + "的cookie有误，请重新填写")
        }
      sleep(2000)
  }
}

bark(message);
pushplus(message);
serverchan(message);
