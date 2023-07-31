// 推送bark消息
function bark(message){
  let push = Application.Range("E"+2).Text
  let bark_id = Application.Range("D"+2).Text
  if(push == "是" && bark_id != ""){
    let url = 'https://api.day.app/' + bark_id + "/" + message;
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

var message= "【有道云笔记】";
var line = 21;  // 默认支持20个账户
var url1 = 'https://note.youdao.com/yws/mapi/user?method=checkin'
for (let i = 2; i <= line; i++){
  var cookie = Application.Range("A"+i).Text
  var exec = Application.Range("B"+i).Text
  if(cookie != "" && exec == "是"){
        try{
          headers = {
            'cookie': cookie,
            'User-Agent': 'YNote',
            'Host': 'note.youdao.com'
          }

          let resp = HTTP.fetch(url1,{
            method: "post",
            headers: headers
          })

          if (resp.status == 200) {
              resp = resp.json()
              total = resp['total'] / 1048576
              space = resp['space'] / 1048576
              message += '帐号：' + "单元格A" + i + '签到成功，本次获取 ' + space + ' M, 总共获取 ' + total + ' M '
              console.log('帐号：' + "单元格A" + i + '签到成功，本次获取 ' + space + ' M, 总共获取 ' + total + ' M ')
          }else
          {
            message += '帐号：' + "单元格A" + i + '签到失败 '
            console.log('帐号：' + "单元格A" + i + '签到失败 ')
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
