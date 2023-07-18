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
