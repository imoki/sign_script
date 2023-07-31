// 吾爱破解论坛自动签到（多用户版，支持bark推送）
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


function cookie_to_json(cookies){
    var cookie_text = cookies;
    var arr = [];
    var text_to_split = cookie_text.split(";");
    for(var i in text_to_split){
        var tmp = text_to_split[i].split("=");
        arr.push('"'+tmp.shift().trim()+'":"'+tmp.join(":").trim()+'"')
    };
    var res ='{\n'+arr.join(",\n")+'\n}';
    return JSON.parse(res);
}

function sleep(d){
  for(var t = Date.now();Date.now() - t <= d;);
}

var message= "【52pojie】";
var line = 21;  // 默认支持20个账户
var url1 = "https://www.52pojie.cn/CSPDREL2hvbWUucGhwP21vZD10YXNrJmRvPWRyYXcmaWQ9Mg==?wzwscspd=MC4wLjAuMA=="
var url2 = 'https://www.52pojie.cn/home.php?mod=task&do=apply&id=2&referer=%2F'
var url3 = 'https://www.52pojie.cn/home.php?mod=task&do=draw&id=2'
for (let i = 2; i <= line; i++){
  var cookie = Application.Range("A"+i).Text
  var exec = Application.Range("B"+i).Text
  if(cookie != "" && exec == "是"){
        cookie_json = cookie_to_json(cookie)
        try{
          htVC_2132_saltkey = cookie_json['htVC_2132_saltkey']
          htVC_2132_auth = cookie_json['htVC_2132_auth']
          cookie = "htVC_2132_saltkey=" + htVC_2132_saltkey + "; htVC_2132_auth=" + htVC_2132_auth + ";"
          headers={
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9",
            "Accept-Encoding": "gzip, deflate, br",
            "Accept-Language": "zh-CN,zh;q=0.9",
            "Cache-Control": "no-cache",
            "Connection": "keep-alive",
            "Cookie": cookie,
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/109.0.0.0 Safari/537.36",
          }

          let res = HTTP.fetch(url1,{
            method: "get",
            headers: headers
          })
          cookie_set = res.headers['set-cookie']
          cookie = cookie + cookie_set
          sleep(1000)

          headers["Cookie"] = cookie
          let res2 = HTTP.fetch(url2,{
            method: "get",
            headers: headers
          })
          cookie_set = res2.headers['set-cookie']
          cookie = cookie + cookie_set
          sleep(1000)

          headers["Cookie"] = cookie
          let res3 = HTTP.fetch(url3,{
            method: "get",
            headers: headers
          })
          console.log("签到完成")
          message =  message + "单元格A" + i + "签到 "
          // 若需要修改bark推送的内容为昵称，则将上面一行改为如下的形式
          // message += Application.Range("C"+i).Text + "签到 "
        }catch{
          console.log("单元格A" + i + "的cookie有误，请重新填写")
        }
      sleep(2000)
  }
}

bark(message);
pushplus(message);
serverchan(message);
