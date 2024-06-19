// 百度贴吧自动签到（多用户版，支持bark推送）
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

// 获取10 位时间戳
function getts(){
  var ts = Math.round(new Date().getTime()/1000).toString()
  return ts
}

// 获取关注列表需要用到的sign
function getsign(cookie, ts){
  var key_1 = "BDUSS=" + cookie + "_client_id=wappc_1534235498291_488_client_type=2_client_version=9.7.8.0_phone_imei=000000000000000from=1008621ymodel=MI+5net_type=1page_no=1page_size=200timestamp=" + ts + "vcode_tag=11";
  var SIGN_KEY = 'tiebaclient!!!'
  key_1 = key_1 + SIGN_KEY
  var sign = Crypto.createHash("md5").update(key_1,"utf8").digest("hex").toUpperCase().toString()
  return sign
}

// 关注列表data,包含了sign
function getdata(cookie, ts, sign){
  var data_favorite = "BDUSS=" + cookie + "&_client_type=2&_client_id=wappc_1534235498291_488&_client_version=9.7.8.0&_phone_imei=000000000000000&from=1008621y&page_no=1&page_size=200&model=MI%2B5&net_type=1&timestamp=" + ts + "&vcode_tag=11&sign=" + sign
  return data_favorite
}

// 签到需要用到的sign
function getsign2(cookie, ts, tbs, fid, kw){
  var key_1 = "BDUSS=" + cookie + "_client_type=2_client_version=9.7.8.0_phone_imei=000000000000000fid=" +fid + "kw=" + kw + "model=MI+5net_type=1tbs="+tbs+ "timestamp=" + ts
  var SIGN_KEY = 'tiebaclient!!!'
  key_1 = key_1 + SIGN_KEY
  var sign = Crypto.createHash("md5").update(key_1).digest("hex").toUpperCase().toString()
  return sign
}

// 签到用到的data包含了sign
function getsigndata(cookie, tbs, fid, kw){
  var ts = getts();
  var sign = getsign2(cookie, ts, tbs, fid, kw);
  var data = "_client_type=2&_client_version=9.7.8.0&_phone_imei=000000000000000&model=MI%2B5&net_type=1&BDUSS=" + cookie + "&fid=" +fid + "&kw=" + kw + "&tbs="+tbs+ "&timestamp=" + ts + "&sign=" + sign
  return data
}

// 签到请求
function client_sign(cookie, tbs, fid, kw){
    var data = getsigndata(cookie, tbs, fid, kw)
    let res= HTTP.post("http://c.tieba.baidu.com/c/c/forum/sign",
    data
    )
    console.log(res.text())
}

function sleep(d){
  for(var t = Date.now();Date.now() - t <= d;);
}

var message= "【tieba】";
var line = 21;  // 默认支持20个账户
for (let i = 2; i <= line; i++){
  var cookie = Application.Range("A"+i).Text
  var exec = Application.Range("B"+i).Text
  if(cookie != "" && exec == "是"){
        // 获取tbs
        let res = HTTP.fetch("http://tieba.baidu.com/dc/common/tbs", {
          method: "get",
          headers: {
            "Cookie": "BDUSS=" + cookie
          }
        })
        res = res.json()
        var tbs = res["tbs"]
        sleep(1000)

        // 获取关注列表
        var ts = getts();
        var sign = getsign(cookie, ts);
        var data_favorite = getdata(cookie, ts, sign)
        let res_favorite = HTTP.post("http://c.tieba.baidu.com/c/f/forum/like",
          data_favorite
        )
        res_favorite = res_favorite.json()
        console.log(res_favorite['forum_list']["non-gconforum"])
        sleep(1000)

        // 签到
        var arr_favorite = res_favorite['forum_list']["non-gconforum"]
        for(var j = 0; j < arr_favorite.length; j++){
          client_sign(cookie, tbs, arr_favorite[j]["id"], arr_favorite[j]["name"])
          message = message +arr_favorite[j]["name"] + "签到 "
          sleep(20000)
        }
      sleep(2000)
  }
}

bark(message);
pushplus(message);
serverchan(message);
