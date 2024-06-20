/*
    name: "哔哩哔哩"
    cron: 10 30 12 * * *
    脚本兼容: 金山文档， 青龙
    更新时间：20240620
*/

const logo = "艾默库 : https://github.com/imoki/sign_script"    // 仓库地址
let sheetNameSubConfig = "bilibili"; // 分配置表名称
let pushHeader = "【哔哩哔哩】";
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

// =================青龙适配开始===================

// 艾默库青龙适配代码
// v2.4.0  

var userContent=[["\u0063\u006f\u006f\u006b\u0069\u0065\u0028\u9ed8\u8ba4\u0032\u0030\u4e2a\u0029",")\u5426/\u662F(\u884C\u6267\u5426\u662F".split("").reverse().join(""),"\u8d26\u53f7\u540d\u79f0\u0028\u53ef\u4e0d\u586b\u5199\u0029"]];var configContent=[["\u79F0\u540D\u7684\u8868\u4F5C\u5DE5".split("").reverse().join(""),"\u5907\u6ce8","\u53ea\u63a8\u9001\u5931\u8d25\u6d88\u606f\uff08\u662f\u002f\u5426\uff09","\u63a8\u9001\u6635\u79f0\uff08\u662f\u002f\u5426\uff09"],[sheetNameSubConfig,pushHeader,"\u5426","\u662f"]];var qlpushFlag=0xdb0d5^0xdb0d5;var qlSheet=[];var colNum=["\u0041","\u0042","\u0043","\u0044","\u0045","\u0046","\u0047","\u0048","\u0049","\u004a","\u004b","\u004c",'M',"\u004e","\u004f","\u0050",'Q'];qlConfig={'CONFIG':configContent,'SUBCONFIG':userContent};var posHttp=0x9809b^0x9809b;var flagFinish=0x9a483^0x9a483;var flagResultFinish=0x18a91^0x18a91;var HTTPOverwrite={'get':function get(_0x4d5ae9,_0x277b01){_0x277b01=_0x277b01["\u0068\u0065\u0061\u0064\u0065\u0072\u0073"];let _0x24c742=userContent["\u006c\u0065\u006e\u0067\u0074\u0068"]-qlpushFlag;method="\u0067\u0065\u0074";resp=fetch(_0x4d5ae9,{'method':method,"\u0068\u0065\u0061\u0064\u0065\u0072\u0073":_0x277b01})['then'](function(_0x1587f7){return _0x1587f7['text']()['then'](_0x14e094=>{return{"\u0073\u0074\u0061\u0074\u0075\u0073":_0x1587f7['status'],"\u0068\u0065\u0061\u0064\u0065\u0072\u0073":_0x1587f7["\u0068\u0065\u0061\u0064\u0065\u0072\u0073"],"\u0074\u0065\u0078\u0074":_0x14e094,"\u0072\u0065\u0073\u0070\u006f\u006e\u0073\u0065":_0x1587f7,"\u0070\u006f\u0073":_0x24c742};});})["\u0074\u0068\u0065\u006e"](function(_0x40b681){try{data=JSON["\u0070\u0061\u0072\u0073\u0065"](_0x40b681["\u0074\u0065\u0078\u0074"]);return{"\u0073\u0074\u0061\u0074\u0075\u0073":_0x40b681['status'],'headers':_0x40b681['headers'],'json':function _0x17a182(){return data;},'text':function _0x249d49(){return _0x40b681['text'];},"\u0070\u006f\u0073":_0x40b681["\u0070\u006f\u0073"]};}catch(_0x19c001){return{"\u0073\u0074\u0061\u0074\u0075\u0073":_0x40b681["\u0073\u0074\u0061\u0074\u0075\u0073"],'headers':_0x40b681["\u0068\u0065\u0061\u0064\u0065\u0072\u0073"],"\u006a\u0073\u006f\u006e":null,'text':function _0x196b5c(){return _0x40b681["\u0074\u0065\u0078\u0074"];},'pos':_0x40b681["\u0070\u006f\u0073"]};}})['then'](_0x575f1d=>{_0x24c742=_0x575f1d["\u0070\u006f\u0073"];flagResultFinish=resultHandle(_0x575f1d,_0x24c742);if(flagResultFinish==(0x1990c^0x1990d)){i=_0x24c742+(0x6733e^0x6733f);for(;i<=line;i++){var _0x21fa27=Application["\u0052\u0061\u006e\u0067\u0065"]('A'+i)['Text'];var _0xf32cc4=Application['Range']("\u0042"+i)["\u0054\u0065\u0078\u0074"];if(_0x21fa27=="".split("").reverse().join("")){break;}if(_0xf32cc4=='是'){console['log']('🧑\x20开始执行用户：'+(parseInt(i)-(0x2a98a^0x2a98b)));flagResultFinish=0x273b7^0x273b7;execHandle(_0x21fa27,i);break;}}}if(_0x24c742==userContent['length']&&flagResultFinish==(0xd629e^0xd629f)){flagFinish=0x25089^0x25088;}if(qlpushFlag==0x0&&flagFinish==0x1){console['log']("\u9001\u63A8\u8D77\u53D1\u9F99\u9752 \uDE80\uD83D".split("").reverse().join(""));message=messageMerge();const{sendNotify:_0x13202b}=require('./sendNotify.js');_0x13202b(pushHeader,message);qlpushFlag=-(0x31873^0x31817);}})['catch'](_0x3e566b=>{console["\u0065\u0072\u0072\u006f\u0072"]('Fetch\x20error:',_0x3e566b);});},'post':function post(_0x5ee0bd,_0x5c53b8,_0x1dd893,_0x2ca2cc){_0x1dd893=_0x1dd893['headers'];contentType=_0x1dd893["\u0043\u006f\u006e\u0074\u0065\u006e\u0074\u002d\u0054\u0079\u0070\u0065"];contentType2=_0x1dd893["\u0063\u006f\u006e\u0074\u0065\u006e\u0074\u002d\u0074\u0079\u0070\u0065"];var _0x1a6999="".split("").reverse().join("");if(contentType!=undefined&&contentType!="".split("").reverse().join("")||contentType2!=undefined&&contentType2!="".split("").reverse().join("")){if(contentType=="dedocnelru-mrof-www-x/noitacilppa".split("").reverse().join("")){console['log']("\u5F0F\u683C\u5355\u8868 :\u4E3A\u4F53\u6C42\u8BF7\u9001\u53D1\u5230\u6D4B\u68C0 \uDF73\uD83C".split("").reverse().join(""));_0x1a6999=dataToFormdata(_0x5c53b8);}else{try{console['log']("\u5F0F\u683CNOSJ :\u4E3A\u4F53\u6C42\u8BF7\u9001\u53D1\u5230\u6D4B\u68C0 \uDF73\uD83C".split("").reverse().join(""));_0x1a6999=JSON["\u0073\u0074\u0072\u0069\u006e\u0067\u0069\u0066\u0079"](_0x5c53b8);}catch{console["\u006c\u006f\u0067"]('🍳\x20检测到发送请求体为:\x20表单格式');_0x1a6999=_0x5c53b8;}}}else{console["\u006c\u006f\u0067"]('🍳\x20检测到发送请求体为:\x20JSON格式');_0x1a6999=JSON['stringify'](_0x5c53b8);}if(_0x2ca2cc=='get'||_0x2ca2cc=="\u0047\u0045\u0054"){let _0x4ff5dd=userContent["\u006c\u0065\u006e\u0067\u0074\u0068"]-qlpushFlag;method='get';resp=fetch(_0x5ee0bd,{'method':method,"\u0068\u0065\u0061\u0064\u0065\u0072\u0073":_0x1dd893})['then'](function(_0x40c629){return _0x40c629["\u0074\u0065\u0078\u0074"]()["\u0074\u0068\u0065\u006e"](_0x495e66=>{return{"\u0073\u0074\u0061\u0074\u0075\u0073":_0x40c629['status'],"\u0068\u0065\u0061\u0064\u0065\u0072\u0073":_0x40c629['headers'],"\u0074\u0065\u0078\u0074":_0x495e66,'response':_0x40c629,"\u0070\u006f\u0073":_0x4ff5dd};});})["\u0074\u0068\u0065\u006e"](function(_0x53dc7b){try{_0x5c53b8=JSON['parse'](_0x53dc7b['text']);return{'status':_0x53dc7b['status'],'headers':_0x53dc7b['headers'],'json':function _0x51d78b(){return _0x5c53b8;},'text':function _0x1ea301(){return _0x53dc7b['text'];},'pos':_0x53dc7b['pos']};}catch(_0x256b1e){return{'status':_0x53dc7b["\u0073\u0074\u0061\u0074\u0075\u0073"],"\u0068\u0065\u0061\u0064\u0065\u0072\u0073":_0x53dc7b['headers'],'json':null,"\u0074\u0065\u0078\u0074":function _0x15fa3a(){return _0x53dc7b['text'];},'pos':_0x53dc7b['pos']};}})["\u0074\u0068\u0065\u006e"](_0x4ad3a2=>{_0x4ff5dd=_0x4ad3a2['pos'];flagResultFinish=resultHandle(_0x4ad3a2,_0x4ff5dd);if(flagResultFinish==0x1){i=_0x4ff5dd+0x1;for(;i<=line;i++){var _0xfe207b=Application['Range']('A'+i)['Text'];var _0x337b12=Application['Range']("\u0042"+i)["\u0054\u0065\u0078\u0074"];if(_0xfe207b==''){break;}if(_0x337b12=='是'){console['log']("\uFF1A\u6237\u7528\u884C\u6267\u59CB\u5F00 \uDDD1\uD83E".split("").reverse().join("")+(parseInt(i)-0x1));flagResultFinish=0x0;execHandle(_0xfe207b,i);break;}}}if(_0x4ff5dd==userContent['length']&&flagResultFinish==(0x65592^0x65593)){flagFinish=0x1;}if(qlpushFlag==(0x28748^0x28748)&&flagFinish==0x1){console['log']('🚀\x20青龙发起推送');message=messageMerge();const{sendNotify:_0x45c282}=require("\u002e\u002f\u0073\u0065\u006e\u0064\u004e\u006f\u0074\u0069\u0066\u0079\u002e\u006a\u0073");_0x45c282(pushHeader,message);qlpushFlag=-0x64;}})['catch'](_0x37ca1f=>{console['error'](":rorre hcteF".split("").reverse().join(""),_0x37ca1f);});}else{let _0x4cbbf1=userContent["\u006c\u0065\u006e\u0067\u0074\u0068"]-qlpushFlag;method='post';resp=fetch(_0x5ee0bd,{'method':method,'headers':_0x1dd893,'body':_0x1a6999})["\u0074\u0068\u0065\u006e"](function(_0x194f78){return _0x194f78['text']()['then'](_0x1e3939=>{return{'status':_0x194f78['status'],'headers':_0x194f78['headers'],'text':_0x1e3939,'response':_0x194f78,'pos':_0x4cbbf1};});})["\u0074\u0068\u0065\u006e"](function(_0x29586c){try{_0x5c53b8=JSON["\u0070\u0061\u0072\u0073\u0065"](_0x29586c['text']);return{'status':_0x29586c["\u0073\u0074\u0061\u0074\u0075\u0073"],'headers':_0x29586c['headers'],'json':function _0x5782f8(){return _0x5c53b8;},'text':function _0x466d53(){return _0x29586c['text'];},'pos':_0x29586c['pos']};}catch(_0x3a28fe){return{'status':_0x29586c['status'],'headers':_0x29586c['headers'],'json':null,"\u0074\u0065\u0078\u0074":function _0x4f0f06(){return _0x29586c['text'];},'pos':_0x29586c["\u0070\u006f\u0073"]};}})['then'](_0x2a6953=>{_0x4cbbf1=_0x2a6953['pos'];flagResultFinish=resultHandle(_0x2a6953,_0x4cbbf1);if(flagResultFinish==(0x99a44^0x99a45)){i=_0x4cbbf1+(0xf3437^0xf3436);for(;i<=line;i++){var _0x49c979=Application['Range']("\u0041"+i)['Text'];var _0x4300e0=Application['Range']("\u0042"+i)['Text'];if(_0x49c979==''){break;}if(_0x4300e0=="\u662f"){console['log']('🧑\x20开始执行用户：'+(parseInt(i)-0x1));flagResultFinish=0x5c079^0x5c079;execHandle(_0x49c979,i);break;}}}if(_0x4cbbf1==userContent['length']&&flagResultFinish==(0x29a4a^0x29a4b)){flagFinish=0x1;}if(qlpushFlag==0x0&&flagFinish==0x1){console['log']('🚀\x20青龙发起推送');let _0x3f8f2f=messageMerge();const{sendNotify:_0x3d6c5c}=require('./sendNotify.js');_0x3d6c5c(pushHeader,_0x3f8f2f);qlpushFlag=-0x64;}})['catch'](_0x4df542=>{console['error'](":rorre hcteF".split("").reverse().join(""),_0x4df542);});}}};var ApplicationOverwrite={'Range':function Range(_0x810fe5){charFirst=_0x810fe5['substring'](0x0,0x1);qlRow=_0x810fe5['substring'](0x6c194^0x6c195,_0x810fe5["\u006c\u0065\u006e\u0067\u0074\u0068"]);qlCol=0x1;for(num in colNum){if(colNum[num]==charFirst){break;}qlCol+=0xe3795^0xe3794;}try{result=qlSheet[qlRow-0x1][qlCol-(0xe7d58^0xe7d59)];}catch{result='';}dict={'Text':result};return dict;},"\u0053\u0068\u0065\u0065\u0074\u0073":{'Item':function(_0x1cb341){return{'Name':_0x1cb341,'Activate':function(){flag=0x1;qlSheet=qlConfig[_0x1cb341];if(qlSheet==undefined){qlSheet=qlConfig['SUBCONFIG'];}console['log']("\uFF1A\u8868\u4F5C\u5DE5\u6D3B\u6FC0\u9F99\u9752 \uDF73\uD83C".split("").reverse().join("")+_0x1cb341);return flag;}};}}};var CryptoOverwrite={'createHash':function createHash(_0x2d9616){return{'update':function _0x21e64b(_0x2fa01e,_0x239590){return{"\u0064\u0069\u0067\u0065\u0073\u0074":function _0x1a574f(_0x3676e2){return{'toUpperCase':function _0x228aa2(){return{"\u0074\u006f\u0053\u0074\u0072\u0069\u006e\u0067":function _0x93bcc6(){let _0x316696=require('crypto-js');let _0x21631d=_0x316696['MD5'](_0x2fa01e)['toString']();_0x21631d=_0x21631d['toUpperCase']();return _0x21631d;}};},'toString':function _0x57a8d9(){const _0x481928=require('crypto-js');const _0x1612d5=_0x481928['MD5'](_0x2fa01e)['toString']();return _0x1612d5;}};}};}};}};function dataToFormdata(_0x48ce47){result="".split("").reverse().join("");values=Object['values'](_0x48ce47);values['forEach']((_0x2620b1,_0x37be5d)=>{key=Object['keys'](_0x48ce47)[_0x37be5d];content=key+'='+_0x2620b1+'&';result+=content;});result=result['substring'](0xe8d17^0xe8d17,result["\u006c\u0065\u006e\u0067\u0074\u0068"]-0x1);console['log'](result);return result;}function cookiesTocookieMin(_0x30170f){let _0x11e905=_0x30170f;let _0x945bb7=[];var _0x2ef727=_0x11e905['split']("\u0023");for(let _0x145676 in _0x2ef727){_0x945bb7[_0x145676]=_0x2ef727[_0x145676];}return _0x945bb7;}function checkEscape(_0x5e4c34,_0x3d3504){cookieArrynew=[];j=0x0;for(i=0x0;i<_0x5e4c34['length'];i++){result=_0x5e4c34[i];lastChar=result['substring'](result["\u006c\u0065\u006e\u0067\u0074\u0068"]-(0xe0507^0xe0506),result["\u006c\u0065\u006e\u0067\u0074\u0068"]);if(lastChar=='\x5c'&&i<=_0x5e4c34["\u006c\u0065\u006e\u0067\u0074\u0068"]-(0x18942^0x18940)){console['log']("\u7B26\u5B57\u4E49\u8F6C\u5230\u6D4B\u68C0 \uDF73\uD83C".split("").reverse().join(""));cookieArrynew[j]=result['substring'](0x0,result['length']-0x1)+_0x3d3504+_0x5e4c34[parseInt(i)+0x1];i+=0x1;}else{cookieArrynew[j]=_0x5e4c34[i];}j+=0x1;}return cookieArrynew;}function cookiesTocookie(_0x3bc1df){let _0x573053=_0x3bc1df;let _0x1bc4ee=[];let _0x467c4a=[];let _0x2d24c1=_0x573053['split']('@');_0x2d24c1=checkEscape(_0x2d24c1,'@');for(let _0x3942c3 in _0x2d24c1){_0x467c4a=[];let _0x4014e2=Number(_0x3942c3)+0x1;_0x1bc4ee=cookiesTocookieMin(_0x2d24c1[_0x3942c3]);_0x1bc4ee=checkEscape(_0x1bc4ee,"\u0023");_0x467c4a['push'](_0x1bc4ee[0x6527f^0x6527f]);_0x467c4a['push']('是');_0x467c4a['push']("\u79F0\u6635".split("").reverse().join("")+_0x4014e2);if(_0x1bc4ee["\u006c\u0065\u006e\u0067\u0074\u0068"]>0x0){for(let _0x2c94be=0x3;_0x2c94be<_0x1bc4ee['length']+0x2;_0x2c94be++){_0x467c4a['push'](_0x1bc4ee[_0x2c94be-(0x4b7ae^0x4b7ac)]);}}userContent['push'](_0x467c4a);}qlpushFlag=userContent['length']-0x1;}var qlSwitch=0x0;try{qlSwitch=process['env'][sheetNameSubConfig];qlSwitch=0x7cd16^0x7cd17;console['log']('♻️\x20当前环境为青龙');console['log']('♻️\x20开始适配青龙环境，执行青龙代码');try{fetch=require('node-fetch');console['log']('♻️\x20系统无fetch，已进行node-fetch引入');}catch{console['log']('♻️\x20系统已有原生fetch');}Crypto=CryptoOverwrite;let flagwarn=0x0;const a='da11990c';const b="0b854f216a9662fb".split("").reverse().join("");encode=getsign(logo);let len=encode['length'];if(a+'ec4dce09'==encode['substring'](0x3de44^0x3de44,len/(0xb6827^0xb6825))&&b==encode['substring'](0x4*(0x4e95f^0x4e95b),len)){console['log']('✨\x20'+logo);cookies=process['env'][sheetNameSubConfig];}else{console["\u006c\u006f\u0067"]('🔨\x20请使用艾默库代码\x20:\x20https://github.com/imoki/sign_script');flagwarn=0xd711f^0xd711e;}let flagwarn2=0x49f1d^0x49f1c;const welcome='Welcome\x20to\x20use\x20MOKU\x20code';const mo=welcome["\u0073\u006c\u0069\u0063\u0065"](0xf,0xb3e06^0xb3e17)['toLowerCase']();const ku=welcome['split']('\x20')[0x4-0x1]['slice'](0x2,0x4);if(mo['substring'](0x0,0x1)=='m'){if(ku=='KU'){if(mo['substring'](0x1,0xb0f02^0xb0f00)==String['fromCharCode'](0x6f)){cookiesTocookie(cookies);flagwarn2=0x0;console['log']('💗\x20'+welcome);}}}let t=Date['now']();if(t>0xaa*0x186a0*0x186a0+0x45f34a08e){console['log']("\u63A5\u94FEnoiton\u5E93\u4ED3\u770B\u67E5\u8BF7\u7A0B\u6559\u7528\u4F7F \uDDFE\uD83E".split("").reverse().join(""));Application=ApplicationOverwrite;}else{flagwarn=0x3c83f^0x3c83e;}if(Date['now']()<(0x4048d^0x40445)*0x186a0*0x186a0){console['log']('🤝\x20欢迎各种形式的贡献');HTTP=HTTPOverwrite;}else{flagwarn2=0xb5c8c^0xb5c8d;}if(flagwarn==0x1||flagwarn2==0x1){console['log']('🔨\x20请使用艾默库代码\x20:\x20https://github.com/imoki/sign_script');}}catch{qlSwitch=0x0;console['log']('♻️\x20当前环境为金山文档');console['log']("\u7801\u4EE3\u6863\u6587\u5C71\u91D1\u884C\u6267\uFF0C\u6863\u6587\u5C71\u91D1\u914D\u9002\u59CB\u5F00 \uFE0F\u267B".split("").reverse().join(""));}

// =================青龙适配结束===================

// =================金山适配开始===================
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
    console.log("🍳 消息为空不推送");
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
  // console.log("🍳 已发送邮件至：" + sender);
  console.log("🍳 已发送邮件");
  sleep(5000);
}

// 邮箱配置
function emailConfig() {
  console.log("🍳 开始读取邮箱配置");
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
          console.log("🍳 开始读取邮箱表");
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

// =================金山适配结束===================

// =================共用开始===================
flagConfig = ActivateSheet(sheetNameConfig); // 激活推送表
// 主配置工作表存在
if (flagConfig == 1) {
  console.log("🍳 开始读取主配置表");
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
        console.log("🍳 只推送错误消息");
      }

      if (nickname == "是") {
        messageNickname = 1;
        console.log("🍳 单元格用昵称替代");
      }

      break; // 提前退出，提高效率
    }
  }
}

flagPush = ActivateSheet(sheetNamePush); // 激活推送表
// 推送工作表存在
if (flagPush == 1) {
  console.log("🍳 开始读取推送工作表");
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
  console.log("🍳 开始读取分配置表");

    if(qlSwitch != 1){  // 金山文档
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
    }else{
        for (let i = 2; i <= line; i++) {
            var cookie = Application.Range("A" + i).Text;
            var exec = Application.Range("B" + i).Text;
            if (cookie == "") {
                // 如果为空行，则提前结束读取
                break;
            }
            if (exec == "是") {
                console.log("🧑 开始执行用户：" + "1" )
                execHandle(cookie, i);
                break;  // 只取一个
            }
        } 
    }

}

// 激活工作表函数
function ActivateSheet(sheetName) {
    let flag = 0;
    try {
      // 激活工作表
      let sheet = Application.Sheets.Item(sheetName);
      sheet.Activate();
      console.log("🥚 激活工作表：" + sheet.Name);
      flag = 1;
    } catch {
      flag = 0;
      console.log("🍳 无法激活工作表，工作表可能不存在");
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

// 将消息数组融合为一条总消息
function messageMerge(){
    // console.log(messageArray)
    let message = ""
  for(i=0; i<messageArray.length; i++){
    if(messageArray[i] != "" && messageArray[i] != null)
    {
      message += messageHeader[i] + messageArray[i] + ""; // 加上推送头
    }
  }
  if(message != "")
  {
    console.log("✨✨✨✨✨✨✨✨✨✨✨✨✨✨✨✨✨✨")
    console.log(message)  // 打印总消息
    console.log("✨✨✨✨✨✨✨✨✨✨✨✨✨✨✨✨✨✨")
  }
  return message
}

function sleep(d) {
  for (var t = Date.now(); Date.now() - t <= d; );
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

// =================共用结束===================

// 青龙适配
// 结果处理函数
function resultHandle(resp, pos){
    // 每次进来resultHandle则加一次请求
    posHttp += 1    // 青龙适配，青龙微适配
    // console.log(posHttp)

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
    messageHeader[posLabel] = "👨‍🚀 " + messageName

    if (resp.status == 200) {
        // {"code":1011040,"message":"今日已签到过,无法重复签到","ttl":1,"data":{}}
        // { code: -101, message: '账号未登录', ttl: 1 }
        // {
        //     code: 0,
        //     message: '0',
        //     ttl: 1,
        //     data: {
        //         text: '3000点用户经验,2根辣条',
        //         specialText: '再签到4天可以获得666银瓜子',
        //         allDays: 30,
        //         hadSignDays: 1,
        //         isBonusDay: 0
        //     }
        // }
        resp = resp.json()
        console.log(resp)
        code = resp["code"]
        if (code != null){
        content = ""
        if (code == 0){
            content = "🎉 " + resp["data"]["text"] + "\n"
            // console.log(resp["data"]["text"])
            }else{
                content = "📢 " + resp["message"] + "\n"
                // console.log(resp["message"])
            }
            messageSuccess += content;
            console.log(content)
        }else
        {
            content =  "❌ " + "签到失败\n"
            messageFail += content;
            console.log(content);
        }
      
    } else {
        //   console.log(resp.text());
        content =  "❌ " + "签到失败\n";
        messageFail += content
        console.log(content);
    }

    // 青龙适配，青龙微适配
    flagResultFinish = 1; // 签到结束

  // 以最后一次的取到的消息为主
    if (messageOnlyError == 1) {
        messageArray[posLabel] = messageFail;
    } else {
        messageArray[posLabel] = messageFail + " " + messageSuccess;
    }

    if(messageArray[posLabel] != "")
    {
        if(messageFail != ""){
            messageArray[posLabel] = messageFail + " " + messageSuccess;
        }else{
            messageArray[posLabel] = messageSuccess;
        }
    }

    // console.log(messageArray)

    return flagResultFinish
}


// 具体的执行函数
function execHandle(cookie, pos) {
    // 清零操作，保证不同用户的消息的独立
    // 青龙适配，青龙微适配
    posHttp = 0 // 置空请求
    qlpushFlag -= 1 // 一个用户只会执行一次execHandle，因此可用于记录当前用户
    messageSuccess = "";
    messageFail = "";

  // try {
    var url = "https://api.live.bilibili.com/xlive/web-ucenter/v1/sign/DoSign"; // 直播签到

    headers = {
        'Cookie':cookie,
        'User-Agent': 'HD1910(Android/7.1.2) (pediy.UNICFBC0DD/1.0.5) Weex/0.26.0 720x1280',
    }

    if(qlSwitch != 1){  // 金山文档
        // 直播签到
        resp = HTTP.get(url, {
            headers: headers,
        })
    }else{  // 青龙
        data = {}
        option = "get"
        resp = HTTP.post(
            url,
            data,
            { headers: headers },
            option
        );
    }
    
    // 青龙适配，青龙微适配
    if(qlSwitch != 1){  // 选择金山文档
        resultHandle(resp, pos)
    }
    
}