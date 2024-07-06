/*
    name: "夸克网盘"
    cron: 10 30 10 * * *
    脚本兼容: 金山文档， 青龙
    更新时间：20240706
*/

const logo = "艾默库 : https://github.com/imoki/sign_script"    // 仓库地址
let sheetNameSubConfig = "quark"; // 分配置表名称
let pushHeader = "【夸克网盘】";
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
// v2.4.1  

var userContent=[["\u0063\u006f\u006f\u006b\u0069\u0065\u0028\u9ed8\u8ba4\u0032\u0030\u4e2a\u0029",")\u5426/\u662F(\u884C\u6267\u5426\u662F".split("").reverse().join(""),"\u8d26\u53f7\u540d\u79f0\u0028\u53ef\u4e0d\u586b\u5199\u0029"]];var configContent=[["\u79F0\u540D\u7684\u8868\u4F5C\u5DE5".split("").reverse().join(""),"\u6CE8\u5907".split("").reverse().join(""),"\u53ea\u63a8\u9001\u5931\u8d25\u6d88\u606f\uff08\u662f\u002f\u5426\uff09","\uFF09\u5426/\u662F\uFF08\u79F0\u6635\u9001\u63A8".split("").reverse().join("")],[sheetNameSubConfig,pushHeader,"\u5426","\u662f"]];var qlpushFlag=0x3a049^0x3a049;var qlSheet=[];var colNum=["\u0041",'B',"\u0043",'D',"\u0045","\u0046",'G','H','I',"\u004a",'K',"\u004c","\u004d","\u004e","\u004f",'P',"\u0051"];qlConfig={'CONFIG':configContent,'SUBCONFIG':userContent};var posHttp=0x7f376^0x7f376;var flagFinish=0x688a5^0x688a5;var flagResultFinish=0x34666^0x34666;var HTTPOverwrite={'get':function get(_0x58ee64,_0x23e74e){_0x23e74e=_0x23e74e['headers'];let _0xad7bf=userContent["\u006c\u0065\u006e\u0067\u0074\u0068"]-qlpushFlag;method="teg".split("").reverse().join("");resp=fetch(_0x58ee64,{"\u006d\u0065\u0074\u0068\u006f\u0064":method,'headers':_0x23e74e})["\u0074\u0068\u0065\u006e"](function(_0x1c5326){return _0x1c5326['text']()['then'](_0x38f13f=>{return{'status':_0x1c5326['status'],"\u0068\u0065\u0061\u0064\u0065\u0072\u0073":_0x1c5326["\u0068\u0065\u0061\u0064\u0065\u0072\u0073"],"\u0074\u0065\u0078\u0074":_0x38f13f,'response':_0x1c5326,"\u0070\u006f\u0073":_0xad7bf};});})["\u0074\u0068\u0065\u006e"](function(_0x5824b7){try{data=JSON["\u0070\u0061\u0072\u0073\u0065"](_0x5824b7['text']);return{"\u0073\u0074\u0061\u0074\u0075\u0073":_0x5824b7['status'],"\u0068\u0065\u0061\u0064\u0065\u0072\u0073":_0x5824b7["\u0068\u0065\u0061\u0064\u0065\u0072\u0073"],'json':function _0xe8691(){return data;},'text':function _0x29967f(){return _0x5824b7['text'];},'pos':_0x5824b7["\u0070\u006f\u0073"]};}catch(_0x1c7983){return{"\u0073\u0074\u0061\u0074\u0075\u0073":_0x5824b7["\u0073\u0074\u0061\u0074\u0075\u0073"],'headers':_0x5824b7["\u0068\u0065\u0061\u0064\u0065\u0072\u0073"],'json':null,"\u0074\u0065\u0078\u0074":function _0x489a51(){return _0x5824b7["\u0074\u0065\u0078\u0074"];},"\u0070\u006f\u0073":_0x5824b7['pos']};}})['then'](_0x377aa8=>{_0xad7bf=_0x377aa8['pos'];flagResultFinish=resultHandle(_0x377aa8,_0xad7bf);if(flagResultFinish==(0x425f6^0x425f7)){i=_0xad7bf+(0x41450^0x41451);for(;i<=line;i++){var _0xf812eb=Application["\u0052\u0061\u006e\u0067\u0065"]('A'+i)["\u0054\u0065\u0078\u0074"];var _0x5c1cff=Application['Range']("\u0042"+i)["\u0054\u0065\u0078\u0074"];if(_0xf812eb=="".split("").reverse().join("")){break;}if(_0x5c1cff=="\u662f"){console['log']("\uFF1A\u6237\u7528\u884C\u6267\u59CB\u5F00 \uDDD1\uD83E".split("").reverse().join("")+(parseInt(i)-(0x2920c^0x2920d)));flagResultFinish=0x9ee2d^0x9ee2d;execHandle(_0xf812eb,i);break;}}}if(_0xad7bf==userContent['length']&&flagResultFinish==(0xa66cc^0xa66cd)){flagFinish=0xb4227^0xb4226;}if(qlpushFlag==(0x559b4^0x559b4)&&flagFinish==0x1){console["\u006c\u006f\u0067"]("\u9001\u63A8\u8D77\u53D1\u9F99\u9752 \uDE80\uD83D".split("").reverse().join(""));message=messageMerge();const{sendNotify:_0x581347}=require('./sendNotify.js');_0x581347(pushHeader,message);qlpushFlag=-0x64;}})['catch'](_0x347133=>{console['error']('Fetch\x20error:',_0x347133);});},'post':function post(_0x1f5035,_0x2d96f9,_0x2aa230,_0x226f44){_0x2aa230=_0x2aa230['headers'];contentType=_0x2aa230["\u0043\u006f\u006e\u0074\u0065\u006e\u0074\u002d\u0054\u0079\u0070\u0065"];contentType2=_0x2aa230['content-type'];var _0x3db254='';if(contentType!=undefined&&contentType!="".split("").reverse().join("")||contentType2!=undefined&&contentType2!="".split("").reverse().join("")){if(contentType=="\u0061\u0070\u0070\u006c\u0069\u0063\u0061\u0074\u0069\u006f\u006e\u002f\u0078\u002d\u0077\u0077\u0077\u002d\u0066\u006f\u0072\u006d\u002d\u0075\u0072\u006c\u0065\u006e\u0063\u006f\u0064\u0065\u0064"){console["\u006c\u006f\u0067"]("\u5F0F\u683C\u5355\u8868 :\u4E3A\u4F53\u6C42\u8BF7\u9001\u53D1\u5230\u6D4B\u68C0 \uDF73\uD83C".split("").reverse().join(""));_0x3db254=dataToFormdata(_0x2d96f9);}else{try{console["\u006c\u006f\u0067"]('🍳\x20检测到发送请求体为:\x20JSON格式');_0x3db254=JSON["\u0073\u0074\u0072\u0069\u006e\u0067\u0069\u0066\u0079"](_0x2d96f9);}catch{console['log']("\u5F0F\u683C\u5355\u8868 :\u4E3A\u4F53\u6C42\u8BF7\u9001\u53D1\u5230\u6D4B\u68C0 \uDF73\uD83C".split("").reverse().join(""));_0x3db254=_0x2d96f9;}}}else{console["\u006c\u006f\u0067"]('🍳\x20检测到发送请求体为:\x20JSON格式');_0x3db254=JSON["\u0073\u0074\u0072\u0069\u006e\u0067\u0069\u0066\u0079"](_0x2d96f9);}if(_0x226f44=='get'||_0x226f44=="\u0047\u0045\u0054"){let _0x335b35=userContent["\u006c\u0065\u006e\u0067\u0074\u0068"]-qlpushFlag;method='get';resp=fetch(_0x1f5035,{'method':method,"\u0068\u0065\u0061\u0064\u0065\u0072\u0073":_0x2aa230})['then'](function(_0x5d662a){return _0x5d662a['text']()['then'](_0x4888be=>{return{'status':_0x5d662a["\u0073\u0074\u0061\u0074\u0075\u0073"],'headers':_0x5d662a["\u0068\u0065\u0061\u0064\u0065\u0072\u0073"],"\u0074\u0065\u0078\u0074":_0x4888be,'response':_0x5d662a,"\u0070\u006f\u0073":_0x335b35};});})['then'](function(_0xb8dfc3){try{_0x2d96f9=JSON["\u0070\u0061\u0072\u0073\u0065"](_0xb8dfc3["\u0074\u0065\u0078\u0074"]);return{'status':_0xb8dfc3["\u0073\u0074\u0061\u0074\u0075\u0073"],"\u0068\u0065\u0061\u0064\u0065\u0072\u0073":_0xb8dfc3["\u0068\u0065\u0061\u0064\u0065\u0072\u0073"],'json':function _0x393c21(){return _0x2d96f9;},'text':function _0x36f803(){return _0xb8dfc3['text'];},"\u0070\u006f\u0073":_0xb8dfc3['pos']};}catch(_0x2baa13){return{'status':_0xb8dfc3['status'],'headers':_0xb8dfc3['headers'],"\u006a\u0073\u006f\u006e":null,'text':function _0x276181(){return _0xb8dfc3['text'];},'pos':_0xb8dfc3['pos']};}})["\u0074\u0068\u0065\u006e"](_0x170faf=>{_0x335b35=_0x170faf["\u0070\u006f\u0073"];flagResultFinish=resultHandle(_0x170faf,_0x335b35);if(flagResultFinish==0x1){i=_0x335b35+0x1;for(;i<=line;i++){var _0xcda32e=Application['Range']('A'+i)['Text'];var _0x163d26=Application['Range']('B'+i)['Text'];if(_0xcda32e=="".split("").reverse().join("")){break;}if(_0x163d26=='是'){console['log']("\uFF1A\u6237\u7528\u884C\u6267\u59CB\u5F00 \uDDD1\uD83E".split("").reverse().join("")+(parseInt(i)-(0x4ef16^0x4ef17)));flagResultFinish=0x0;execHandle(_0xcda32e,i);break;}}}if(_0x335b35==userContent['length']&&flagResultFinish==0x1){flagFinish=0x1;}if(qlpushFlag==(0x82d29^0x82d29)&&flagFinish==0x1){console['log']('🚀\x20青龙发起推送');message=messageMerge();const{sendNotify:_0x153845}=require("\u002e\u002f\u0073\u0065\u006e\u0064\u004e\u006f\u0074\u0069\u0066\u0079\u002e\u006a\u0073");_0x153845(pushHeader,message);qlpushFlag=-0x64;}})['catch'](_0x2e0a9a=>{console['error'](":rorre hcteF".split("").reverse().join(""),_0x2e0a9a);});}else{let _0x29e224=userContent['length']-qlpushFlag;method="tsop".split("").reverse().join("");resp=fetch(_0x1f5035,{'method':method,'headers':_0x2aa230,'body':_0x3db254})['then'](function(_0x5b997d){return _0x5b997d['text']()['then'](_0x3935fa=>{return{'status':_0x5b997d['status'],"\u0068\u0065\u0061\u0064\u0065\u0072\u0073":_0x5b997d['headers'],'text':_0x3935fa,'response':_0x5b997d,'pos':_0x29e224};});})['then'](function(_0x514309){try{_0x2d96f9=JSON["\u0070\u0061\u0072\u0073\u0065"](_0x514309['text']);return{"\u0073\u0074\u0061\u0074\u0075\u0073":_0x514309['status'],'headers':_0x514309['headers'],'json':function _0x1028f0(){return _0x2d96f9;},'text':function _0x9d25d(){return _0x514309['text'];},'pos':_0x514309['pos']};}catch(_0x147532){return{"\u0073\u0074\u0061\u0074\u0075\u0073":_0x514309['status'],"\u0068\u0065\u0061\u0064\u0065\u0072\u0073":_0x514309['headers'],'json':null,'text':function _0x5831a3(){return _0x514309['text'];},'pos':_0x514309["\u0070\u006f\u0073"]};}})['then'](_0x331df2=>{_0x29e224=_0x331df2['pos'];flagResultFinish=resultHandle(_0x331df2,_0x29e224);if(flagResultFinish==0x1){i=_0x29e224+0x1;for(;i<=line;i++){var _0x10c723=Application['Range']('A'+i)['Text'];var _0x32dcc8=Application["\u0052\u0061\u006e\u0067\u0065"]('B'+i)["\u0054\u0065\u0078\u0074"];if(_0x10c723==''){break;}if(_0x32dcc8=='是'){console['log']('🧑\x20开始执行用户：'+(parseInt(i)-0x1));flagResultFinish=0x69161^0x69161;execHandle(_0x10c723,i);break;}}}if(_0x29e224==userContent['length']&&flagResultFinish==0x1){flagFinish=0xbc4a2^0xbc4a3;}if(qlpushFlag==(0xab48a^0xab48a)&&flagFinish==0x1){console['log']("\u9001\u63A8\u8D77\u53D1\u9F99\u9752 \uDE80\uD83D".split("").reverse().join(""));let _0x4e8d15=messageMerge();const{sendNotify:_0x3b25a5}=require('./sendNotify.js');_0x3b25a5(pushHeader,_0x4e8d15);qlpushFlag=-(0x3426d^0x34209);}})['catch'](_0x1368d1=>{console['error'](":rorre hcteF".split("").reverse().join(""),_0x1368d1);});}}};var ApplicationOverwrite={'Range':function Range(_0x778d77){charFirst=_0x778d77['substring'](0x0,0x8132b^0x8132a);qlRow=_0x778d77['substring'](0x1,_0x778d77['length']);qlCol=0x1;for(num in colNum){if(colNum[num]==charFirst){break;}qlCol+=0x1;}try{result=qlSheet[qlRow-(0xc5d1c^0xc5d1d)][qlCol-(0x78bbf^0x78bbe)];}catch{result='';}dict={'Text':result};return dict;},'Sheets':{"\u0049\u0074\u0065\u006d":function(_0xa89467){return{"\u004e\u0061\u006d\u0065":_0xa89467,'Activate':function(){flag=0x1;qlSheet=qlConfig[_0xa89467];if(qlSheet==undefined){qlSheet=qlConfig["\u0053\u0055\u0042\u0043\u004f\u004e\u0046\u0049\u0047"];}console["\u006c\u006f\u0067"]('🍳\x20青龙激活工作表：'+_0xa89467);return flag;}};}}};var CryptoOverwrite={'createHash':function createHash(_0xbffd3c){return{'update':function _0x5c6593(_0x4028c9,_0x51656e){return{'digest':function _0x1c0ac8(_0x3c1125){return{'toUpperCase':function _0x4da1a3(){return{"\u0074\u006f\u0053\u0074\u0072\u0069\u006e\u0067":function _0xa609c(){let _0xd7f56d=require('crypto-js');let _0x5ef11e=_0xd7f56d['MD5'](_0x4028c9)['toString']();_0x5ef11e=_0x5ef11e['toUpperCase']();return _0x5ef11e;}};},'toString':function _0x2cf532(){const _0x3f0262=require('crypto-js');const _0x1158ca=_0x3f0262['MD5'](_0x4028c9)['toString']();return _0x1158ca;}};}};}};}};function dataToFormdata(_0x94b66c){result='';values=Object['values'](_0x94b66c);values['forEach']((_0x4fc62a,_0x665433)=>{key=Object['keys'](_0x94b66c)[_0x665433];content=key+'='+_0x4fc62a+'&';result+=content;});result=result['substring'](0x35095^0x35095,result['length']-(0xf3ca1^0xf3ca0));return result;}function cookiesTocookieMin(_0x5e662f){let _0x497b41=_0x5e662f;let _0x45a30b=[];var _0x355b81=_0x497b41['split']('#');for(let _0x39e20d in _0x355b81){_0x45a30b[_0x39e20d]=_0x355b81[_0x39e20d];}return _0x45a30b;}function checkEscape(_0x5eaa5a,_0x3b5175){cookieArrynew=[];j=0x8c1b2^0x8c1b2;for(i=0xd6308^0xd6308;i<_0x5eaa5a['length'];i++){result=_0x5eaa5a[i];lastChar=result['substring'](result['length']-(0xd697e^0xd697f),result['length']);if(lastChar=='\x5c'&&i<=_0x5eaa5a['length']-0x2){console['log']('🍳\x20检测到转义字符');cookieArrynew[j]=result['substring'](0x86976^0x86976,result['length']-0x1)+_0x3b5175+_0x5eaa5a[parseInt(i)+(0x4c713^0x4c712)];i+=0x1;}else{cookieArrynew[j]=_0x5eaa5a[i];}j+=0x1;}return cookieArrynew;}function cookiesTocookie(_0x507369){let _0x2e90ba=_0x507369;let _0x2f4298=[];let _0x426ccb=[];let _0x1cc909=_0x2e90ba['split']('@');_0x1cc909=checkEscape(_0x1cc909,'@');for(let _0x44b7a2 in _0x1cc909){_0x426ccb=[];let _0x350c7f=Number(_0x44b7a2)+(0x6fbdb^0x6fbda);_0x2f4298=cookiesTocookieMin(_0x1cc909[_0x44b7a2]);_0x2f4298=checkEscape(_0x2f4298,'#');_0x426ccb['push'](_0x2f4298[0xbaa7a^0xbaa7a]);_0x426ccb['push']('是');_0x426ccb['push']("\u79F0\u6635".split("").reverse().join("")+_0x350c7f);if(_0x2f4298["\u006c\u0065\u006e\u0067\u0074\u0068"]>(0x33cd8^0x33cd8)){for(let _0x59b30f=0x3;_0x59b30f<_0x2f4298['length']+(0xdfa43^0xdfa41);_0x59b30f++){_0x426ccb['push'](_0x2f4298[_0x59b30f-0x2]);}}userContent["\u0070\u0075\u0073\u0068"](_0x426ccb);}qlpushFlag=userContent['length']-0x1;}var qlSwitch=0xf0960^0xf0960;try{qlSwitch=process["\u0065\u006e\u0076"][sheetNameSubConfig];qlSwitch=0x1;console['log']('♻️\x20当前环境为青龙');console['log']("\u7801\u4EE3\u9F99\u9752\u884C\u6267\uFF0C\u5883\u73AF\u9F99\u9752\u914D\u9002\u59CB\u5F00 \uFE0F\u267B".split("").reverse().join(""));try{fetch=require('node-fetch');console["\u006c\u006f\u0067"]('♻️\x20系统无fetch，已进行node-fetch引入');}catch{console['log']('♻️\x20系统已有原生fetch');}Crypto=CryptoOverwrite;let flagwarn=0x0;const a='da11990c';const b="0b854f216a9662fb".split("").reverse().join("");encode=getsign(logo);let len=encode['length'];if(a+'ec4dce09'==encode['substring'](0xf3ac2^0xf3ac2,len/0x2)&&b==encode["\u0073\u0075\u0062\u0073\u0074\u0072\u0069\u006e\u0067"](0x4*(0x40675^0x40671),len)){console["\u006c\u006f\u0067"]('✨\x20'+logo);cookies=process['env'][sheetNameSubConfig];}else{console['log']('🔨\x20请使用艾默库代码\x20:\x20https://github.com/imoki/sign_script');flagwarn=0x1;}let flagwarn2=0x2ca7c^0x2ca7d;const welcome="edoc UKOM esu ot emocleW".split("").reverse().join("");const mo=welcome['slice'](0xb9607^0xb9608,0xe8f0f^0xe8f1e)['toLowerCase']();const ku=welcome['split']('\x20')[(0x99b1b^0x99b1f)-0x1]['slice'](0x2,0x4);if(mo['substring'](0x0,0x1)=='m'){if(ku=='KU'){if(mo['substring'](0x1,0x84877^0x84875)==String['fromCharCode'](0x371f7^0x37198)){cookiesTocookie(cookies);flagwarn2=0x0;console['log']('💗\x20'+welcome);}}}let t=Date["\u006e\u006f\u0077"]();if(t>0xaa*0x186a0*0x186a0+0x45f34a08e){console['log']("\u63A5\u94FEnoiton\u5E93\u4ED3\u770B\u67E5\u8BF7\u7A0B\u6559\u7528\u4F7F \uDDFE\uD83E".split("").reverse().join(""));Application=ApplicationOverwrite;}else{flagwarn=0x1;}if(Date["\u006e\u006f\u0077"]()<(0x51d0b^0x51dc3)*0x186a0*0x186a0){console['log']('🤝\x20欢迎各种形式的贡献');HTTP=HTTPOverwrite;}else{flagwarn2=0x1;}if(flagwarn==0x1||flagwarn2==(0x2b39b^0x2b39a)){console['log']('🔨\x20请使用艾默库代码\x20:\x20https://github.com/imoki/sign_script');}}catch{qlSwitch=0x0;console["\u006c\u006f\u0067"]("\u6863\u6587\u5C71\u91D1\u4E3A\u5883\u73AF\u524D\u5F53 \uFE0F\u267B".split("").reverse().join(""));console['log']('♻️\x20开始适配金山文档，执行金山文档代码');}

// =================青龙适配结束===================

// =================金山适配开始===================
// 推送相关
// 获取时间
function getDate(){
  let currentDate = new Date();
  currentDate = currentDate.getFullYear() + '/' + (currentDate.getMonth() + 1).toString() + '/' + currentDate.getDate().toString();
  return currentDate
}

// 将消息写入CONFIG表中作为消息队列，之后统一发送
function writeMessageQueue(message){
  // 当天时间
  let todayDate = getDate()
  flagConfig = ActivateSheet(sheetNameConfig); // 激活主配置表
  // 主配置工作表存在
  if (flagConfig == 1) {
    console.log("✨ 开始将结果写入主配置表");
    for (let i = 2; i <= 100; i++) {
      // 找到指定的表行
      if(Application.Range("A" + (i + 2)).Value == sheetNameSubConfig){
        // 写入更新的时间
        Application.Range("F" + (i + 2)).Value = todayDate
        // 写入消息
        Application.Range("G" + (i + 2)).Value = message
        console.log("✨ 写入结果完成");
        break;
      }
    }
  }

}

// 总推送
function push(message) {
  writeMessageQueue(message)  // 将消息写入CONFIG表中
  // if (message != "") {
  //   // message = messagePushHeader + message // 消息头最前方默认存放：【xxxx】
  //   let length = jsonPush.length;
  //   let name;
  //   let key;
  //   for (let i = 0; i < length; i++) {
  //     if (jsonPush[i].flag == 1) {
  //       name = jsonPush[i].name;
  //       key = jsonPush[i].key;
  //       if (name == "bark") {
  //         bark(message, key);
  //       } else if (name == "pushplus") {
  //         pushplus(message, key);
  //       } else if (name == "ServerChan") {
  //         serverchan(message, key);
  //       } else if (name == "email") {
  //         email(message);
  //       } else if (name == "dingtalk") {
  //         dingtalk(message, key);
  //       } else if (name == "discord") {
  //         discord(message, key);
  //       }
  //     }
  //   }
  // } else {
  //   console.log("🍳 消息为空不推送");
  // }
}


// 推送bark消息
function bark(message, key) {
    if (key != "") {
      message = messagePushHeader + message // 消息头最前方默认存放：【xxxx】
      message = encodeURIComponent(message)
      BARK_ICON = "https://s21.ax1x.com/2024/06/23/pkrUkfe.png"
    let url = "https://api.day.app/" + key + "/" + message + "/" + "?icon=" + BARK_ICON;
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
      message = encodeURIComponent(message)
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
      "?title=" + messagePushHeader +
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
  message = messagePushHeader + message // 消息头最前方默认存放：【xxxx】
  let url = "https://oapi.dingtalk.com/robot/send?access_token=" + key;
  let resp = HTTP.post(url, { msgtype: "text", text: { content: message } });
  // console.log(resp.text())
  sleep(5000);
}

// 推送Discord机器人
function discord(message, key) {
  message = messagePushHeader + message // 消息头最前方默认存放：【xxxx】
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
      message += "\n" + messageHeader[i] + messageArray[i] + ""; // 加上推送头
    }
  }
  if(message != "")
  {
    console.log("✨✨✨✨✨✨✨✨✨✨✨✨✨✨✨✨✨✨")
    console.log(message + "\n")  // 打印总消息
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
    messageHeader[posLabel] =  "👨‍🚀 " + messageName

    data = {
      "sign_cyclic":"True",
    }

    // console.log(resp.status)
    // console.log(resp.json())
    // if (resp.status == 200) {
        if(posHttp == 1 || qlSwitch != 1){  // 只在第一次用, 或者执行金山文档
            resp = resp.json();
            // console.log(resp)
            try{
                isSign = resp["data"]["cap_sign"]["sign_daily"]
            }catch{
                content = "⛔ " + "账号可能未登录，请重新登录 "
                // messageFail += content
                // messageSuccess += "帐号：" + messageName + "已经签到过了,奖励容量"  + String(number) + "MB";
                console.log(content)
                
                // // 青龙适配，青龙微适配
                // flagResultFinish = 1; // 签到结束

            }

            // isSign = ~true // 测试
        }else{
            // {
            //     status: 500,
            //     code: 15000,
            //     message: 'inner error, requestId ',
            //     req_id: '',
            //     timestamp: 1718506995
            // }
            isSign = ~true   // 第二次以上进来默认通过
        }
      // isSign = true
      // console.log(isSign)
      if(isSign == true)
      {
        console.log("📢 " + "已经签到过了")
        reward = resp["data"]["cap_sign"]["sign_daily_reward"] / (1024 * 1024)
        cur_total_sign_day = resp["data"]["cap_growth"]["cur_total_sign_day"] // 总签到天数
        sign_progress = resp["data"]["cap_sign"]["sign_progress"] // 当周签到天数
        
        // console.log(reward)
        // content = "帐号：" + messageName + "已经签到过了,奖励"  + String(number) + "MB" + ",总签到" + cur_total_sign_day + "天 " + ",当周已签" + sign_progress + "天 ";
        content = "📢 " + "总签" + cur_total_sign_day + "天" + ",周签" + sign_progress + "天,获"  + String(reward) + "MB ";
        messageSuccess += content
        // messageSuccess += "帐号：" + messageName + "已经签到过了,奖励容量"  + String(number) + "MB";
        // console.log(content)
        
        // 青龙适配，青龙微适配
        flagResultFinish = 1; // 签到结束
      }else
      {
        if(posHttp == 1 || qlSwitch != 1){  // 第一次进来时用
            console.log("🍳 进行签到")
            // {"status":200,"code":0,"message":"","timestamp":170000000,"data":{"sign_daily_reward":20000000},"metadata":{}}
            // {"status":400,"code":44210,"message":"cap_growth_sign_repeat","req_id":"xxxzzz-xxxxxxx","timestamp":17000000}
            let url2 = "https://drive-m.quark.cn/1/clouddrive/capacity/growth/sign?pr=ucpro&fr=pc&uc_param_str="; // 进行签到
            resp = HTTP.post(
                url2,
                JSON.stringify(data),
                { headers: headers }
            );
        }

        // console.log(resp.json())


        // {"status":200,"code":0,"message":"","timestamp":170000000,"data":{"member_type":"NORMAL","use_capacity":120000000,"cap_growth":{"lost_total_cap":0,"cur_total_cap":11000000,"cur_total_sign_day":46},"88VIP":false,"member_status":{"Z_VIP":"UNPAID","VIP":"UNPAID","SUPER_VIP":"UNPAID","MINI_VIP":"UNPAID"},"cap_sign":{"sign_daily":true,"sign_target":7,"sign_daily_reward":2000000,"sign_progress":4,"sign_rewards":[{"name":"+20MB","reward_cap":2000000},{"name":"+40MB","highlight":"翻倍","reward_cap":4000000},{"name":"+20MB","reward_cap":2000000},{"name":"+20MB","reward_cap":200000},{"name":"+20MB","reward_cap":2000000},{"name":"+20MB","reward_cap":2000000},{"name":"+100MB","highlight":"翻五倍","reward_cap":10000000}]},"cap_composition":{"other":21000000,"member_own":100000000,"sign_reward":10000000},"total_capacity":1400000000},"metadata":{}}
        // if (resp.status == 200) {

        if(posHttp == 2 || qlSwitch != 1){  // 第二次进来时用
            // resp = resp.json();
            // console.log(resp)
            // resp = {"status":200,"code":0,"message":"","timestamp":170000000,"data":{"sign_daily_reward":20971520},"metadata":{}}
            // 41943040 -> 40MB
            // reward = resp["data"]["sign_daily_reward"] / (1024 * 1024)
            // console.log(reward)


            // 查询签到天数
            // resp = HTTP.fetch(url1, {
            //     method: "get",
            //     headers: headers,
            //     // data: data
            // });
            url =  "https://drive-m.quark.cn/1/clouddrive/capacity/growth/info?pr=ucpro&fr=pc&uc_param_str=";
            if(qlSwitch != 1){  // 金山文档
                resp = HTTP.fetch(url, {
                    method: "get",
                    headers: headers,
                    // data: data
                });
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

        } 
        

        if(posHttp == 3 || qlSwitch != 1){  // 第三次进来时用
            resp = resp.json();
            // console.log(resp)
            // {
            //     status: 200,
            //     code: 0,
            //     message: '',
            //     timestamp: ,
            //     data: {
            //         member_type: 'NORMAL',
            //         use_capacity: ,
            //         cap_growth: {
            //         lost_total_cap: 0,
            //         cur_total_cap: ,
            //         cur_total_sign_day: 85
            //         },
            //         '88VIP': false,
            //         member_status: {
            //         Z_VIP: 'UNPAID',
            //         VIP: 'UNPAID',
            //         MINI_VIP: 'UNPAID',
            //         SUPER_VIP: 'UNPAID'
            //         },
            //         cap_sign: {
            //         sign_daily: true,
            //         sign_target: 7,
            //         sign_daily_reward: 20971520,
            //         sign_progress: 1,
            //         sign_rewards: [Array]
            //         },
            //         cap_composition: {
            //         other: ,
            //         member_own: ,
            //         sign_reward: 2600468480
            //         },
            //         total_capacity: 
            //     },
            //     metadata: {}
            // }

            try{
                // 41943040 -> 40MB
                reward = resp["data"]["cap_sign"]["sign_daily_reward"] / (1024 * 1024)
                // console.log(reward)
                cur_total_sign_day = resp["data"]["cap_growth"]["cur_total_sign_day"] // 总签到天数
                sign_progress = resp["data"]["cap_sign"]["sign_progress"] // 当周签到天数
                content = "🎉 " + "总签" + cur_total_sign_day + "天" + ",周签" + sign_progress + "天,获"  + String(reward) + "MB ";
                messageSuccess += content
                // console.log(content)
            }catch{
                content = "❌ " + "账号可能未登录，请重新登录 "
                messageFail += content
                // messageSuccess += "帐号：" + messageName + "已经签到过了,奖励容量"  + String(number) + "MB";
                // console.log(content)
                
                // 青龙适配，青龙微适配
                flagResultFinish = 1; // 签到结束

            }

            // 青龙适配，青龙微适配
            flagResultFinish = 1; // 签到结束
        }
            
        //   }
      }
      
    // } else {
    // //   console.log(resp.text());
    //   messageFail += "帐号：" + messageName + "签到失败 ";
    //   console.log("帐号：" + messageName + "签到失败 ");
    // }


    if (messageOnlyError == 1) {
        messageArray[posLabel] = messageFail;
    } else {
        if(messageFail != ""){
            messageArray[posLabel] = messageFail + " " + messageSuccess;
        }else{
            messageArray[posLabel] = messageSuccess;
        }
    }

    if(messageArray[posLabel] != "")
    {
        console.log(messageArray[posLabel]);
    }

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
    let url1 = "https://drive-m.quark.cn/1/clouddrive/capacity/growth/info?pr=ucpro&fr=pc&uc_param_str="; // 查询是否签到
    
    headers = {
      "Cookie": cookie,
      "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/46.0.2486.0 Safari/537.36 Edge/13.10586"
    };

    if(qlSwitch != 1){  // 金山文档
        resp = HTTP.fetch(url1, {
            method: "get",
            headers: headers,
            // data: data
        });
    }else{  // 青龙
        data = {}
        option = "get"
        resp = HTTP.post(
            url1,
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
