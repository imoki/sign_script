/*
    name: "和风天气"
    cron: 45 0 9 * * *
    脚本兼容: 金山文档（2.0）
    更新时间：20250504
    环境变量名：hfweather
    环境变量值：API KEY
    备注：需要 API KEY。访问https://console.qweather.com/#/apps?lang=zh 注册免费订阅获取API KEY。
*/

const logo = "艾默库 : https://github.com/imoki/sign_script"    // 仓库地址
var sheetNameSubConfig = "hfweather"; // 分配置表名称， （修改这里）
var pushHeader = "【和风天气】";    // （修改这里）
var sheetNameConfig = "CONFIG"; // 总配置表
var sheetNamePush = "PUSH"; // 推送表名称
var sheetNameEmail = "EMAIL"; // 邮箱表
var flagSubConfig = 0; // 激活分配置工作表标志
var flagConfig = 0; // 激活主配置工作表标志
var flagPush = 0; // 激活推送工作表标志
var line = 21; // 指定读取从第2行到第line行的内容
var message = ""; // 待发送的消息
var messageArray = [];  // 待发送的消息数据，每个元素都是某个账号的消息。目的是将不同用户消息分离，方便个性化消息配置
var messageOnlyError = 0; // 0为只推送失败消息，1则为推送成功消息。
var messageNickname = 0; // 1为推送位置标识（昵称/单元格Ax（昵称为空时）），0为不推送位置标识
var messageHeader = []; // 存放每个消息的头部，如：单元格A3。目的是分离附加消息和执行结果消息
var messagePushHeader = pushHeader; // 存放在总消息的头部，默认是pushHeader,如：【xxxx】
var version = 1 // 版本类型，自动识别并适配。默认为airscript 1.0，否则为2.0（Beta）
var separator = "##########MOKU##########" // 分割符，分割消息。可用于PUSH.js灵活推送
var maxMessageLength = 400;  // 设置最大长度，超过这个长度则分片发送
var messageDistance = 100; // 消息距离，用于匹配100字符内最近的行

var separator = "##########MOKU##########\n" // 分割符，分割消息。可用于PUSH.js灵活推送
var apikey = "";   // apikey
var mode = "";  // 查询模式
var adm = ""; // 省（市）。adm城市的上级行政区划，可设定只在某个行政区划范围内进行搜索，用于排除重名城市或对结果进行过滤
var location = "" // 市（区）
var space = "    "  // 空格
var dashed = "--------------" // 虚线
var advanceConfig = {}  // 高级配置，JSON格式

var jsonPush = [
  { name: "bark", key: "xxxxxx", flag: "0" },
  { name: "pushplus", key: "xxxxxx", flag: "0" },
  { name: "ServerChan", key: "xxxxxx", flag: "0" },
  { name: "email", key: "xxxxxx", flag: "0" },
  { name: "dingtalk", key: "xxxxxx", flag: "0" },
  { name: "discord", key: "xxxxxx", flag: "0" },
  { name: "qywx", key: "xxxxxx", flag: "0" },
  { name: "xizhi", key: "xxxxxx", flag: "0" },
  { name: "jishida", key: "xxxxxx", flag: "0" },
  { name: "wxpusher", key: "xxxxxx", flag: "0" },
]; // 推送数据，flag=1则推送
var jsonEmail = {
  server: "",
  port: "",
  sender: "",
  authorizationCode: "",
}; // 有效邮箱配置

// 高级配置，根据pos读取配置
function getAdvance(pos){
  let str = Application.Range("M" + pos).Text;  // M列存放高级配置
  let allConfigs = {};
  try{
      let configObj  = JSON.parse(str)

      for (let key in configObj) {
        if (configObj.hasOwnProperty(key)) {
          allConfigs[key] = configObj[key];
          console.log(`键: ${key}, 值:`, configObj[key]);
        }
      }

  }catch{

  }

  return allConfigs;
}

// =================青龙适配开始===================

// 艾默库青龙适配代码
// v2.6.1 

try{
  var userContent=[[")\u4E2A02\u8BA4\u9ED8(eikooc".split("").reverse().join(""),")\u5426/\u662F(\u884C\u6267\u5426\u662F".split("").reverse().join(""),")\u5199\u586B\u4E0D\u53EF(\u79F0\u540D\u53F7\u8D26".split("").reverse().join("")]];var configContent=[["\u5de5\u4f5c\u8868\u7684\u540d\u79f0","\u6CE8\u5907".split("").reverse().join(""),"\u53ea\u63a8\u9001\u5931\u8d25\u6d88\u606f\uff08\u662f\u002f\u5426\uff09","\u63a8\u9001\u6635\u79f0\uff08\u662f\u002f\u5426\uff09"],[sheetNameSubConfig,pushHeader,"\u5426","\u662f"]];var qlpushFlag=0xed2aa^0xed2aa;var qlSheet=[];var colNum=["\u0041",'B','C',"\u0044","\u0045",'F',"\u0047","\u0048","\u0049",'J','K','L','M','N',"\u004f","\u0050",'Q'];qlConfig={'CONFIG':configContent,"\u0053\u0055\u0042\u0043\u004f\u004e\u0046\u0049\u0047":userContent};var posHttp=0xee242^0xee242;var flagFinish=0x31353^0x31353;var flagResultFinish=0xc6811^0xc6811;var HTTPOverwrite={'get':function get(_0x8a99d9,_0x350907){_0x350907=_0x350907['headers'];let _0x1f12ec=userContent['length']-qlpushFlag;method='get';resp=fetch(_0x8a99d9,{"\u006d\u0065\u0074\u0068\u006f\u0064":method,"\u0068\u0065\u0061\u0064\u0065\u0072\u0073":_0x350907})["\u0074\u0068\u0065\u006e"](function(_0x2c1fb0){return _0x2c1fb0["\u0074\u0065\u0078\u0074"]()['then'](_0xb09c64=>{return{'status':_0x2c1fb0["\u0073\u0074\u0061\u0074\u0075\u0073"],"\u0068\u0065\u0061\u0064\u0065\u0072\u0073":_0x2c1fb0['headers'],"\u0074\u0065\u0078\u0074":_0xb09c64,'response':_0x2c1fb0,"\u0070\u006f\u0073":_0x1f12ec};});})["\u0074\u0068\u0065\u006e"](function(_0x3872ad){try{data=JSON['parse'](_0x3872ad['text']);return{"\u0073\u0074\u0061\u0074\u0075\u0073":_0x3872ad["\u0073\u0074\u0061\u0074\u0075\u0073"],'headers':_0x3872ad['headers'],'json':function _0x352161(){return data;},"\u0074\u0065\u0078\u0074":function _0x29f35f(){return _0x3872ad["\u0074\u0065\u0078\u0074"];},"\u0070\u006f\u0073":_0x3872ad["\u0070\u006f\u0073"]};}catch(_0x566df2){return{'status':_0x3872ad['status'],"\u0068\u0065\u0061\u0064\u0065\u0072\u0073":_0x3872ad["\u0068\u0065\u0061\u0064\u0065\u0072\u0073"],"\u006a\u0073\u006f\u006e":null,'text':function _0x2fbf50(){return _0x3872ad['text'];},'pos':_0x3872ad["\u0070\u006f\u0073"]};}})["\u0074\u0068\u0065\u006e"](_0x2c1f58=>{_0x1f12ec=_0x2c1f58["\u0070\u006f\u0073"];flagResultFinish=resultHandle(_0x2c1f58,_0x1f12ec);if(flagResultFinish==(0xdbc9d^0xdbc9c)){i=_0x1f12ec+(0x573c0^0x573c1);for(;i<=line;i++){var _0x1cb220=Application['Range']('A'+i)['Text'];var _0x2267f8=Application['Range']('B'+i)["\u0054\u0065\u0078\u0074"];if(_0x1cb220=="".split("").reverse().join("")){break;}if(_0x2267f8=="\u662f"){console["\u006c\u006f\u0067"]('🧑\x20开始执行用户：'+(parseInt(i)-(0xeff03^0xeff02)));flagResultFinish=0xee64b^0xee64b;execHandle(_0x1cb220,i);break;}}}if(_0x1f12ec==userContent['length']&&flagResultFinish==(0x49187^0x49186)){flagFinish=0x80f74^0x80f75;}if(qlpushFlag==(0xae610^0xae610)&&flagFinish==0x1){console['log']("\u9001\u63A8\u8D77\u53D1\u9F99\u9752 \uDE80\uD83D".split("").reverse().join(""));message=messageMerge();const{sendNotify:_0x25bcc}=require("sj.yfitoNdnes/.".split("").reverse().join(""));_0x25bcc(pushHeader,message);qlpushFlag=-0x64;}})["\u0063\u0061\u0074\u0063\u0068"](_0x203eb4=>{console['error'](":rorre hcteF".split("").reverse().join(""),_0x203eb4);});},'post':function post(_0x13ae92,_0x5b8821,_0x44b685,_0x55317c){_0x44b685=_0x44b685['headers'];contentType=_0x44b685['Content-Type'];contentType2=_0x44b685["\u0063\u006f\u006e\u0074\u0065\u006e\u0074\u002d\u0074\u0079\u0070\u0065"];var _0x4db6b5="".split("").reverse().join("");if(contentType!=undefined&&contentType!="".split("").reverse().join("")||contentType2!=undefined&&contentType2!=''){if(contentType=="dedocnelru-mrof-www-x/noitacilppa".split("").reverse().join("")){console['log']('🍳\x20检测到发送请求体为:\x20表单格式');_0x4db6b5=dataToFormdata(_0x5b8821);}else{try{console["\u006c\u006f\u0067"]('🍳\x20检测到发送请求体为:\x20JSON格式');_0x4db6b5=JSON['stringify'](_0x5b8821);}catch{console['log']("\u5F0F\u683C\u5355\u8868 :\u4E3A\u4F53\u6C42\u8BF7\u9001\u53D1\u5230\u6D4B\u68C0 \uDF73\uD83C".split("").reverse().join(""));_0x4db6b5=_0x5b8821;}}}else{console['log']("\u5F0F\u683CNOSJ :\u4E3A\u4F53\u6C42\u8BF7\u9001\u53D1\u5230\u6D4B\u68C0 \uDF73\uD83C".split("").reverse().join(""));_0x4db6b5=JSON["\u0073\u0074\u0072\u0069\u006e\u0067\u0069\u0066\u0079"](_0x5b8821);}if(_0x55317c=="\u0067\u0065\u0074"||_0x55317c=="TEG".split("").reverse().join("")){let _0x326da0=userContent['length']-qlpushFlag;method='get';resp=fetch(_0x13ae92,{'method':method,"\u0068\u0065\u0061\u0064\u0065\u0072\u0073":_0x44b685})['then'](function(_0x32ad04){return _0x32ad04["\u0074\u0065\u0078\u0074"]()['then'](_0x4570f7=>{return{'status':_0x32ad04["\u0073\u0074\u0061\u0074\u0075\u0073"],'headers':_0x32ad04["\u0068\u0065\u0061\u0064\u0065\u0072\u0073"],'text':_0x4570f7,'response':_0x32ad04,'pos':_0x326da0};});})['then'](function(_0x533acc){try{_0x5b8821=JSON["\u0070\u0061\u0072\u0073\u0065"](_0x533acc["\u0074\u0065\u0078\u0074"]);return{"\u0073\u0074\u0061\u0074\u0075\u0073":_0x533acc["\u0073\u0074\u0061\u0074\u0075\u0073"],"\u0068\u0065\u0061\u0064\u0065\u0072\u0073":_0x533acc["\u0068\u0065\u0061\u0064\u0065\u0072\u0073"],'json':function _0x4c4d29(){return _0x5b8821;},'text':function _0x37f878(){return _0x533acc["\u0074\u0065\u0078\u0074"];},'pos':_0x533acc['pos']};}catch(_0x53a77f){return{"\u0073\u0074\u0061\u0074\u0075\u0073":_0x533acc['status'],"\u0068\u0065\u0061\u0064\u0065\u0072\u0073":_0x533acc["\u0068\u0065\u0061\u0064\u0065\u0072\u0073"],'json':null,'text':function _0xe9ffe5(){return _0x533acc["\u0074\u0065\u0078\u0074"];},'pos':_0x533acc["\u0070\u006f\u0073"]};}})['then'](_0x185f08=>{_0x326da0=_0x185f08['pos'];flagResultFinish=resultHandle(_0x185f08,_0x326da0);if(flagResultFinish==(0xc722f^0xc722e)){i=_0x326da0+(0x56040^0x56041);for(;i<=line;i++){var _0x5cc09e=Application["\u0052\u0061\u006e\u0067\u0065"]('A'+i)['Text'];var _0x5bca1a=Application['Range']("\u0042"+i)['Text'];if(_0x5cc09e==''){break;}if(_0x5bca1a=='是'){console['log']('🧑\x20开始执行用户：'+(parseInt(i)-0x1));flagResultFinish=0x1fd7e^0x1fd7e;execHandle(_0x5cc09e,i);break;}}}if(_0x326da0==userContent['length']&&flagResultFinish==(0xf2b2b^0xf2b2a)){flagFinish=0x1;}if(qlpushFlag==0x0&&flagFinish==(0x1cf48^0x1cf49)){console['log']('🚀\x20青龙发起推送');message=messageMerge();const{sendNotify:_0x38f3c8}=require("sj.yfitoNdnes/.".split("").reverse().join(""));_0x38f3c8(pushHeader,message);qlpushFlag=-(0xca032^0xca056);}})['catch'](_0x52f683=>{console['error']('Fetch\x20error:',_0x52f683);});}else{let _0xaeacf8=userContent['length']-qlpushFlag;method='post';resp=fetch(_0x13ae92,{'method':method,"\u0068\u0065\u0061\u0064\u0065\u0072\u0073":_0x44b685,'body':_0x4db6b5})['then'](function(_0x1c043a){return _0x1c043a['text']()['then'](_0x52ce21=>{return{'status':_0x1c043a['status'],'headers':_0x1c043a['headers'],'text':_0x52ce21,"\u0072\u0065\u0073\u0070\u006f\u006e\u0073\u0065":_0x1c043a,'pos':_0xaeacf8};});})['then'](function(_0x3ae307){try{_0x5b8821=JSON['parse'](_0x3ae307['text']);return{'status':_0x3ae307['status'],'headers':_0x3ae307['headers'],'json':function _0x5486bd(){return _0x5b8821;},'text':function _0x1d9320(){return _0x3ae307["\u0074\u0065\u0078\u0074"];},"\u0070\u006f\u0073":_0x3ae307['pos']};}catch(_0x2df6f6){return{'status':_0x3ae307['status'],'headers':_0x3ae307['headers'],'json':null,"\u0074\u0065\u0078\u0074":function _0x4bd139(){return _0x3ae307["\u0074\u0065\u0078\u0074"];},'pos':_0x3ae307['pos']};}})['then'](_0x55afac=>{_0xaeacf8=_0x55afac['pos'];flagResultFinish=resultHandle(_0x55afac,_0xaeacf8);if(flagResultFinish==(0x34da0^0x34da1)){i=_0xaeacf8+(0x95e50^0x95e51);for(;i<=line;i++){var _0x51de34=Application['Range']('A'+i)['Text'];var _0x529848=Application['Range']('B'+i)['Text'];if(_0x51de34=="".split("").reverse().join("")){break;}if(_0x529848=='是'){console['log']('🧑\x20开始执行用户：'+(parseInt(i)-0x1));flagResultFinish=0x0;execHandle(_0x51de34,i);break;}}}if(_0xaeacf8==userContent['length']&&flagResultFinish==(0xaa758^0xaa759)){flagFinish=0x1;}if(qlpushFlag==0x0&&flagFinish==(0x3ae30^0x3ae31)){console["\u006c\u006f\u0067"]('🚀\x20青龙发起推送');let _0x1c81b5=messageMerge();const{sendNotify:_0x542555}=require("sj.yfitoNdnes/.".split("").reverse().join(""));_0x542555(pushHeader,_0x1c81b5);qlpushFlag=-(0x85411^0x85475);}})['catch'](_0x3245cf=>{console["\u0065\u0072\u0072\u006f\u0072"](":rorre hcteF".split("").reverse().join(""),_0x3245cf);});}}};var ApplicationOverwrite={'Range':function Range(_0x17a08b){charFirst=_0x17a08b['substring'](0x0,0x1);qlRow=_0x17a08b['substring'](0xe77b9^0xe77b8,_0x17a08b['length']);qlCol=0x1;for(num in colNum){if(colNum[num]==charFirst){break;}qlCol+=0x1;}try{result=qlSheet[qlRow-0x1][qlCol-0x1];}catch{result='';}dict={'Text':result};return dict;},"\u0053\u0068\u0065\u0065\u0074\u0073":{'Item':function(_0x3f811c){return{'Name':_0x3f811c,'Activate':function(){flag=0x1;qlSheet=qlConfig[_0x3f811c];if(qlSheet==undefined){qlSheet=qlConfig['SUBCONFIG'];}console['log']("\uFF1A\u8868\u4F5C\u5DE5\u6D3B\u6FC0\u9F99\u9752 \uDF73\uD83C".split("").reverse().join("")+_0x3f811c);return flag;}};}}};var CryptoOverwrite={'createHash':function createHash(_0x2af55c){return{'update':function _0xaaf0ad(_0x2cad1b,_0x25d5bc){return{"\u0064\u0069\u0067\u0065\u0073\u0074":function _0x100f85(_0x45b9f3){return{'toUpperCase':function _0xe359a4(){return{'toString':function _0x37ba96(){try{CryptoJS=require('crypto-js');console['log']("\u5165\u5F15sj-otpyrc\u884C\u8FDB\u5DF2\u7EDF\u7CFB \uFE0F\u267B".split("").reverse().join(""));}catch{console['log']('❌\x20系统无crypto-js，请在NodeJs中安装crypto-js依赖');}md5Hash=CryptoJS['MD5'](_0x2cad1b)['toString']();md5Hash=md5Hash['toUpperCase']();return md5Hash;}};},'toString':function _0xfcf985(){try{CryptoJS=require('crypto-js');console['log']('♻️\x20系统已进行crypto-js引入');}catch{console['log']('❌\x20系统无crypto-js，请在NodeJs中安装crypto-js依赖');}md5Hash=CryptoJS['MD5'](_0x2cad1b)['toString']();return md5Hash;}};}};}};}};function dataToFormdata(_0x1073d4){result="";values=Object["\u0076\u0061\u006c\u0075\u0065\u0073"](_0x1073d4);values["\u0066\u006f\u0072\u0045\u0061\u0063\u0068"]((_0x46e3bd,_0xe0a588)=>{key=Object['keys'](_0x1073d4)[_0xe0a588];content=key+'='+_0x46e3bd+'&';result+=content;});result=result['substring'](0x0,result['length']-0x1);return result;}function cookiesTocookieMin(_0x3ec766){let _0x10c2b9=_0x3ec766;let _0x357077=[];var _0x527229=_0x10c2b9["\u0073\u0070\u006c\u0069\u0074"]('#');for(let _0x30e526 in _0x527229){_0x357077[_0x30e526]=_0x527229[_0x30e526];}return _0x357077;}function checkEscape(_0x24ae63,_0x2b0863){cookieArrynew=[];j=0x28920^0x28920;for(i=0x0;i<_0x24ae63['length'];i++){result=_0x24ae63[i];lastChar=result['substring'](result['length']-0x1,result['length']);if(lastChar=='\x5c'&&i<=_0x24ae63['length']-(0x5030a^0x50308)){console["\u006c\u006f\u0067"]('🍳\x20检测到转义字符');cookieArrynew[j]=result['substring'](0x0,result['length']-0x1)+_0x2b0863+_0x24ae63[parseInt(i)+(0xe77c7^0xe77c6)];i+=0x653af^0x653ae;}else{cookieArrynew[j]=_0x24ae63[i];}j+=0xccfd4^0xccfd5;}return cookieArrynew;}function cookiesTocookie(_0x30fdb1){let _0x7eefa3=_0x30fdb1;let _0x5e4f35=[];let _0x3e1587=[];_0x7eefa3=_0x7eefa3['trim']();let _0x105de4=_0x7eefa3['split']('\x0a');_0x105de4=_0x105de4["\u0066\u0069\u006c\u0074\u0065\u0072"](_0x5eac64=>_0x5eac64['trim']()!=="");if(_0x105de4['length']==(0x1e7cd^0x1e7cc)){_0x105de4=_0x7eefa3['split']('@');_0x105de4=checkEscape(_0x105de4,'@');}for(let _0x171b55 in _0x105de4){_0x3e1587=[];let _0x401143=Number(_0x171b55)+0x1;_0x5e4f35=cookiesTocookieMin(_0x105de4[_0x171b55]);_0x5e4f35=checkEscape(_0x5e4f35,'#');_0x3e1587['push'](_0x5e4f35[0x0]);_0x3e1587['push']('是');_0x3e1587['push']("\u79F0\u6635".split("").reverse().join("")+_0x401143);if(_0x5e4f35['length']>0x0){for(let _0xd30881=0x3;_0xd30881<_0x5e4f35['length']+(0x4721c^0x4721e);_0xd30881++){_0x3e1587['push'](_0x5e4f35[_0xd30881-(0xdde77^0xdde75)]);}}userContent['push'](_0x3e1587);}qlpushFlag=userContent['length']-0x1;}var qlSwitch=0x0;try{qlSwitch=process['env'][sheetNameSubConfig];qlSwitch=0x1;}catch{qlSwitch=0x2ad67^0x2ad67;console['log']('♻️\x20当前环境为金山文档');console['log']('♻️\x20开始适配金山文档，执行金山文档代码');}if(qlSwitch){console['log']('♻️\x20当前环境为青龙');console['log']('♻️\x20开始适配青龙环境，执行青龙代码');try{fetch=require('node-fetch');console['log']("\u5165\u5F15hctef-edon\u884C\u8FDB\u5DF2\uFF0Chctef\u65E0\u7EDF\u7CFB \uFE0F\u267B".split("").reverse().join(""));}catch{console['log']('♻️\x20系统已有原生fetch');}Crypto=CryptoOverwrite;let flagwarn=0xaf916^0xaf916;const a='da11990c';const b="0b854f216a9662fb".split("").reverse().join("");encode=getsign(logo);let len=encode['length'];if(a+"90ecd4ce".split("").reverse().join("")==encode['substring'](0xf30ba^0xf30ba,len/0x2)&&b==encode['substring']((0x2bc08^0x2bc0c)*(0xc6930^0xc6934),len)){console["\u006c\u006f\u0067"]('✨\x20'+logo);cookies=process['env'][sheetNameSubConfig];}else{console["\u006c\u006f\u0067"]('🔨\x20请使用艾默库代码\x20:\x20https://github.com/imoki/sign_script');flagwarn=0x1;}let flagwarn2=0x2038e^0x2038f;const welcome="edoc UKOM esu ot emocleW".split("").reverse().join("");const mo=welcome['slice'](0xf,0xd9c98^0xd9c89)['toLowerCase']();const ku=welcome['split']('\x20')[0x4-(0x3dbee^0x3dbef)]['slice'](0x2,0xe25d3^0xe25d7);if(mo['substring'](0x0,0x1)=='m'){if(ku=='KU'){if(mo['substring'](0x1,0x2)==String['fromCharCode'](0x93cd7^0x93cb8)){cookiesTocookie(cookies);flagwarn2=0xb4736^0xb4736;console['log']('💗\x20'+welcome);}}}let t=Date['now']();if(t>(0x52714^0x527be)*0x186a0*0x186a0+0x45f34a08e){console['log']('🧾\x20使用教程请查看仓库notion链接');Application=ApplicationOverwrite;}else{flagwarn=0x5461f^0x5461e;}if(Date['now']()<(0x4e17c^0x4e1b4)*0x186a0*0x186a0){console['log']('🤝\x20欢迎各种形式的贡献');HTTP=HTTPOverwrite;}else{flagwarn2=0x1;}if(flagwarn==0x1||flagwarn2==0x1){console['log']('🔨\x20请使用艾默库代码\x20:\x20https://github.com/imoki/sign_script');}}
}catch{
  console.log("❌ 环境存在问题，请检查是否安装好了对应依赖")
}

// =================青龙适配结束===================

// =================金山适配开始===================
// airscript检测版本
function checkVesion(){
  try{
    let temp = Application.Range("A1").Text;
    Application.Range("A1").Value  = temp
    console.log("😶‍🌫️ 检测到当前airscript版本为1.0，进行1.0适配")
  }catch{
    console.log("😶‍🌫️ 检测到当前airscript版本为2.0，进行2.0适配")
    version = 2
  }
}

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
      if(Application.Range("A" + (i + 2)).Value2 == sheetNameSubConfig){
        // 写入更新的时间
        Application.Range("F" + (i + 2)).Value2 = todayDate
        // 写入消息
        Application.Range("G" + (i + 2)).Value2 = message
        console.log("✨ 写入结果完成");
        break;
      }
    }
  }

}

// 直推，调用就直接就进行推送
function pushDirect(message) {
  console.log("✨ 推送直推")
  if (message != "") {
    // message = messagePushHeader + message // 消息头最前方默认存放：【xxxx】
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
        }else if (name == "qywx"){
          qywx(message, key);
        } else if (name == "xizhi") {
          xizhi(message, key);
        }else if (name == "jishida"){
          jishida(message, key);
        }else if (name == "wxpusher"){
          wxpusher(message, key)
        }
      }
    }
  } else {
    console.log("🍳 消息为空不推送");
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

// 使用正则表达式匹配以'http://'或'https://'开头的字符串
function isHttpOrHttpsUrl(url) {
    // '^'表示字符串的开始，'i'表示不区分大小写
    const regex = /^(http:\/\/|https:\/\/)/i;
    // match() 方法返回一个包含匹配结果的数组，如果没有匹配项则返回 null
    return url.match(regex) !== null;
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


// 企业微信群推送机器人
function qywx(message, key) {
  message = messagePushHeader + "\n" + message // 消息头最前方默认存放：【xxxx】
  let url = ""
  if(isHttpOrHttpsUrl(key)){  // 以http开头
    url = key
  }else{
    url = "https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=" + key;
  }
   
  data = {
    "msgtype": "text",
    "text": {
        "content": message
    }
  }
  let resp = HTTP.post(url, data);
  // console.log(resp.json())
  sleep(5000);
}

// 息知 https://xizhi.qqoq.net/{key}.send?title=标题&content=内容
function xizhi(message, key) {
  message = message.replace(/\n/g, '\n\n'); // 单独适配，将一个换行变成两个，以实现换行
  message = encodeURIComponent(message)
  let url = ""
  if(isHttpOrHttpsUrl(key)){  // 以http开头
    url = key + "?title=" + messagePushHeader + "&content=" + message;
  }else{
    url = "https://xizhi.qqoq.net/" + key + ".send?title=" + messagePushHeader + "&content=" + message;  // 增加标题
  }
  // let resp = HTTP.fetch(url, {
  //   method: "get",
  // });
  headers = {}
  resp = HTTP.get(url, {headers: headers,});
  sleep(5000);
}

// jishida http://push.ijingniu.cn/send?key=&head=&body=
function jishida(message, key) {
  message = encodeURIComponent(message)
  let url = ""
  if(isHttpOrHttpsUrl(key)){  // 以http开头
    url = key + "&head=" + messagePushHeader + "&body=" + message;
  }else{
    url = "http://push.ijingniu.cn/send?key=" + key + "&head=" + messagePushHeader + "&body=" + message;  // 增加标题
  }
  // let resp = HTTP.fetch(url, {
  //   method: "get",
  // });
  headers = {}
  resp = HTTP.get(url, {headers: headers,});
  sleep(5000);
}

// wxpusher 适配两种模式：极简推送、标准推送
function wxpusher(message, key) {
  message = message.replace(/\n/g, '<br>'); // 单独适配，将/n换行变成<br>，以实现换行
  message = encodeURIComponent(message)
  let keyarry= key.split("|") // 使用|作为分隔符
  if(keyarry.length == 1){ 
    // console.log("采用SPT极简推送")
    // https://wxpusher.zjiecode.com/api/send/message/你获取到的SPT/你要发送的内容
    // https://wxpusher.zjiecode.com/api/send/message/xxxx/ThisIsSendContent
    let url = ""
    if(isHttpOrHttpsUrl(key)){  // 以http开头
      // end = key.slice(-1)
      if(key.endsWith("/")){
        // 形如：https://wxpusher.zjiecode.com/api/send/message/你获取到的SPT/
        url = key + message 
      }else if(key.endsWith("ThisIsSendContent")){
        // 形如：https://wxpusher.zjiecode.com/api/send/message/xxxx/ThisIsSendContent
        key = key.slice(0, -"ThisIsSendContent".length);  // 去掉末尾的"ThisIsSendContent"
        url = key + message 
      }else{
        // 形如：https://wxpusher.zjiecode.com/api/send/message/你获取到的SPT
        url = key + "/" + message  
      }
    }else{
      // 形如：你获取到的SPT
      url = "https://wxpusher.zjiecode.com/api/send/message/" + key + "/" + message
    }
    // console.log(url)
    // let resp = HTTP.fetch(url, {
    //   method: "get",
    // });
    headers = {}
    resp = HTTP.get(url, {headers: headers,});
    // console.log(resp.text())
  }else{
    // console.log("采用标准推送")
    let appToken = keyarry[0]
    let uid = keyarry[1]
    let url = ""
    if(isHttpOrHttpsUrl(key)){  // 以http开头
      url = key + "&verifyPayType=0&content=" + message 
    }else{
      url = "https://wxpusher.zjiecode.com/api/send/message/?appToken=" + appToken + "&uid=" + uid + "&verifyPayType=0&content=" + message 
    }
    // console.log(url)
    // let resp = HTTP.fetch(url, {
    //   method: "get",
    // });
    headers = {}
    resp = HTTP.get(url, {headers: headers,});
    // console.log(resp.json())
  }
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


// =================自定义的非正常MD5开始======================
function rotateLeft(n, d) {
    return (n << d) | (n >>> (32 - d));
}

function F(x, y, z) {
    return (x & y) | (~x & z);
}

function G(x, y, z) {
    return (x & z) | (y & ~z);
}

function H(x, y, z) {
    return x ^ y ^ z;
}

function I(x, y, z) {
    return y ^ (x | ~z);
}

function FF(a, b, c, d, x, s, ac) {
    a = a + F(b, c, d) + x + ac;
    return (rotateLeft(a, s) + b) | 0;
}

function GG(a, b, c, d, x, s, ac) {
    a = a + G(b, c, d) + x + ac;
    return (rotateLeft(a, s) + b) | 0;
}

function HH(a, b, c, d, x, s, ac) {
    a = a + H(b, c, d) + x + ac;
    return (rotateLeft(a, s) + b) | 0;
}

function II(a, b, c, d, x, s, ac) {
    a = a + I(b, c, d) + x + ac;
    return (rotateLeft(a, s) + b) | 0;
}

function cryptoMd5(message) {
    const S11 = 7, S12 = 12, S13 = 17, S14 = 22;
    const S21 = 5, S22 = 9, S23 = 14, S24 = 20;
    const S31 = 4, S32 = 11, S33 = 16, S34 = 23;
    const S41 = 6, S42 = 10, S43 = 15, S44 = 21;

    const K1 = 0xd76aa478, K2 = 0xe8c7b756, K3 = 0x242070db, K4 = 0xc1bdceee;
    const K5 = 0xf57c0faf, K6 = 0x4787c62a, K7 = 0xa8304613, K8 = 0xfd469501;
    const K9 = 0x698098d8, K10 = 0x8b44f7af, K11 = 0xffff5bb1, K12 = 0x895cd7be;
    const K13 = 0x6b901122, K14 = 0xfd987193, K15 = 0xa679438e, K16 = 0x49b40821;
    const K17 = 0xf61e2562, K18 = 0xc040b340, K19 = 0x265e5a51, K20 = 0xe9b6c7aa;
    const K21 = 0xd62f105d, K22 = 0x2441453, K23 = 0xd8a1e681, K24 = 0xe7d3fbc8;
    const K25 = 0x21e1cde6, K26 = 0xc33707d6, K27 = 0xf4d50d87, K28 = 0x455a14ed;
    const K29 = 0xa9e3e905, K30 = 0xfcefa3f8, K31 = 0x676f02d9, K32 = 0x8d2a4c8a;
    const K33 = 0xfffa3942, K34 = 0x8771f681, K35 = 0x6d9d6122, K36 = 0xfde5380c;
    const K37 = 0xa4beea44, K38 = 0x4bdecfa9, K39 = 0xf6bb4b60, K40 = 0xbebfbc70;
    const K41 = 0x289b7ec6, K42 = 0xeaa127fa, K43 = 0xd4ef3085, K44 = 0x4881d05;
    const K45 = 0xd9d4d039, K46 = 0xe6db99e5, K47 = 0x1fa27cf8, K48 = 0xc4ac5665;
    const K49 = 0xf4292244, K50 = 0x432aff97, K51 = 0xab9423a7, K52 = 0xfc93a039;
    const K53 = 0x655b59c3, K54 = 0x8f0ccc92, K55 = 0xffeff47d, K56 = 0x85845dd1;
    const K57 = 0x6fa87e4f, K58 = 0xfe2ce6e0, K59 = 0xa3014314, K60 = 0x4e0811a1;
    const K61 = 0xf7537e82, K62 = 0xbd3af235, K63 = 0x2ad7d2bb, K64 = 0xeb86d391;

    let len = message.length * 8;
    message += '\x80';
    while ((message.length * 8) % 512 != 448) {
        message += '\x00';
    }
    message += String.fromCharCode(len >> 24, (len >> 16) & 0xFF, (len >> 8) & 0xFF, len & 0xFF);

    let a = 0x67452301, b = 0xefcdab89, c = 0x98badcfe, d = 0x10325476;

    for (let i = 0; i < message.length; i += 64) {
        let olda = a, oldb = b, oldc = c, oldd = d;

        for (let j = 0; j < 64; j++) {
            let t;
            if (j < 16) {
                t = FF(a, b, c, d, message.charCodeAt(i + (j * 4)) << 24 | message.charCodeAt(i + (j * 4 + 1)) << 16 | message.charCodeAt(i + (j * 4 + 2)) << 8 | message.charCodeAt(i + (j * 4 + 3)), S11, K1);
            } else if (j < 32) {
                t = GG(a, b, c, d, message.charCodeAt(i + (1 + (j * 4) % 16)) << 24 | message.charCodeAt(i + (6 + (j * 4) % 16)) << 16 | message.charCodeAt(i + (11 + (j * 4) % 16)) << 8 | message.charCodeAt(i + (0 + (j * 4) % 16)), S21, K2);
            } else if (j < 48) {
                t = HH(a, b, c, d, message.charCodeAt(i + (5 + (j * 4) % 16)) << 24 | message.charCodeAt(i + (8 + (j * 4) % 16)) << 16 | message.charCodeAt(i + (11 + (j * 4) % 16)) << 8 | message.charCodeAt(i + (14 + (j * 4) % 16)), S31, K3);
            } else {
                t = II(a, b, c, d, message.charCodeAt(i + (0 + (j * 4) % 16)) << 24 | message.charCodeAt(i + (7 + (j * 4) % 16)) << 16 | message.charCodeAt(i + (14 + (j * 4) % 16)) << 8 | message.charCodeAt(i + (5 + (j * 4) % 16)), S41, K4);
            }

            a = d;
            d = c;
            c = b;
            b = t;
        }

        a = (a + olda) | 0;
        b = (b + oldb) | 0;
        c = (c + oldc) | 0;
        d = (d + oldd) | 0;
    }

    return (a >>> 0).toString(16).padStart(8, '0') +
           (b >>> 0).toString(16).padStart(8, '0') +
           (c >>> 0).toString(16).padStart(8, '0') +
           (d >>> 0).toString(16).padStart(8, '0');
}
// =================自定义的非正常MD5结束======================

// =========================天气处理接口结束========================
// 根据输入，获取模式
function getmodeAdvance(mode){

  // 模式1（mode==1）：今日天气预报
  // 模式2（mode==2）：实时天气预报
  // 模式3（mode==3）：当天生活指数
  // 模式4（mode==4）：天气灾害预警
  // 模式5（mode==5）：逐渐小时天气预报
  // 模式6（mode==6）：分钟级降水
  if(mode == "1" || mode == "今日天气预报" || mode == "每日天气预报"){
    mode = 1
  }else if(mode == "2" || mode == "实时天气预报"){
    mode = 2
  }else if(mode == "3" || mode == "当天生活指数"){
    mode = 3
  }else if(mode == "4" || mode == "天气灾害预警"){
    mode = 4
  }else if(mode == "5" || mode == "逐小时天气预报"){
    mode = 5
  }else if(mode == "6" || mode == "分钟级降水" || mode == "分钟级降雨"){
    mode = 6
  }else{
    mode = 1
  }

  mode = parseInt(mode)
  return mode
}

function getmode(pos){
  let result = [0,0,0,0,0,0]
  todyWeather = Application.Range("E" + pos).Text; // 今日天气预报
  nowWeather = Application.Range("F" + pos).Text  // 实时天气预报
  indexWeather = Application.Range("G" + pos).Text  // 当天生活指数
  disasterWeather = Application.Range("H" + pos).Text  // 天气灾害预报
  hourWeather = Application.Range("I" + pos).Text    // 逐小时天气预报
  minuteWeather = Application.Range("J" + pos).Text  // 分钟级降水

  if(todyWeather == "是"){
    result[0] = 1
  }
  
  if(nowWeather == "是"){
    result[1] = 1
  }
  
  if(indexWeather == "是"){
    result[2] = 1
  }
  
  if(disasterWeather == "是"){
    result[3] = 1
  }
  
  if(hourWeather == "是"){
    result[4] = 1
  }
  
  if(minuteWeather == "是"){
    result[5] = 1
  }
  

  // console.log(result)
  return result
}

// 通过geoapi获得地址信息，包含id、经纬度等。返回某个数据和冗余信息
function getGeo(location){
  console.log("✨️ 获取地理坐标数据")
  let result = []
  url = `https://geoapi.qweather.com/v2/city/lookup?location=${location}&key=${apikey}`
  // console.log(url)

  resp = HTTP.get(
    url,
    // data,
    // headers 
  );
  resp = resp.json()

  sleep(2000)
  // console.log(resp)
  // respcode = resp["code"]
  // if(respcode == "200"){

  // }else{

  // }

  locationlist = resp["location"]
  id = locationlist[0]["id"]  // id
  lat = locationlist[0]["lat"]  // lat
  lon = locationlist[0]["lon"]  // lon
  country = locationlist[0]["country"]  // country
  name = locationlist[0]["name"]  // name 海淀
  adm1 = locationlist[0]["adm1"]  // adm1 北京市
  adm2 = locationlist[0]["adm2"]  // adm2 北京


  let dict = {
    "name":name,
    "id":id,
    "lat":lat,
    "lon":lon,
    "adm2":adm2,
    "adm1":adm1,
    "country":country,
  }

  // redundancy = String(locationlist) // 冗余信息
  // 将地理信息提出出来，并按行拼接，方便在表格中查看冗余信息
  redundancy = locationlist.map(location => {
    // return `名称: ${location.name}, ID: ${location.id}, 纬度: ${location.lat}, 经度: ${location.lon}, 城市: ${location.adm2}, 省份: ${location.adm1}, 国家: ${location.country}, 时区: ${location.tz}, UTC偏移: ${location.utcOffset}, 是否夏令时: ${location.isDst}, 类型: ${location.type}, 排名: ${location.rank}, 预报链接: ${location.fxLink}\n`;
    return `地区: ${location.adm1}${location.adm2}${location.name}, ID: ${location.id}, 经纬度: ${location.lon},${location.lat} 。`;
    
  })  // .join('');

  result = [dict, redundancy]
  return result
}

// 写入数据到表格中，写入一致性校验
function writeData(pos, location, locationReal, locationId, lonlat, redundancy){
  console.log("✨️ 写入地理坐标数据")
  Application.WrapText = false
  Application.Range("N" + pos).Value2 = locationReal // 服务器返回的定位
  Application.Range("O" + pos).Value2 = locationId // 服务器返回的地区ID
  Application.Range("P" + pos).Value2 = lonlat // 服务器返回的lonlat经纬度
  
  consistency= getConsistency(location, locationReal, locationId)
  Application.Range("Q" + pos).Value2 = consistency // 一致性校验值
  Application.Range("R" + pos).Value2 = redundancy // 冗余值
  // Application.WrapText = true
}

// 一致性校验值
function getConsistency(location, locationReal, locationId){
  console.log("🔒 生成一致性校验值")
  let md5 = ""
  // 计算md5
  let sign = String(location) + "|"  + String(locationReal) + "|"  + String(locationId)
  // console.log(sign)
  md5 = cryptoMd5(sign)
  // console.log(md5)
  return md5
}

// 根据对应的模式进行处理
function modeHandel(pos, mode){
  let messageSuccess = "";
  let messageFail = "";
  
  let result = []

  // 读取高级配置
  advanceConfig = getAdvance(pos)

  // 获取表格内容
  location = Application.Range("D" + pos).Text; // 地区
  // modeAdvance = Application.Range("E" + pos).Text
  filter = Application.Range("L" + pos).Text  // 消息过滤器
  advance = Application.Range("M" + pos).Text  // 高级配置
  locationReal = Application.Range("N" + pos).Text    // 实际定位
  locationId = Application.Range("O" + pos).Text  // 地区ID
  lonlat = Application.Range("P" + pos).Text  // 经纬度
  consistency = Application.Range("Q" + pos).Text  // 一致性校验

  // console.log(mode)
  modeList = getmode(pos); // 模式
  // console.log(mode)

  let geoList = []  // 地理信息
  // 一致性校验
  consistencyChallenge = getConsistency(location, locationReal, locationId)
  if(consistencyChallenge == consistency){
    console.log("✅ 一致性校验通过")
    console.log("⚡️ 使用缓存数据")
  }else{
    console.log("🐧 获取最新地理信息")
    // 获取地理信息
    geoList = getGeo(location)
    geoDict = geoList[0]  // 地理信息字典
    redundancy = geoList[1] // 冗余信息
    // console.log(redundancy)

    locationId = geoDict["id"]  // id

    
    lon = geoDict["lon"]  // lon  longitude经度
    lat = geoDict["lat"]  // lat  latitude纬度
    lonlat = String(lon) + "," + String(lat)
    // console.log(lonlat)
    
    // country = geoDict["country"]  // country
    name = geoDict["name"]  // name 海淀
    adm1 = geoDict["adm1"]  // adm1 北京市
    adm2 = geoDict["adm2"]  // adm2 北京 

    if(name == adm1){ // 如果相同则只写入一个
      // locationReal = country + adm2 + name
      locationReal = adm2 + name
    }else if(name == adm2){
      // locationReal = country + adm1 + name
      locationReal = adm1 + name
    }else if(adm1 == adm2){
      // locationReal = country + adm1 + name
      locationReal = adm1 + name
    }else{
      // locationReal = country + adm1 + adm2 + name
      locationReal = adm1 + adm2 + name
    }
    
    // console.log(locationReal)

    // 写入地理信息
    writeData(pos, location, locationReal, locationId, lonlat, redundancy)

  }  

  let count = 0
  let respcode = 0
  for(let i = 1; i<=modeList.length; i++){
    if(modeList[i-1] != 0){

      if(count >= 1){
        messageSuccess += separator;
      }

      mode = i  
      // console.log("=============")
      // console.log(modeList[i-1])
      // console.log(mode)

      // 利用locationId、lonlat查询天气情况。根据模式进行数据选择
      // console.log(mode)
      if(mode == 1){
        console.log("⏲️ 今日天气查询")
        url = `https://devapi.qweather.com/v7/weather/3d?location=${locationId}&key=${apikey}`
        // console.log(url)
        resp = HTTP.get(
          url,
          // data,
          // headers 
        );

        resp = resp.json()
        respcode = resp["code"]
        if(respcode == 200){
          // 天气数据
          weatherData = resp["daily"][0]  // 只要一天，即今天
          // console.log(weatherData)
          updateTime = resp["updateTime"] // 更新时间
          weatherData = weatherfilter(mode, weatherData, filter)
          content =   dashed + "今日天气预报" + dashed +  "\n" + "📅 " + locationReal + " " + weatherData + "\n"
          messageSuccess += content;
          // console.log(content);

        }else{
          content = "❌ 今日天气查询失败" + "\n"
          messageFail += content;
          console.log(content);
        }

      }else if(mode == 2){

        console.log("⏰️ 实时天气查询")
        url = `https://devapi.qweather.com/v7/weather/now?location=${locationId}&key=${apikey}`;

        resp = HTTP.get(
          url,
          // data,
          // headers 
        );

        resp = resp.json()
        respcode = resp["code"]
        if(respcode == 200){
          // 天气数据
          weatherData = resp["now"] // 当前
          // console.log(weatherData)
          updateTime = resp["updateTime"] // 更新时间
          
          weatherData = weatherfilter(mode, weatherData, filter)
          content =   dashed +  "实时天气预报"  + dashed +  "\n" + "📅 " + locationReal + " " + weatherData + "\n"
          messageSuccess += content;
          // console.log(content);

        }else{
          content = "❌ 实时天气查询失败" + "\n"
          messageFail += content;
          console.log(content);
        }

      }else if(mode == 3){

        console.log("🎫 当天生活指数")
        // type = "0" // 0
        // type = "1,2,3,5,6,8,9,14,15"
        type = "1,3,9,15"
        url = `https://devapi.qweather.com/v7/indices/1d?type=${type}&location=${locationId}&key=${apikey}`;

        resp = HTTP.get(
          url,
          // data,
          // headers 
        );

        resp = resp.json()
        respcode = resp["code"]
        if(respcode == 200){
          // 天气数据
          //  weatherData = resp["daily"][0]  // 只要一天，即今天
          weatherData = resp
          // console.log(weatherData)
          updateTime = resp["updateTime"] // 更新时间
          weatherData = weatherfilter(mode, weatherData, filter)
          content =  dashed +  "当天生活指数" + dashed + "\n" +  "📅 " + locationReal + " " + weatherData + "\n"
          messageSuccess += content;
          // console.log(content);

        }else{
          content = "❌ 当天生活指数查询失败" + "\n"
          messageFail += content;
          console.log(content);
        }

      }else if(mode == 4){

        console.log("🚨 天气灾害预警查询")
        url = `https://devapi.qweather.com/v7/warning/now?location=${locationId}&key=${apikey}`;

        resp = HTTP.get(
          url,
          // data,
          // headers 
        );

        resp = resp.json()
        respcode = resp["code"]
        if(respcode == 200){
          // 天气数据
          //  weatherData = resp["warning"][0]  // 只要一天，即今天
          weatherData = resp
          // console.log(weatherData)
          updateTime = resp["updateTime"] // 更新时间
          weatherData = weatherfilter(mode, weatherData, filter)
          content =  dashed + "天气灾害预警" + dashed + "\n" + "📅 " + locationReal + " " + weatherData + "\n"
          messageSuccess += content;
          // console.log(content);

        }else{
          content = "❌ 天气灾害预警查询失败" + "\n"
          messageFail += content;
          console.log(content);
        }
      }else if(mode == 5){
        
        console.log("🕰️ 逐小时天气预报查询")
        url = `https://devapi.qweather.com/v7/weather/24h?location=${locationId}&key=${apikey}`;
        // console.log(url)

        resp = HTTP.get(
          url,
          // data,
          // headers 
        );

        resp = resp.json()
        respcode = resp["code"]
        if(respcode == 200){
          weatherData = resp
          // console.log(weatherData)
          updateTime = resp["updateTime"] // 更新时间
          weatherData = weatherfilter(mode, weatherData, filter)
          content =  dashed + "逐小时天气预报" + dashed + "\n" + "📅 " + locationReal + " " + weatherData + "\n"
          messageSuccess += content;
          // console.log(content);

        }else{
          content = "❌ 逐小时天气预报查询失败" + "\n"
          messageFail += content;
          console.log(content);
        }
      }else if(mode == 6){

        console.log("⏱️ 分钟级降水查询")
        // 需要传入经纬度
        url = `https://devapi.qweather.com/v7/minutely/5m?location=${lonlat}&key=${apikey}`;
        // console.log(url)

        resp = HTTP.get(
          url,
          // data,
          // headers 
        );

        resp = resp.json()
        respcode = resp["code"]
        if(respcode == 200){
          weatherData = resp
          // console.log(weatherData)
          updateTime = resp["updateTime"] // 更新时间
          weatherData = weatherfilter(mode, weatherData, filter)
          content =  dashed + "分钟级降水" + dashed + "\n" + "📅 " + locationReal + " " + weatherData + "\n"
          messageSuccess += content;
          // console.log(content);

        }else{
          content = "❌ 分钟级降水查询失败" + "\n"
          messageFail += content;
          console.log(content);
        }
      }
      
      count += 1
      sleep(2000)

    }
  }

  result =[respcode, messageSuccess, messageFail]
  return result

}

// =========================天气处理接口结束========================


// =================消息分片处理相关开始===================
// 去除首尾换行和空格
function customTrim(str) {
  return str.replace(/^\s+|\s+$/g, '');
}

// 纯长度分片
function splitMessageSimple(data) {
    let chunks = [];
    for (let i = 0; i < data.length; i += maxMessageLength) {
      console.log(i, i + maxMessageLength)
        chunks.push(data.slice(i, i + maxMessageLength));
    }
    return chunks

    // chunks.forEach((chunk, index) => {
    //     // let message = `${index + 1}/${chunks.length}: ${chunk}`;
    //     bark(message, key)
    // });
}


// 消息分片，以换行符为分割，自动检索切割位置符号
function splitMessage(data) {
    let chunks = [];
    let start = 0;

    while (start < data.length) {
        let end = start + maxMessageLength;
        if (end >= data.length) {
            chunks.push(data.slice(start));
            break;
        }

        // 查找距离 maxMessageLength 在 20 字符以内的最近的换行符
        let newlineIndex = data.lastIndexOf('【', end + parseInt(messageDistance));
        // console.log(newlineIndex)
        if (newlineIndex > start && newlineIndex >= end - parseInt(messageDistance)) {
            end = newlineIndex;
        }
        chunks.push(data.slice(start, end));
        start = end;
    }

     return chunks
}

// =================消息分片处理相关结束===================

// 结果处理函数
function resultHandle(resp, pos){
  // 每次进来resultHandle则加一次请求
  posHttp += 1    // 青龙适配，青龙微适配

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
  // messageHeader[posLabel] = "👨‍🚀 " + messageName + "\n"
  // messageHeader[posLabel] = "\n"
  messageHeader[posLabel] = ""
  // console.log(messageName)

  // 实际运行
  result = modeHandel(pos, mode)
  respcode = result[0]
  messageSuccess = result[1]
  messageFail = result[2]

  // messageSuccess = "测试" // 测试


  // 青龙适配，青龙微适配
  flagResultFinish = 1; // 结束

  // 检查是否直接推送
  flag_pushdirect = Application.Range("S" + pos).Text
  if(flag_pushdirect == "是") {
    // console.log("🚀 直接推送")
    // pushDirect(messageSuccess);
    message = messageSuccess
    // 方式1：按照指定分割符分片  separator
    shards = message.split(separator); // // 分割符分片数据，一级分割
    let chunks = []
    for(let i=0; i<shards.length; i++){
      strTrim = customTrim(shards[i]) + "\n\n" // 消息内间隔。去除首位空格和换行，然后在末尾拼接2个换行。
      chunks = splitMessage(strTrim)  // 长度限制分割，二级分割
      // console.log(chunks)
      // console.log(chunks.length)
      for (let j = 0; j < chunks.length; j++) {
          // pushMessage(chunks[j], msgCurrentDict.methodPush, "【" + msgCurrentDict.note + "】",)
          message = chunks[j]
          // console.log(message)
          pushDirect(message);
          sleep(2000)
      }
    }

  } else {
    if (messageOnlyError == 1) {
      messageArray[posLabel] =  messageFail;
    } else {
        if(messageFail != ""){
          messageArray[posLabel] = messageFail + " " + messageSuccess;
        }else{
          messageArray[posLabel] = messageSuccess;
        }
    }
  }

  sleep(2000);

  // if(messageArray[posLabel] != "")
  // {
  //   // console.log(messageArray[posLabel]);
  // }
//   console.log(messageArray)

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

    apikey = cookie // apikey

    resp = ""
    if(qlSwitch != 1){  // 选择金山文档
        resultHandle(resp, pos)
    }
}

// 2024-11-27 23:21:42转ISO格式2024-11-27T15:31:57.158Z
function ymdhmstoISO(dateTimeString){
  // 解析日期时间字符串为Date对象
  // console.log(dateTimeString)
  const date = new Date(dateTimeString.replace(' ', 'T'));

  // 手动构建ISO格式的字符串
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  const hours = String(date.getHours()).padStart(2, '0');
  const minutes = String(date.getMinutes()).padStart(2, '0');
  const seconds = String(date.getSeconds()).padStart(2, '0');
  const milliseconds = String(date.getMilliseconds()).padStart(3, '0');

  const localISOFormat = `${year}-${month}-${day}T${hours}:${minutes}:${seconds}.${milliseconds}Z`;

  // console.log(localISOFormat); // 输出类似：2024-11-27T23:21:42.000
  return localISOFormat
}

// 当前时间2024-11-27 23:21:42
function currentTime() {
    let now = new Date();
    let year = now.getFullYear();
    let month = String(now.getMonth() + 1).padStart(2, '0');
    let day = String(now.getDate()).padStart(2, '0');
    let hours = String(now.getHours()).padStart(2, '0');
    let minutes = String(now.getMinutes()).padStart(2, '0');
    let seconds = String(now.getSeconds()).padStart(2, '0');

    let formattedDate = `${year}-${month}-${day} ${hours}:${minutes}:${seconds}`;

    return formattedDate
}

// 获取当前时间 ISO格式
function getNowTimeISO(){
  let now = currentTime()
  now = ymdhmstoISO(now)
  return now
}

// 当前时间hh:mm:ss
function getNowTime() {
  return new Date().toTimeString().split(' ')[0]
}

// ISO->hh:mm:ss
function getTimeFromISO(isoString) {
  const date = new Date(isoString);
  return date.toTimeString().split(' ')[0];
}

// ISO->hh:mm
function isoTohm(isoString) {
  const now = new Date(isoString);
  let hours = now.getHours().toString().padStart(2, '0');
  let minutes = now.getMinutes().toString().padStart(2, '0');
  return `${hours}:${minutes}`;
}

// 获取小时.hh:mm - > h
function getHour(time) {
  // 使用冒号分隔时间字符串
  const [hour, minute] = time.split(':');
  // 返回小时部分
  return hour;
}


// 比较hh:mm:ss。前者大返回1，后者大-1，相等0
function compareHHMMSS(time1, time2) {
  // 解析时间字符串
  const [hours1, minutes1, seconds1] = time1.split(':').map(Number);
  const [hours2, minutes2, seconds2] = time2.split(':').map(Number);

  // 比较小时
  if (hours1 < hours2) return -1;
  if (hours1 > hours2) return 1;

  // 比较分钟
  if (minutes1 < minutes2) return -1;
  if (minutes1 > minutes2) return 1;

  // 比较秒
  if (seconds1 < seconds2) return -1;
  if (seconds1 > seconds2) return 1;

  // 时间相等
  return 0;
}

// 比较ISO时间大小，前者时间大返回1
function compareISOTimes(time1, time2) {
    // 解析ISO格式的时间字符串
    const dt1 = new Date(time1);
    const dt2 = new Date(time2);

    // 比较两个Date对象
    if (dt1 < dt2) return -1;
    if (dt1 > dt2) return 1;
    return 0; // 相等
}

// 年月日转月日
function ymdtomd(dateStr){
  // const dateStr = '2024-11-27';
  const [year, month, day] = dateStr.split('-');
  const formattedDate = `${month}-${day}`;
  // console.log(formattedDate); // 输出: 11-27
  return formattedDate
}

// ISO获得对应emoji
function timeToemoji(isoTime) {
  // 定义每个小时对应的emoji
  const hourEmojis = [
    { hour: 0, emoji: "🕛" },
    { hour: 1, emoji: "🕐" },
    { hour: 2, emoji: "🕑" },
    { hour: 3, emoji: "🕒" },
    { hour: 4, emoji: "🕓" },
    { hour: 5, emoji: "🕔" },
    { hour: 6, emoji: "🕕" },
    { hour: 7, emoji: "🕖" },
    { hour: 8, emoji: "🕗" },
    { hour: 9, emoji: "🕘" },
    { hour: 10, emoji: "🕙" },
    { hour: 11, emoji: "🕚" },
    { hour: 12, emoji: "🕛" },
    { hour: 13, emoji: "🕐" },
    { hour: 14, emoji: "🕑" },
    { hour: 15, emoji: "🕒" },
    { hour: 16, emoji: "🕓" },
    { hour: 17, emoji: "🕔" },
    { hour: 18, emoji: "🕕" },
    { hour: 19, emoji: "🕖" },
    { hour: 20, emoji: "🕗" },
    { hour: 21, emoji: "🕘" },
    { hour: 22, emoji: "🕙" },
    { hour: 23, emoji: "🕚" }
  ];

    // 解析ISO时间
    const date = new Date(isoTime);
    hours = date.getHours();

  // 找到对应的emoji
  const closestEmoji = hourEmojis.find(timeEmoji => timeEmoji.hour === hours);

  return closestEmoji ? closestEmoji.emoji : "⏲️"; // 默认返回闹钟emoji
}

// hour获取对应emoji
function hourToemoji(hours) {
  // 定义每个小时对应的emoji
  const hourEmojis = [
    "🕛", "🕐", "🕑", "🕒", "🕓", "🕔", "🕕", "🕖", "🕗", "🕘", "🕙", "🕚"
  ];

  // 将小时数转换为0-11范围
  const adjustedHour = hours % 12;

  // 返回对应的emoji，如果小时数不在0-23范围内，返回默认的闹钟emoji
  return adjustedHour >= 0 && adjustedHour < 12 ? hourEmojis[adjustedHour] : "⏲️";
}

// 若需要自定义天气格式请改这个函数的内容
// =========================天气消息开始========================

// 模式1（mode==1）：今日天气预报
// 模式2（mode==2）：实时天气预报
// 模式3（mode==3）：当天生活指数
// 模式4（mode==4）：天气灾害预警
// 模式5（mode==5）：逐渐小时天气预报
// 模式6（mode==6）：分钟级降水

// 若需要自定义天气格式请改这个函数的内容
// =========================天气消息开始========================
// 带文字的天气消息数据
// 天气数据过滤处理，合成消息数据
function weatherfilter(mode, weatherData, filter){
    let weatherMessage = ""
    filter = String(filter)
  
    if(mode == 1){
      // 天气数据变量
      let fxDate = ""
      let sunrise = ""
      let sunset = ""
      let moonrise = ""
      let moonset = ""
      let moonPhase = ""
      let moonPhaseIcon = ""
      let tempMax = ""
      let tempMin = ""
      let iconDay = ""
      let textDay = ""
      let iconNight = ""
      let textNight = ""
      let wind360Day = ""
      let windDirDay = ""
      let windScaleDay = ""
      let windSpeedDay = ""
      let wind360Night = ""
      let windDirNight = ""
      let windScaleNight = ""
      let windSpeedNight = ""
      let humidity = ""
      let precip = ""
      let pressure = ""
      let vis = ""
      let cloud = ""
      let uvIndex = ""
  
      fxDate = weatherData["fxDate"];
      sunrise = weatherData["sunrise"];
      sunset = weatherData["sunset"];
      moonrise = weatherData["moonrise"];
      moonset = weatherData["moonset"];
      moonPhase = weatherData["moonPhase"];
      moonPhaseIcon = weatherData["moonPhaseIcon"];
      tempMax = weatherData["tempMax"];
      tempMin = weatherData["tempMin"];
      iconDay = weatherData["iconDay"];
      textDay = weatherData["textDay"];
      iconNight = weatherData["iconNight"];
      textNight = weatherData["textNight"];
      wind360Day = weatherData["wind360Day"];
      windDirDay = weatherData["windDirDay"];
      windScaleDay = weatherData["windScaleDay"];
      windSpeedDay = weatherData["windSpeedDay"];
      wind360Night = weatherData["wind360Night"];
      windDirNight = weatherData["windDirNight"];
      windScaleNight = weatherData["windScaleNight"];
      windSpeedNight = weatherData["windSpeedNight"];
      humidity = weatherData["humidity"];
      precip = weatherData["precip"];
      pressure = weatherData["pressure"];
      vis = weatherData["vis"];
      cloud = weatherData["cloud"];
      uvIndex = weatherData["uvIndex"];
      
      // 根据过滤进行处理
      if(filter == "" || filter == "undefined"){
        // weatherMessage += "📅 " + fxDate + "\n"

        // 可自行拓展
        // 根据高级配置进行个性化服务

        // // 行内分割符号
        // lineDelimiter = advanceConfig["lineDelimiter"]  // 行内分割符号
        // if(lineDelimiter == undefined){
        //   // console.log("无lineDelimiter配置")
        //   lineDelimiter = ""
        // }

        // // 行内数量
        // lineNumber = advanceConfig["lineNumber"] 
        // if(lineNumber == undefined){
        //   lineNumber = 2  // 默认为2
        // }

        // 行内分割符号，默认为空
        let lineDelimiter = advanceConfig["lineDelimiter"] || "";
        // 行内数量，默认为2
        let lineNumber = advanceConfig["lineNumber"] || 2;

        // console.log(lineDelimiter)
        // console.log(lineNumber)

        fxDate = ymdtomd(fxDate)
        weatherMessage += "" + fxDate + "\n"

        weatherMessage += "🌞 白天: " + textDay + "\t"
        weatherMessage += "🌙 夜间: " + textNight + "\n"

        weatherMessage += "🌅 日出: " + sunrise + "\t"
        weatherMessage += "🌇 日落: " + sunset + "\n"

        weatherMessage += "🌙 月升: " + moonrise + "\t"
        weatherMessage += "🌛 月落: " + moonset + "\n"

        weatherMessage += "☁️ 云量: " + cloud + "%" + "\t"
        weatherMessage += "🌙 月相: " + moonPhase + "\n"

        weatherMessage += "👀 能见度: " + vis + "km"  + "\t"
        weatherMessage += "☀️ 紫外线指数: " + uvIndex + "\n"
        
        weatherMessage += "🌬️ 白天风: " + windDirDay + " " + windSpeedDay + "km/h (" + windScaleDay + "级)\n"
        weatherMessage += "🌬️ 夜间风: " + windDirNight + " " + windSpeedNight + "km/h (" + windScaleNight + "级)\n"

        weatherMessage += "💧 相对湿度: " + humidity + "%" + "\t"
        weatherMessage += "🌧️ 降水量: " + precip + "mm" + "\n"

        weatherMessage += "🌡️ 最高温/最低温: " + tempMax + "°C / " + tempMin + "°C\n"

        weatherMessage += "🎈 大气压强: " + pressure + "hPa"  + "\n"

        // weatherMessage += "📅  " + fxDate + "\t";
        // weatherMessage += "🌅  日出: " + sunrise + "\t";
        // weatherMessage += "🌇  日落: " + sunset + "\n";

        // weatherMessage += "🌙  月出: " + moonrise + "\t";
        // weatherMessage += "🌛  月落: " + moonset + "\t";
        // weatherMessage += "🌙  月相: " + moonPhase + "\n";

        // weatherMessage += "🌡️ 最高温度/最低温度: " + tempMax + "°C / " + tempMin + "°C\n"

        // weatherMessage += "🌞  白天: " + textDay + " (" + iconDay + ")\t";
        // weatherMessage += "🌙  夜间: " + textNight + " (" + iconNight + ")\n";

        // weatherMessage += "🌬️  白天风: " + windDirDay + " " + windSpeedDay + "km/h (" + windScaleDay + "级)\t";
        // weatherMessage += "🌬️  夜间风: " + windDirNight + " " + windSpeedNight + "km/h (" + windScaleNight + "级)\n";

        // weatherMessage += "💧  湿度: " + humidity + "%\t";
        // weatherMessage += "🌧️  降水量: " + precip + "mm\t";
        // weatherMessage += "🎈  气压: " + pressure + "hPa\n";

        // weatherMessage += "👀  能见度: " + vis + "km\t";
        // weatherMessage += "☁️  云量: " + cloud + "%\t";
        // weatherMessage += "☀️  紫外线指数: " + uvIndex + "\n";

        // // 拼接天气消息
        // weatherMessage += "📅 预报日期: " + fxDate + "\n"

        // weatherMessage += "🌅 日出时间: " + sunrise + "\t"
        // weatherMessage += "🌇 日落时间: " + sunset + "\n"

        // weatherMessage += "🌙 月升时间: " + moonrise + "\t"
        // weatherMessage += "🌛 月落时间: " + moonset + "\n"

        

        // weatherMessage += "🌡️ 最高温度/最低温度: " + tempMax + "°C / " + tempMin + "°C\n"

        // weatherMessage += "🌞 白天天气状况: " + textDay + " (" + iconDay + ")\n"
        // weatherMessage += "🌙 夜间天气状况: " + textNight + " (" + iconNight + ")\n"
        // weatherMessage += "🌬️ 白天风向: " + windDirDay + " " + windSpeedDay + "km/h (" + windScaleDay + "级)\n"
        // weatherMessage += "🌬️ 夜间风向: " + windDirNight + " " + windSpeedNight + "km/h (" + windScaleNight + "级)\n"

        // weatherMessage += "💧 相对湿度: " + humidity + "%" + "\t"
        // weatherMessage += "🌧️ 降水量: " + precip + "mm" + "\n"

        //  weatherMessage += "👀 能见度: " + vis + "km"  + "\t"
        // weatherMessage += "🎈 大气压强: " + pressure + "hPa"  + "\n"

        // weatherMessage += "☁️ 云量: " + cloud + "%" + " "
        // weatherMessage += "🌙 月相: " + moonPhase + " "
        // weatherMessage += "☀️ 紫外线指数: " + uvIndex + "\n"
        
  
      }else{
        // 带过滤
        // 遍历 filterArray 并构建 weatherMessage
        let filterArray = filter.split('&');
        filterArray.forEach(key => {
            switch (key) {
                case 'fxDate':
                    weatherMessage += "📅 预报日期: " + fxDate + "\n";
                    break;
                case 'sunrise':
                    weatherMessage += "🌅 日出时间: " + sunrise + "\n";
                    break;
                case 'sunset':
                    weatherMessage += "🌇 日落时间: " + sunset + "\n";
                    break;
                case 'moonrise':
                    weatherMessage += "🌙 月升时间: " + moonrise + "\n";
                    break;
                case 'moonset':
                    weatherMessage += "🌛 月落时间: " + moonset + "\n";
                    break;
                case 'moonPhase':
                    weatherMessage += "🌙 月相: " + moonPhase + "\n";
                    break;
                case 'tempMax':
                case 'tempMin':
                    weatherMessage += "🌡️ 最高温度/最低温度: " + tempMax + "°C / " + tempMin + "°C\n";
                    break;
                case 'textDay':
                    weatherMessage += "🌞 白天天气状况: " + textDay + " (" + iconDay + ")\n";
                    break;
                case 'textNight':
                    weatherMessage += "🌙 夜间天气状况: " + textNight + " (" + iconNight + ")\n";
                    break;
                case 'windDirDay':
                case 'windSpeedDay':
                case 'windScaleDay':
                    weatherMessage += "🌬️ 白天风向: " + windDirDay + " " + windSpeedDay + "km/h (" + windScaleDay + "级)\n";
                    break;
                case 'windDirNight':
                case 'windSpeedNight':
                case 'windScaleNight':
                    weatherMessage += "🌬️ 夜间风向: " + windDirNight + " " + windSpeedNight + "km/h (" + windScaleNight + "级)\n";
                    break;
                case 'humidity':
                    weatherMessage += "💧 相对湿度: " + humidity + "%\n";
                    break;
                case 'precip':
                    weatherMessage += "🌧️ 降水量: " + precip + "mm\n";
                    break;
                case 'pressure':
                    weatherMessage += "🎈 大气压强: " + pressure + "hPa\n";
                    break;
                case 'vis':
                    weatherMessage += "👀 能见度: " + vis + "km\n";
                    break;
                case 'cloud':
                    weatherMessage += "☁️ 云量: " + cloud + "%\n";
                    break;
                case 'uvIndex':
                    weatherMessage += "☀️ 紫外线指数: " + uvIndex + "\n";
                    break;
            }
        });
  
      }
  
    }else if (mode == 2){
      // 🕒 记录时间
      let obsTime = ""
      // 🌡️ 温度
      let temp = ""
      // 🌡️ 体感温度
      let feelsLike = ""
      // 🌤️ 天气图标代码
      let iconDay = ""
      // 🌤️ 天气描述
      let textDay = ""
      // 💨 风向角度
      let wind360 = ""
      // 💨 风向
      let windDir = ""
      // 💨 风力等级
      let windScale = ""
      // 💨 风速
      let windSpeed = ""
      // 💧 湿度
      let humidity = ""
      // 🌦️ 降水量
      let precip = ""
      // 📈 气压
      let pressure = ""
      // 👀 能见度
      let vis = ""
      // ☁️ 云量
      let cloud = ""
      // 🌞 露点温度
      let dew = ""
  
      obsTime = weatherData["obsTime"]
      temp = weatherData["temp"]
      feelsLike = weatherData["feelsLike"]
      iconDay = weatherData["icon"]
      textDay = weatherData["text"]
      wind360 = weatherData["wind360"]
      windDir = weatherData["windDir"]
      windScale = weatherData["windScale"]
      windSpeed = weatherData["windSpeed"]
      humidity = weatherData["humidity"]
      precip = weatherData["precip"]
      pressure = weatherData["pressure"]
      vis = weatherData["vis"]
      cloud = weatherData["cloud"]
      dew = weatherData["dew"]
  
      if(filter == "" || filter == "undefined"){
        // weatherMessage += "🕒 观测时间: " + obsTime + "\n"
        fxDate = isoTohm(obsTime)
        weatherMessage += "" + fxDate + "\n"

        weatherMessage += "🌞 天气: " + textDay + "\n"

        weatherMessage += "🌡️ 温度: " + temp + "°C" + "\t"
        weatherMessage += "🌡️ 体感温度: " + feelsLike + "°C" + "\n"

       
        // weatherMessage += "🌤️ 天气图标代码: " + iconDay + "\n"

        weatherMessage += "🌬️ 风向角度: " + wind360 + "°" + "\t"
        weatherMessage += "💨 风向: " + windDir + "\n"

        weatherMessage += "🍃 风力等级: " + windScale + "\t"
        weatherMessage += "💨 风速: " + windSpeed + " km/h" + "\n"

        weatherMessage += "💧 湿度: " + humidity + "%" + "\t"
        weatherMessage += "🌧️ 降水量: " + precip + " mm" + "\n"

        weatherMessage += "📈 气压: " + pressure + " hPa" + "\t"
        weatherMessage += "👀 能见度: " + vis + " km" + "\n"

        weatherMessage += "☁️ 云量: " + cloud + "%" + "\t"
        weatherMessage += "🌞 露点温度: " + dew + "°C" + "\n"
  
      }else{
        // 带过滤
        let filterArray = filter.split('&');
        filterArray.forEach(key => {
            switch (key) {
                case 'obsTime':
                    weatherMessage += "🕒 观测时间: " + obsTime + "\n";
                    break;
                case 'temp':
                    weatherMessage += "🌡️ 温度: " + temp + "°C\n";
                    break;
                case 'feelsLike':
                    weatherMessage += "🌡️ 体感温度: " + feelsLike + "°C\n";
                    break;
                case 'icon':
                    weatherMessage += "🌤️ 天气图标代码: " + iconDay + "\n";
                    break;
                case 'text':
                    weatherMessage += "🌞 天气状况: " + textDay + "\n";
                    break;
                case 'wind360':
                    weatherMessage += "🌬️ 风向角度: " + wind360 + "°\n";
                    break;
                case 'windDir':
                    weatherMessage += "💨 风向: " + windDir + "\n";
                    break;
                case 'windScale':
                    weatherMessage += "🍃 风力等级: " + windScale + "\n";
                    break;
                case 'windSpeed':
                    weatherMessage += "💨 风速: " + windSpeed + " km/h\n";
                    break;
                case 'humidity':
                    weatherMessage += "💧 湿度: " + humidity + "%\n";
                    break;
                case 'precip':
                    weatherMessage += "🌧️ 降水量: " + precip + " mm\n";
                    break;
                case 'pressure':
                    weatherMessage += "📈 气压: " + pressure + " hPa\n";
                    break;
                case 'vis':
                    weatherMessage += "👀 能见度: " + vis + " km\n";
                    break;
                case 'cloud':
                    weatherMessage += "☁️ 云量: " + cloud + "%\n";
                    break;
                case 'dew':
                    weatherMessage += "🌞 露点温度: " + dew + "°C\n";
                    break;
                default:
                    break;
            }
        });
      }
    }else if (mode == 3) {

      // 处理生活指数数据
      let daily = weatherData.daily;
      // weatherMessage += "📅 " + daily[0].date + "\n";
      fxDate = ymdtomd(daily[0].date)
      weatherMessage += "" + fxDate + "\n"


      if (filter == "" || filter == "undefined") {
          // 默认显示所有生活指数
          daily.forEach(index => {
              weatherMessage += "🏷️ " + index.name + "\t";
              // weatherMessage += "💪 " + index.level + "\t";
              // weatherMessage += "🔍 " + index.category + "\n";
              weatherMessage += "：" + index.category + "\n";
              weatherMessage += "💬 " + index.text + "\n";
          });
      } else {
          // 带过滤
          let filterArray = filter.split('&');
          daily.forEach(index => {
              if (filterArray.includes(index.type)) {
                  weatherMessage += "📅 " + index.date + "\n";
                  weatherMessage += "🏷️ " + index.name + "\n";
                  weatherMessage += "💪 " + index.level + "\n";
                  weatherMessage += "🔍 " + index.category + "\n";
                  weatherMessage += "💬 " + index.text + "\n\n";
              }
          });
      }
    }else if (mode == 4) {
        // 处理气象预警数据
        // console.log(weatherData)
        const { updateTime, fxLink, warning, refer } = weatherData;
        
        if (warning.length > 0) {
            weatherMessage +=  "\n" + "🚨 气象预警信息 🚨\n";
            emoji = timeToemoji(new Date(updateTime).toLocaleString())
            weatherMessage += emoji + " 更新时间: " + new Date(updateTime).toLocaleString() + "\n";
            weatherMessage += "🔗 预警详情链接: " + fxLink + "\n\n";

            warning.forEach(warn => {
                weatherMessage += "📢 发布单位: " + warn.sender + "\n";
                weatherMessage += "⏰ 发布时间: " + new Date(warn.pubTime).toLocaleString() + "\n";
                weatherMessage += "🎯 标题: " + warn.title + "\n";
                weatherMessage += "⏰ 开始时间: " + new Date(warn.startTime).toLocaleString() + "\n";
                weatherMessage += "⏰ 结束时间: " + new Date(warn.endTime).toLocaleString() + "\n";
                weatherMessage += "🚧 状态: " + warn.status + "\n";
                weatherMessage += "🎚️ 预警严重等级: " + warn.level + " (" + warn.severity + ")\n";
                weatherMessage += "⚠️ 类型: " + warn.typeName + "\n";
                weatherMessage += "📝 内容: " + warn.text + "\n\n";
            });

            weatherMessage += "📚 数据来源: " + refer.sources.join(", ") + "\n";
            weatherMessage += "📜 许可证: " + refer.license.join(", ") + "\n";
        } else {
            weatherMessage += "🫧 当前没有气象预警信息。";
        }
    }else if (mode == 5) {
        // 逐小时天气预报
        const { code, updateTime, fxLink, hourly, refer } = weatherData;

        if (code !== "200") {
            weatherMessage += "\n" + "🚨 请求失败，状态码: " + code;
        } else {
            // emoji = timeToemoji(new Date(updateTime).toLocaleString())
            // weatherMessage += "\n" + emoji + " 更新时间: " + new Date(updateTime).toLocaleString() + "\n";
            // weatherMessage += "🔗 预报详情链接: " + fxLink + "\n\n";

            weatherMessage += "\n"

            if (filter == "" || filter == "undefined") {
                // 默认显示所有逐小时天气预报

                // 只要在此时间之后的
                // let now = new Date();
                
                let now = getNowTimeISO()
                // let now = new Date()
                // console.log(now)
                let count = 4;  // 最多需要的个数

                hourly.forEach(hour => {

                    // hourTime = getTimeFromISO(hour.fxTime)
                    hourTime = isoTohm(hour.fxTime)
                    // console.log("预报时间:", hourTime);
                    // comp = compareHHMMSS(now, hourTime)
                    
                    comp = compareISOTimes(now, hour.fxTime)
                    if(comp >= 0 && count > 0){
                      // console.log(count)
                      count -= 1

                      // weatherMessage += "🕒 预报时间: " + new Date(hour.fxTime).toLocaleString() + "\n";
                      emoji = timeToemoji(hour.fxTime)
                      weatherMessage += "-----------" + emoji + " " + hourTime + "-----------" + "\n";

                      // weatherMessage += "🌤️ 天气图标代码: " + hour.icon + "\n";
                      
                      weatherMessage += "🌡️ 温度: " + hour.temp + "°C" + "\t";
                      weatherMessage += "🌞 天气: " + hour.text + "\n";

                      weatherMessage += "💨 风向: " + hour.windDir + "\t";
                      weatherMessage += "🌬️ 角度: " + hour.wind360 + "°" + "\n"

                      weatherMessage += "💨 风速: " + hour.windSpeed + " km/h" + "\t"
                      weatherMessage += "🍃 等级: " + hour.windScale + "\n";

                      weatherMessage += "💧 相对湿度: " + hour.humidity + "%" + "\t";
                      weatherMessage += "☁️ 云量: " + (hour.cloud ? hour.cloud + "%" : "无数据") + "\n";

                      weatherMessage += "🌧️ 降水量: " + hour.precip + " mm" + "\t";
                      weatherMessage += "📊 概率: " + (hour.pop ? hour.pop + "%" : "无数据") + "\n";

                      weatherMessage += "\n"

                    }

                    
                    
                    // weatherMessage += "📈 气压: " + hour.pressure + " hPa" + "\t";
                    // weatherMessage += "🌞 露点温度: " + (hour.dew ? hour.dew + "°C" : "无数据") + "\n\n";
                })
            } else {
                // 带过滤
                let filterArray = filter.split('&');
                hourly.forEach(hour => {
                    let includeHour = false;
                    filterArray.forEach(key => {
                        if (hour[key] !== undefined) {
                            includeHour = true;
                        }
                    });

                    if (includeHour) {
                        weatherMessage += "🕒 预报时间: " + new Date(hour.fxTime).toLocaleString() + "\n";
                        filterArray.forEach(key => {
                            if (hour[key] !== undefined) {
                                weatherMessage += key + ": " + hour[key] + "\n";
                            }
                        });
                        weatherMessage += "\n";
                    }
                });
            }

            // weatherMessage += "📚 数据来源: " + refer.sources.join(", ") + "\n";
            // weatherMessage += "📜 许可证: " + refer.license.join(", ") + "\n";
        }
    }else if (mode == 6) {
        const { code, updateTime, fxLink, summary, minutely, refer } = weatherData;

        if (code !== "200") {
            weatherMessage += "\n" + "🚨 请求失败，状态码: " + code;
        } else {   
          
            emoji = timeToemoji(new Date(updateTime).toLocaleString())
            // weatherMessage += "\n" + emoji + " 更新时间: " + new Date(updateTime).toLocaleString() + "\n";
            // weatherMessage += "🔗 预报详情链接: " + fxLink + "\n";
            weatherMessage += "\n📢 降水描述: " + summary + "\n\n";

            if (filter == "" || filter == "undefined") {

                let count = 6;  // 最多需要的个数，6*5min = 30min
                
                // 默认显示所有分钟级别降水预报
                minutely.forEach(minute => {
                  if(count > 0){
                    count -= 1
                    fxTime = isoTohm(new Date(minute.fxTime).toLocaleString())
                    // console.log(minute.fxTime)
                    // weatherMessage += "🕒 预报时间: " + new Date(minute.fxTime).toLocaleString() + "\n";
                    emoji = hourToemoji(getHour(fxTime))
                    // weatherMessage += emoji + " 时间: " + fxTime + "\n";
                    weatherMessage += "-----------" + emoji + " " + fxTime + "-----------" + "\n";
                    weatherMessage += "🌧️ 降水量: " + minute.precip + " mm"; + "\t"
                    weatherMessage += "💦 类型: " + (minute.type === "rain" ? "雨" : "雪") + "\n\n";
                  }
                });
            } else {
                // 带过滤
                let filterArray = filter.split('&');
                minutely.forEach(minute => {
                    let includeMinute = false;
                    filterArray.forEach(key => {
                        if (minute[key] !== undefined) {
                            includeMinute = true;
                        }
                    });

                    if (includeMinute) {
                        weatherMessage += "🕒 预报时间: " + new Date(minute.fxTime).toLocaleString() + "\n";
                        filterArray.forEach(key => {
                            if (minute[key] !== undefined) {
                                weatherMessage += key + ": " + minute[key] + "\n";
                            }
                        });
                        weatherMessage += "\n";
                    }
                });
            }

            // weatherMessage += "📚 数据来源: " + refer.sources.join(", ") + "\n";
            // weatherMessage += "📜 许可证: " + refer.license.join(", ") + "\n";
      }
    }
  
    // console.log(weatherMessage);
    return weatherMessage
}
// =========================天气消息结束========================
