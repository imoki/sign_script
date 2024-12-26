/*
    name: "夸克网盘"
    cron: 10 30 10 * * *
    脚本兼容: 金山文档（1.0），金山文档（2.0）， 青龙
    更新时间：20241226
    环境变量名：quark
    环境变量值：填写手机app抓包的一个GET请求包，并且含sign的url，建议含growth/info的url。
                必须是抓夸克APP中的夸克网盘。
*/

const logo = "艾默库 : https://github.com/imoki/sign_script"    // 仓库地址
var sheetNameSubConfig = "quark"; // 分配置表名称
var pushHeader = "【夸克网盘】";
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
// v2.6.2

try{
  var userContent=[[")\u4E2A02\u8BA4\u9ED8(eikooc".split("").reverse().join(""),")\u5426/\u662F(\u884C\u6267\u5426\u662F".split("").reverse().join(""),")\u5199\u586B\u4E0D\u53EF(\u79F0\u540D\u53F7\u8D26".split("").reverse().join("")]];var configContent=[["\u5de5\u4f5c\u8868\u7684\u540d\u79f0","\u6CE8\u5907".split("").reverse().join(""),"\u53ea\u63a8\u9001\u5931\u8d25\u6d88\u606f\uff08\u662f\u002f\u5426\uff09","\u63a8\u9001\u6635\u79f0\uff08\u662f\u002f\u5426\uff09"],[sheetNameSubConfig,pushHeader,"\u5426","\u662f"]];var qlpushFlag=0xed2aa^0xed2aa;var qlSheet=[];var colNum=["\u0041",'B','C',"\u0044","\u0045",'F',"\u0047","\u0048","\u0049",'J','K','L','M','N',"\u004f","\u0050",'Q'];qlConfig={'CONFIG':configContent,"\u0053\u0055\u0042\u0043\u004f\u004e\u0046\u0049\u0047":userContent};var posHttp=0xee242^0xee242;var flagFinish=0x31353^0x31353;var flagResultFinish=0xc6811^0xc6811;var HTTPOverwrite={'get':function get(_0x8a99d9,_0x350907){_0x350907=_0x350907['headers'];let _0x1f12ec=userContent['length']-qlpushFlag;method='get';resp=fetch(_0x8a99d9,{"\u006d\u0065\u0074\u0068\u006f\u0064":method,"\u0068\u0065\u0061\u0064\u0065\u0072\u0073":_0x350907})["\u0074\u0068\u0065\u006e"](function(_0x2c1fb0){return _0x2c1fb0["\u0074\u0065\u0078\u0074"]()['then'](_0xb09c64=>{return{'status':_0x2c1fb0["\u0073\u0074\u0061\u0074\u0075\u0073"],"\u0068\u0065\u0061\u0064\u0065\u0072\u0073":_0x2c1fb0['headers'],"\u0074\u0065\u0078\u0074":_0xb09c64,'response':_0x2c1fb0,"\u0070\u006f\u0073":_0x1f12ec};});})["\u0074\u0068\u0065\u006e"](function(_0x3872ad){try{data=JSON['parse'](_0x3872ad['text']);return{"\u0073\u0074\u0061\u0074\u0075\u0073":_0x3872ad["\u0073\u0074\u0061\u0074\u0075\u0073"],'headers':_0x3872ad['headers'],'json':function _0x352161(){return data;},"\u0074\u0065\u0078\u0074":function _0x29f35f(){return _0x3872ad["\u0074\u0065\u0078\u0074"];},"\u0070\u006f\u0073":_0x3872ad["\u0070\u006f\u0073"]};}catch(_0x566df2){return{'status':_0x3872ad['status'],"\u0068\u0065\u0061\u0064\u0065\u0072\u0073":_0x3872ad["\u0068\u0065\u0061\u0064\u0065\u0072\u0073"],"\u006a\u0073\u006f\u006e":null,'text':function _0x2fbf50(){return _0x3872ad['text'];},'pos':_0x3872ad["\u0070\u006f\u0073"]};}})["\u0074\u0068\u0065\u006e"](_0x2c1f58=>{_0x1f12ec=_0x2c1f58["\u0070\u006f\u0073"];flagResultFinish=resultHandle(_0x2c1f58,_0x1f12ec);if(flagResultFinish==(0xdbc9d^0xdbc9c)){i=_0x1f12ec+(0x573c0^0x573c1);for(;i<=line;i++){var _0x1cb220=Application['Range']('A'+i)['Text'];var _0x2267f8=Application['Range']('B'+i)["\u0054\u0065\u0078\u0074"];if(_0x1cb220=="".split("").reverse().join("")){break;}if(_0x2267f8=="\u662f"){console["\u006c\u006f\u0067"]('🧑\x20开始执行用户：'+(parseInt(i)-(0xeff03^0xeff02)));flagResultFinish=0xee64b^0xee64b;execHandle(_0x1cb220,i);break;}}}if(_0x1f12ec==userContent['length']&&flagResultFinish==(0x49187^0x49186)){flagFinish=0x80f74^0x80f75;}if(qlpushFlag==(0xae610^0xae610)&&flagFinish==0x1){console['log']("\u9001\u63A8\u8D77\u53D1\u9F99\u9752 \uDE80\uD83D".split("").reverse().join(""));message=messageMerge();const{sendNotify:_0x25bcc}=require("sj.yfitoNdnes/.".split("").reverse().join(""));_0x25bcc(pushHeader,message);qlpushFlag=-0x64;}})["\u0063\u0061\u0074\u0063\u0068"](_0x203eb4=>{console['error'](":rorre hcteF".split("").reverse().join(""),_0x203eb4);});},'post':function post(_0x13ae92,_0x5b8821,_0x44b685,_0x55317c){_0x44b685=_0x44b685['headers'];contentType=_0x44b685['Content-Type'];contentType2=_0x44b685["\u0063\u006f\u006e\u0074\u0065\u006e\u0074\u002d\u0074\u0079\u0070\u0065"];var _0x4db6b5="".split("").reverse().join("");if(contentType!=undefined&&contentType!="".split("").reverse().join("")||contentType2!=undefined&&contentType2!=''){if(contentType=="dedocnelru-mrof-www-x/noitacilppa".split("").reverse().join("")){console['log']('🍳\x20检测到发送请求体为:\x20表单格式');_0x4db6b5=dataToFormdata(_0x5b8821);}else{try{console["\u006c\u006f\u0067"]('🍳\x20检测到发送请求体为:\x20JSON格式');_0x4db6b5=JSON['stringify'](_0x5b8821);}catch{console['log']("\u5F0F\u683C\u5355\u8868 :\u4E3A\u4F53\u6C42\u8BF7\u9001\u53D1\u5230\u6D4B\u68C0 \uDF73\uD83C".split("").reverse().join(""));_0x4db6b5=_0x5b8821;}}}else{console['log']("\u5F0F\u683CNOSJ :\u4E3A\u4F53\u6C42\u8BF7\u9001\u53D1\u5230\u6D4B\u68C0 \uDF73\uD83C".split("").reverse().join(""));_0x4db6b5=JSON["\u0073\u0074\u0072\u0069\u006e\u0067\u0069\u0066\u0079"](_0x5b8821);}if(_0x55317c=="\u0067\u0065\u0074"||_0x55317c=="TEG".split("").reverse().join("")){let _0x326da0=userContent['length']-qlpushFlag;method='get';resp=fetch(_0x13ae92,{'method':method,"\u0068\u0065\u0061\u0064\u0065\u0072\u0073":_0x44b685})['then'](function(_0x32ad04){return _0x32ad04["\u0074\u0065\u0078\u0074"]()['then'](_0x4570f7=>{return{'status':_0x32ad04["\u0073\u0074\u0061\u0074\u0075\u0073"],'headers':_0x32ad04["\u0068\u0065\u0061\u0064\u0065\u0072\u0073"],'text':_0x4570f7,'response':_0x32ad04,'pos':_0x326da0};});})['then'](function(_0x533acc){try{_0x5b8821=JSON["\u0070\u0061\u0072\u0073\u0065"](_0x533acc["\u0074\u0065\u0078\u0074"]);return{"\u0073\u0074\u0061\u0074\u0075\u0073":_0x533acc["\u0073\u0074\u0061\u0074\u0075\u0073"],"\u0068\u0065\u0061\u0064\u0065\u0072\u0073":_0x533acc["\u0068\u0065\u0061\u0064\u0065\u0072\u0073"],'json':function _0x4c4d29(){return _0x5b8821;},'text':function _0x37f878(){return _0x533acc["\u0074\u0065\u0078\u0074"];},'pos':_0x533acc['pos']};}catch(_0x53a77f){return{"\u0073\u0074\u0061\u0074\u0075\u0073":_0x533acc['status'],"\u0068\u0065\u0061\u0064\u0065\u0072\u0073":_0x533acc["\u0068\u0065\u0061\u0064\u0065\u0072\u0073"],'json':null,'text':function _0xe9ffe5(){return _0x533acc["\u0074\u0065\u0078\u0074"];},'pos':_0x533acc["\u0070\u006f\u0073"]};}})['then'](_0x185f08=>{_0x326da0=_0x185f08['pos'];flagResultFinish=resultHandle(_0x185f08,_0x326da0);if(flagResultFinish==(0xc722f^0xc722e)){i=_0x326da0+(0x56040^0x56041);for(;i<=line;i++){var _0x5cc09e=Application["\u0052\u0061\u006e\u0067\u0065"]('A'+i)['Text'];var _0x5bca1a=Application['Range']("\u0042"+i)['Text'];if(_0x5cc09e==''){break;}if(_0x5bca1a=='是'){console['log']('🧑\x20开始执行用户：'+(parseInt(i)-0x1));flagResultFinish=0x1fd7e^0x1fd7e;execHandle(_0x5cc09e,i);break;}}}if(_0x326da0==userContent['length']&&flagResultFinish==(0xf2b2b^0xf2b2a)){flagFinish=0x1;}if(qlpushFlag==0x0&&flagFinish==(0x1cf48^0x1cf49)){console['log']('🚀\x20青龙发起推送');message=messageMerge();const{sendNotify:_0x38f3c8}=require("sj.yfitoNdnes/.".split("").reverse().join(""));_0x38f3c8(pushHeader,message);qlpushFlag=-(0xca032^0xca056);}})['catch'](_0x52f683=>{console['error']('Fetch\x20error:',_0x52f683);});}else{let _0xaeacf8=userContent['length']-qlpushFlag;method='post';resp=fetch(_0x13ae92,{'method':method,"\u0068\u0065\u0061\u0064\u0065\u0072\u0073":_0x44b685,'body':_0x4db6b5})['then'](function(_0x1c043a){return _0x1c043a['text']()['then'](_0x52ce21=>{return{'status':_0x1c043a['status'],'headers':_0x1c043a['headers'],'text':_0x52ce21,"\u0072\u0065\u0073\u0070\u006f\u006e\u0073\u0065":_0x1c043a,'pos':_0xaeacf8};});})['then'](function(_0x3ae307){try{_0x5b8821=JSON['parse'](_0x3ae307['text']);return{'status':_0x3ae307['status'],'headers':_0x3ae307['headers'],'json':function _0x5486bd(){return _0x5b8821;},'text':function _0x1d9320(){return _0x3ae307["\u0074\u0065\u0078\u0074"];},"\u0070\u006f\u0073":_0x3ae307['pos']};}catch(_0x2df6f6){return{'status':_0x3ae307['status'],'headers':_0x3ae307['headers'],'json':null,"\u0074\u0065\u0078\u0074":function _0x4bd139(){return _0x3ae307["\u0074\u0065\u0078\u0074"];},'pos':_0x3ae307['pos']};}})['then'](_0x55afac=>{_0xaeacf8=_0x55afac['pos'];flagResultFinish=resultHandle(_0x55afac,_0xaeacf8);if(flagResultFinish==(0x34da0^0x34da1)){i=_0xaeacf8+(0x95e50^0x95e51);for(;i<=line;i++){var _0x51de34=Application['Range']('A'+i)['Text'];var _0x529848=Application['Range']('B'+i)['Text'];if(_0x51de34=="".split("").reverse().join("")){break;}if(_0x529848=='是'){console['log']('🧑\x20开始执行用户：'+(parseInt(i)-0x1));flagResultFinish=0x0;execHandle(_0x51de34,i);break;}}}if(_0xaeacf8==userContent['length']&&flagResultFinish==(0xaa758^0xaa759)){flagFinish=0x1;}if(qlpushFlag==0x0&&flagFinish==(0x3ae30^0x3ae31)){console["\u006c\u006f\u0067"]('🚀\x20青龙发起推送');let _0x1c81b5=messageMerge();const{sendNotify:_0x542555}=require("sj.yfitoNdnes/.".split("").reverse().join(""));_0x542555(pushHeader,_0x1c81b5);qlpushFlag=-(0x85411^0x85475);}})['catch'](_0x3245cf=>{console["\u0065\u0072\u0072\u006f\u0072"](":rorre hcteF".split("").reverse().join(""),_0x3245cf);});}}};var ApplicationOverwrite={'Range':function Range(_0x17a08b){charFirst=_0x17a08b['substring'](0x0,0x1);qlRow=_0x17a08b['substring'](0xe77b9^0xe77b8,_0x17a08b['length']);qlCol=0x1;for(num in colNum){if(colNum[num]==charFirst){break;}qlCol+=0x1;}try{result=qlSheet[qlRow-0x1][qlCol-0x1];}catch{result='';}dict={'Text':result};return dict;},"\u0053\u0068\u0065\u0065\u0074\u0073":{'Item':function(_0x3f811c){return{'Name':_0x3f811c,'Activate':function(){flag=0x1;qlSheet=qlConfig[_0x3f811c];if(qlSheet==undefined){qlSheet=qlConfig['SUBCONFIG'];}console['log']("\uFF1A\u8868\u4F5C\u5DE5\u6D3B\u6FC0\u9F99\u9752 \uDF73\uD83C".split("").reverse().join("")+_0x3f811c);return flag;}};}}};var CryptoOverwrite={'createHash':function createHash(_0x2af55c){return{'update':function _0xaaf0ad(_0x2cad1b,_0x25d5bc){return{"\u0064\u0069\u0067\u0065\u0073\u0074":function _0x100f85(_0x45b9f3){return{'toUpperCase':function _0xe359a4(){return{'toString':function _0x37ba96(){try{CryptoJS=require('crypto-js');console['log']("\u5165\u5F15sj-otpyrc\u884C\u8FDB\u5DF2\u7EDF\u7CFB \uFE0F\u267B".split("").reverse().join(""));}catch{console['log']('❌\x20系统无crypto-js，请在NodeJs中安装crypto-js依赖');}md5Hash=CryptoJS['MD5'](_0x2cad1b)['toString']();md5Hash=md5Hash['toUpperCase']();return md5Hash;}};},'toString':function _0xfcf985(){try{CryptoJS=require('crypto-js');console['log']('♻️\x20系统已进行crypto-js引入');}catch{console['log']('❌\x20系统无crypto-js，请在NodeJs中安装crypto-js依赖');}md5Hash=CryptoJS['MD5'](_0x2cad1b)['toString']();return md5Hash;}};}};}};}};function dataToFormdata(_0x1073d4){result="";values=Object["\u0076\u0061\u006c\u0075\u0065\u0073"](_0x1073d4);values["\u0066\u006f\u0072\u0045\u0061\u0063\u0068"]((_0x46e3bd,_0xe0a588)=>{key=Object['keys'](_0x1073d4)[_0xe0a588];content=key+'='+_0x46e3bd+'&';result+=content;});result=result['substring'](0x0,result['length']-0x1);return result;}function cookiesTocookieMin(_0x3ec766){let _0x10c2b9=_0x3ec766;let _0x357077=[];var _0x527229=_0x10c2b9["\u0073\u0070\u006c\u0069\u0074"]('#');for(let _0x30e526 in _0x527229){_0x357077[_0x30e526]=_0x527229[_0x30e526];}return _0x357077;}function checkEscape(_0x24ae63,_0x2b0863){cookieArrynew=[];j=0x28920^0x28920;for(i=0x0;i<_0x24ae63['length'];i++){result=_0x24ae63[i];lastChar=result['substring'](result['length']-0x1,result['length']);if(lastChar=='\x5c'&&i<=_0x24ae63['length']-(0x5030a^0x50308)){console["\u006c\u006f\u0067"]('🍳\x20检测到转义字符');cookieArrynew[j]=result['substring'](0x0,result['length']-0x1)+_0x2b0863+_0x24ae63[parseInt(i)+(0xe77c7^0xe77c6)];i+=0x653af^0x653ae;}else{cookieArrynew[j]=_0x24ae63[i];}j+=0xccfd4^0xccfd5;}return cookieArrynew;}function cookiesTocookie(_0x30fdb1){let _0x7eefa3=_0x30fdb1;let _0x5e4f35=[];let _0x3e1587=[];_0x7eefa3=_0x7eefa3['trim']();let _0x105de4=_0x7eefa3['split']('\x0a');_0x105de4=_0x105de4["\u0066\u0069\u006c\u0074\u0065\u0072"](_0x5eac64=>_0x5eac64['trim']()!=="");if(_0x105de4['length']==(0x1e7cd^0x1e7cc)){_0x105de4=_0x7eefa3['split']('@');_0x105de4=checkEscape(_0x105de4,'@');}for(let _0x171b55 in _0x105de4){_0x3e1587=[];let _0x401143=Number(_0x171b55)+0x1;_0x5e4f35=cookiesTocookieMin(_0x105de4[_0x171b55]);_0x5e4f35=checkEscape(_0x5e4f35,'#');_0x3e1587['push'](_0x5e4f35[0x0]);_0x3e1587['push']('是');_0x3e1587['push']("\u79F0\u6635".split("").reverse().join("")+_0x401143);if(_0x5e4f35['length']>0x0){for(let _0xd30881=0x3;_0xd30881<_0x5e4f35['length']+(0x4721c^0x4721e);_0xd30881++){_0x3e1587['push'](_0x5e4f35[_0xd30881-(0xdde77^0xdde75)]);}}userContent['push'](_0x3e1587);}qlpushFlag=userContent['length']-0x1;}var qlSwitch=0x0;try{qlSwitch=process['env'][sheetNameSubConfig];qlSwitch=0x1;}catch{qlSwitch=0x2ad67^0x2ad67;console['log']('♻️\x20当前环境为金山文档');console['log']('♻️\x20开始适配金山文档，执行金山文档代码');}if(qlSwitch){console['log']('♻️\x20当前环境为青龙');console['log']('♻️\x20开始适配青龙环境，执行青龙代码');try{fetch=require('node-fetch');console['log']("\u5165\u5F15hctef-edon\u884C\u8FDB\u5DF2\uFF0Chctef\u65E0\u7EDF\u7CFB \uFE0F\u267B".split("").reverse().join(""));}catch{console['log']('♻️\x20系统已有原生fetch');}Crypto=CryptoOverwrite;let flagwarn=0xaf916^0xaf916;const a='da11990c';const b="0b854f216a9662fb".split("").reverse().join("");encode=getsign(logo);let len=encode['length'];if(a+"90ecd4ce".split("").reverse().join("")==encode['substring'](0xf30ba^0xf30ba,len/0x2)&&b==encode['substring']((0x2bc08^0x2bc0c)*(0xc6930^0xc6934),len)){console["\u006c\u006f\u0067"]('✨\x20'+logo);cookies=process['env'][sheetNameSubConfig];}else{console["\u006c\u006f\u0067"]('🔨\x20请使用艾默库代码\x20:\x20https://github.com/imoki/sign_script');flagwarn=0x1;}let flagwarn2=0x2038e^0x2038f;const welcome="edoc UKOM esu ot emocleW".split("").reverse().join("");const mo=welcome['slice'](0xf,0xd9c98^0xd9c89)['toLowerCase']();const ku=welcome['split']('\x20')[0x4-(0x3dbee^0x3dbef)]['slice'](0x2,0xe25d3^0xe25d7);if(mo['substring'](0x0,0x1)=='m'){if(ku=='KU'){if(mo['substring'](0x1,0x2)==String['fromCharCode'](0x93cd7^0x93cb8)){cookiesTocookie(cookies);flagwarn2=0xb4736^0xb4736;console['log']('💗\x20'+welcome);}}}let t=Date['now']();if(t>(0x52714^0x527be)*0x186a0*0x186a0+0x45f34a08e){console['log']('🧾\x20使用教程请查看仓库notion链接');Application=ApplicationOverwrite;}else{flagwarn=0x5461f^0x5461e;}if(Date['now']()<(0x4e17c^0x4e1b4)*0x186a0*0x186a0){console['log']('🤝\x20欢迎各种形式的贡献');HTTP=HTTPOverwrite;}else{flagwarn2=0x1;}if(flagwarn==0x1||flagwarn2==0x1){console['log']('🔨\x20请使用艾默库代码\x20:\x20https://github.com/imoki/sign_script');}}
}catch{
  console.log("❌ 环境存在问题，请检查是否配置好了对应依赖及环境变量")
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
      if(version == 1){
        // 找到指定的表行
        if(Application.Range("A" + (i + 2)).Value == sheetNameSubConfig){
          // 写入更新的时间
          Application.Range("F" + (i + 2)).Value = todayDate
          // 写入消息
          Application.Range("G" + (i + 2)).Value = message
          console.log("✨ 写入结果完成");
          break;
        }
      }else{
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
// main()  // 入口

// function main(){
  checkVesion() // 版本检测，以进行不同版本的适配

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

// }

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

// cookie字符串转json格式
function cookie_to_json(cookies) {
  var cookie_text = cookies;
  var arr = [];
  var text_to_split = cookie_text.split(";");
  for (var i in text_to_split) {
    var tmp = text_to_split[i].split("=");
    arr.push('"' + tmp.shift().trim() + '":"' + tmp.join(":").trim() + '"');
  }
  var res = "{\n" + arr.join(",\n") + "\n}";
  return JSON.parse(res);
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
            // let url2 = "https://drive-m.quark.cn/1/clouddrive/capacity/growth/sign?pr=ucpro&fr=pc&uc_param_str="; // 进行签到
            let url2 = "https://drive-m.quark.cn/1/clouddrive/capacity/growth/sign?" + cookie
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
            // url =  "https://drive-m.quark.cn/1/clouddrive/capacity/growth/info?pr=ucpro&fr=pc&uc_param_str=";
            url = "https://drive-m.quark.cn/1/clouddrive/capacity/growth/info?" + cookie
            if(qlSwitch != 1){  // 金山文档
                // resp = HTTP.fetch(url, {
                //     method: "get",
                //     headers: headers,
                //     // data: data
                // });
                resp = HTTP.get(url, {headers: headers,});
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

    // console.log(cookie)

    params = cookie.split("?")
    cookie = ""
    for(let i= 1; i<params.length; i++){
      cookie +=params[i]
    }
    // console.log(param)
    
  // try {
    // let url1 = "https://drive-m.quark.cn/1/clouddrive/capacity/growth/info?pr=ucpro&fr=pc&uc_param_str="; // 查询是否签到
    let url = "https://drive-m.quark.cn/1/clouddrive/capacity/growth/info?" + cookie
    
    headers = {
      // "Cookie": cookie,
      "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/46.0.2486.0 Safari/537.36 Edge/13.10586"
    };

    if(qlSwitch != 1){  // 金山文档
        // resp = HTTP.fetch(url, {
        //     method: "get",
        //     headers: headers,
        //     // data: data
        // });
        // resp = HTTP.get(url, headers)
        resp = HTTP.get(url, {headers: headers,});
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
