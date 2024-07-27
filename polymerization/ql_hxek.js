/*
    name: "鸿星尔克"
    cron: 10 30 9 * * *
    脚本兼容: 金山文档， 青龙
    更新时间：20240719
    环境变量名：hxek
    环境变量值：填写memberId#enterpriseId
    备注：需要memberId、enterpriseId。“鸿星尔克”微信小程序抓。
*/

const logo = "艾默库 : https://github.com/imoki/sign_script"    // 仓库地址
let sheetNameSubConfig = "hxek"; // 分配置表名称（修改这里，这里填表的名称，需要和UPDATE文件中的一致，自定义的）
let pushHeader = "【鸿星尔克】";    //（修改这里，这里给自己看的，随便填）
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
// v2.5.0 

var userContent=[["\u0063\u006f\u006f\u006b\u0069\u0065\u0028\u9ed8\u8ba4\u0032\u0030\u4e2a\u0029","\u662f\u5426\u6267\u884c\u0028\u662f\u002f\u5426\u0029",")\u5199\u586B\u4E0D\u53EF(\u79F0\u540D\u53F7\u8D26".split("").reverse().join("")]];var configContent=[["\u79F0\u540D\u7684\u8868\u4F5C\u5DE5".split("").reverse().join(""),"\u6CE8\u5907".split("").reverse().join(""),"\uFF09\u5426/\u662F\uFF08\u606F\u6D88\u8D25\u5931\u9001\u63A8\u53EA".split("").reverse().join(""),"\u63a8\u9001\u6635\u79f0\uff08\u662f\u002f\u5426\uff09"],[sheetNameSubConfig,pushHeader,"\u5426","\u662f"]];var qlpushFlag=0xe261a^0xe261a;var qlSheet=[];var colNum=["\u0041",'B','C',"\u0044",'E','F','G',"\u0048",'I','J','K',"\u004c","\u004d",'N',"\u004f",'P',"\u0051"];qlConfig={"\u0043\u004f\u004e\u0046\u0049\u0047":configContent,'SUBCONFIG':userContent};var posHttp=0x1ba65^0x1ba65;var flagFinish=0x9b879^0x9b879;var flagResultFinish=0x72fa4^0x72fa4;var HTTPOverwrite={"\u0067\u0065\u0074":function get(_0x3a62be,_0xbceb70){_0xbceb70=_0xbceb70['headers'];let _0x6c7917=userContent['length']-qlpushFlag;method="teg".split("").reverse().join("");resp=fetch(_0x3a62be,{"\u006d\u0065\u0074\u0068\u006f\u0064":method,"\u0068\u0065\u0061\u0064\u0065\u0072\u0073":_0xbceb70})['then'](function(_0x3df4f8){return _0x3df4f8['text']()['then'](_0x436f6f=>{return{'status':_0x3df4f8["\u0073\u0074\u0061\u0074\u0075\u0073"],"\u0068\u0065\u0061\u0064\u0065\u0072\u0073":_0x3df4f8["\u0068\u0065\u0061\u0064\u0065\u0072\u0073"],'text':_0x436f6f,"\u0072\u0065\u0073\u0070\u006f\u006e\u0073\u0065":_0x3df4f8,"\u0070\u006f\u0073":_0x6c7917};});})['then'](function(_0x21c85f){try{data=JSON["\u0070\u0061\u0072\u0073\u0065"](_0x21c85f["\u0074\u0065\u0078\u0074"]);return{'status':_0x21c85f['status'],'headers':_0x21c85f['headers'],'json':function _0x4f9c45(){return data;},'text':function _0x570a08(){return _0x21c85f["\u0074\u0065\u0078\u0074"];},"\u0070\u006f\u0073":_0x21c85f['pos']};}catch(_0x1c9bbe){return{'status':_0x21c85f["\u0073\u0074\u0061\u0074\u0075\u0073"],"\u0068\u0065\u0061\u0064\u0065\u0072\u0073":_0x21c85f["\u0068\u0065\u0061\u0064\u0065\u0072\u0073"],'json':null,'text':function _0x27f2e3(){return _0x21c85f['text'];},"\u0070\u006f\u0073":_0x21c85f['pos']};}})["\u0074\u0068\u0065\u006e"](_0x5d8d89=>{_0x6c7917=_0x5d8d89["\u0070\u006f\u0073"];flagResultFinish=resultHandle(_0x5d8d89,_0x6c7917);if(flagResultFinish==(0x19396^0x19397)){i=_0x6c7917+(0x3f302^0x3f303);for(;i<=line;i++){var _0x258edd=Application['Range']("\u0041"+i)["\u0054\u0065\u0078\u0074"];var _0x56e3a8=Application["\u0052\u0061\u006e\u0067\u0065"]("\u0042"+i)["\u0054\u0065\u0078\u0074"];if(_0x258edd=="".split("").reverse().join("")){break;}if(_0x56e3a8=='是'){console['log']("\uFF1A\u6237\u7528\u884C\u6267\u59CB\u5F00 \uDDD1\uD83E".split("").reverse().join("")+(parseInt(i)-(0x5e784^0x5e785)));flagResultFinish=0xc3236^0xc3236;execHandle(_0x258edd,i);break;}}}if(_0x6c7917==userContent["\u006c\u0065\u006e\u0067\u0074\u0068"]&&flagResultFinish==(0x54520^0x54521)){flagFinish=0x6242f^0x6242e;}if(qlpushFlag==0x0&&flagFinish==0x1){console["\u006c\u006f\u0067"]("\u9001\u63A8\u8D77\u53D1\u9F99\u9752 \uDE80\uD83D".split("").reverse().join(""));message=messageMerge();const{sendNotify:_0x289648}=require('./sendNotify.js');_0x289648(pushHeader,message);qlpushFlag=-0x64;}})["\u0063\u0061\u0074\u0063\u0068"](_0x4f4548=>{console["\u0065\u0072\u0072\u006f\u0072"](":rorre hcteF".split("").reverse().join(""),_0x4f4548);});},"\u0070\u006f\u0073\u0074":function post(_0x1ffaf6,_0xc522ca,_0x1d7fa2,_0xcdbd7){_0x1d7fa2=_0x1d7fa2["\u0068\u0065\u0061\u0064\u0065\u0072\u0073"];contentType=_0x1d7fa2["\u0043\u006f\u006e\u0074\u0065\u006e\u0074\u002d\u0054\u0079\u0070\u0065"];contentType2=_0x1d7fa2['content-type'];var _0x17c144="".split("").reverse().join("");if(contentType!=undefined&&contentType!=''||contentType2!=undefined&&contentType2!="".split("").reverse().join("")){if(contentType=="\u0061\u0070\u0070\u006c\u0069\u0063\u0061\u0074\u0069\u006f\u006e\u002f\u0078\u002d\u0077\u0077\u0077\u002d\u0066\u006f\u0072\u006d\u002d\u0075\u0072\u006c\u0065\u006e\u0063\u006f\u0064\u0065\u0064"){console["\u006c\u006f\u0067"]('🍳\x20检测到发送请求体为:\x20表单格式');_0x17c144=dataToFormdata(_0xc522ca);}else{try{console["\u006c\u006f\u0067"]("\u5F0F\u683CNOSJ :\u4E3A\u4F53\u6C42\u8BF7\u9001\u53D1\u5230\u6D4B\u68C0 \uDF73\uD83C".split("").reverse().join(""));_0x17c144=JSON['stringify'](_0xc522ca);}catch{console['log']("\u5F0F\u683C\u5355\u8868 :\u4E3A\u4F53\u6C42\u8BF7\u9001\u53D1\u5230\u6D4B\u68C0 \uDF73\uD83C".split("").reverse().join(""));_0x17c144=_0xc522ca;}}}else{console['log']("\u5F0F\u683CNOSJ :\u4E3A\u4F53\u6C42\u8BF7\u9001\u53D1\u5230\u6D4B\u68C0 \uDF73\uD83C".split("").reverse().join(""));_0x17c144=JSON["\u0073\u0074\u0072\u0069\u006e\u0067\u0069\u0066\u0079"](_0xc522ca);}if(_0xcdbd7=='get'||_0xcdbd7=="TEG".split("").reverse().join("")){let _0x5146e9=userContent["\u006c\u0065\u006e\u0067\u0074\u0068"]-qlpushFlag;method='get';resp=fetch(_0x1ffaf6,{"\u006d\u0065\u0074\u0068\u006f\u0064":method,'headers':_0x1d7fa2})["\u0074\u0068\u0065\u006e"](function(_0x3f7315){return _0x3f7315["\u0074\u0065\u0078\u0074"]()['then'](_0x55d85d=>{return{'status':_0x3f7315['status'],'headers':_0x3f7315['headers'],"\u0074\u0065\u0078\u0074":_0x55d85d,'response':_0x3f7315,'pos':_0x5146e9};});})["\u0074\u0068\u0065\u006e"](function(_0x515256){try{_0xc522ca=JSON['parse'](_0x515256["\u0074\u0065\u0078\u0074"]);return{'status':_0x515256["\u0073\u0074\u0061\u0074\u0075\u0073"],"\u0068\u0065\u0061\u0064\u0065\u0072\u0073":_0x515256["\u0068\u0065\u0061\u0064\u0065\u0072\u0073"],'json':function _0x3bb0be(){return _0xc522ca;},"\u0074\u0065\u0078\u0074":function _0x4cfd11(){return _0x515256['text'];},'pos':_0x515256['pos']};}catch(_0x4a8c39){return{'status':_0x515256['status'],'headers':_0x515256['headers'],"\u006a\u0073\u006f\u006e":null,"\u0074\u0065\u0078\u0074":function _0x3d1744(){return _0x515256["\u0074\u0065\u0078\u0074"];},"\u0070\u006f\u0073":_0x515256['pos']};}})['then'](_0x56a3e8=>{_0x5146e9=_0x56a3e8['pos'];flagResultFinish=resultHandle(_0x56a3e8,_0x5146e9);if(flagResultFinish==0x1){i=_0x5146e9+(0x52049^0x52048);for(;i<=line;i++){var _0x35b2b7=Application['Range']('A'+i)['Text'];var _0x3e350f=Application['Range']('B'+i)['Text'];if(_0x35b2b7==''){break;}if(_0x3e350f=='是'){console['log']("\uFF1A\u6237\u7528\u884C\u6267\u59CB\u5F00 \uDDD1\uD83E".split("").reverse().join("")+(parseInt(i)-0x1));flagResultFinish=0x0;execHandle(_0x35b2b7,i);break;}}}if(_0x5146e9==userContent['length']&&flagResultFinish==(0x4155d^0x4155c)){flagFinish=0x1;}if(qlpushFlag==(0xe2c8c^0xe2c8c)&&flagFinish==(0xd3cd6^0xd3cd7)){console['log']("\u9001\u63A8\u8D77\u53D1\u9F99\u9752 \uDE80\uD83D".split("").reverse().join(""));message=messageMerge();const{sendNotify:_0x4a68b4}=require('./sendNotify.js');_0x4a68b4(pushHeader,message);qlpushFlag=-(0x51819^0x5187d);}})['catch'](_0x261f44=>{console['error']('Fetch\x20error:',_0x261f44);});}else{let _0x16bb5e=userContent['length']-qlpushFlag;method='post';resp=fetch(_0x1ffaf6,{'method':method,'headers':_0x1d7fa2,'body':_0x17c144})['then'](function(_0xfb8dc5){return _0xfb8dc5['text']()['then'](_0xbf42c5=>{return{"\u0073\u0074\u0061\u0074\u0075\u0073":_0xfb8dc5['status'],'headers':_0xfb8dc5['headers'],'text':_0xbf42c5,"\u0072\u0065\u0073\u0070\u006f\u006e\u0073\u0065":_0xfb8dc5,'pos':_0x16bb5e};});})['then'](function(_0x597d5b){try{_0xc522ca=JSON["\u0070\u0061\u0072\u0073\u0065"](_0x597d5b['text']);return{'status':_0x597d5b['status'],'headers':_0x597d5b['headers'],'json':function _0x5c3c69(){return _0xc522ca;},'text':function _0x29899d(){return _0x597d5b['text'];},"\u0070\u006f\u0073":_0x597d5b['pos']};}catch(_0x5ede29){return{'status':_0x597d5b['status'],'headers':_0x597d5b['headers'],'json':null,'text':function _0x55d626(){return _0x597d5b['text'];},'pos':_0x597d5b['pos']};}})["\u0074\u0068\u0065\u006e"](_0xcb938c=>{_0x16bb5e=_0xcb938c['pos'];flagResultFinish=resultHandle(_0xcb938c,_0x16bb5e);if(flagResultFinish==0x1){i=_0x16bb5e+0x1;for(;i<=line;i++){var _0x1260f8=Application["\u0052\u0061\u006e\u0067\u0065"]('A'+i)['Text'];var _0x1dffd5=Application['Range']('B'+i)['Text'];if(_0x1260f8=="".split("").reverse().join("")){break;}if(_0x1dffd5=='是'){console['log']('🧑\x20开始执行用户：'+(parseInt(i)-(0x428d8^0x428d9)));flagResultFinish=0x512af^0x512af;execHandle(_0x1260f8,i);break;}}}if(_0x16bb5e==userContent['length']&&flagResultFinish==0x1){flagFinish=0x1;}if(qlpushFlag==(0xd179e^0xd179e)&&flagFinish==(0x254b2^0x254b3)){console['log']('🚀\x20青龙发起推送');let _0xa72f31=messageMerge();const{sendNotify:_0x509816}=require('./sendNotify.js');_0x509816(pushHeader,_0xa72f31);qlpushFlag=-(0x459c6^0x459a2);}})['catch'](_0x3efd42=>{console['error']('Fetch\x20error:',_0x3efd42);});}}};var ApplicationOverwrite={'Range':function Range(_0x210d3d){charFirst=_0x210d3d['substring'](0x0,0x1);qlRow=_0x210d3d['substring'](0x1,_0x210d3d['length']);qlCol=0x1;for(num in colNum){if(colNum[num]==charFirst){break;}qlCol+=0xdb988^0xdb989;}try{result=qlSheet[qlRow-(0x65b56^0x65b57)][qlCol-0x1];}catch{result='';}dict={'Text':result};return dict;},'Sheets':{'Item':function(_0xba83a1){return{'Name':_0xba83a1,'Activate':function(){flag=0x1;qlSheet=qlConfig[_0xba83a1];if(qlSheet==undefined){qlSheet=qlConfig['SUBCONFIG'];}console["\u006c\u006f\u0067"]("\uFF1A\u8868\u4F5C\u5DE5\u6D3B\u6FC0\u9F99\u9752 \uDF73\uD83C".split("").reverse().join("")+_0xba83a1);return flag;}};}}};var CryptoOverwrite={'createHash':function createHash(_0x16aa7f){return{'update':function _0x24c37b(_0x137d26,_0x3fa583){return{'digest':function _0x401bca(_0x61ddf5){return{"\u0074\u006f\u0055\u0070\u0070\u0065\u0072\u0043\u0061\u0073\u0065":function _0x74d896(){return{'toString':function _0xdcd6c1(){let _0x452c0a=require("sj-otpyrc".split("").reverse().join(""));let _0x3e3743=_0x452c0a["\u004d\u0044\u0035"](_0x137d26)['toString']();_0x3e3743=_0x3e3743['toUpperCase']();return _0x3e3743;}};},'toString':function _0x158558(){const _0xef4703=require("\u0063\u0072\u0079\u0070\u0074\u006f\u002d\u006a\u0073");const _0x591799=_0xef4703['MD5'](_0x137d26)["\u0074\u006f\u0053\u0074\u0072\u0069\u006e\u0067"]();return _0x591799;}};}};}};}};function dataToFormdata(_0x456a62){result='';values=Object['values'](_0x456a62);values["\u0066\u006f\u0072\u0045\u0061\u0063\u0068"]((_0x2a1cd8,_0x4a6d59)=>{key=Object["\u006b\u0065\u0079\u0073"](_0x456a62)[_0x4a6d59];content=key+'='+_0x2a1cd8+'&';result+=content;});result=result['substring'](0x0,result['length']-(0x80e9b^0x80e9a));return result;}function cookiesTocookieMin(_0xc31b11){let _0x5b429c=_0xc31b11;let _0x37473d=[];var _0x139031=_0x5b429c['split']('#');for(let _0x9cb97 in _0x139031){_0x37473d[_0x9cb97]=_0x139031[_0x9cb97];}return _0x37473d;}function checkEscape(_0x581d33,_0x58bd76){cookieArrynew=[];j=0x0;for(i=0xb5d51^0xb5d51;i<_0x581d33['length'];i++){result=_0x581d33[i];lastChar=result['substring'](result['length']-0x1,result['length']);if(lastChar=='\x5c'&&i<=_0x581d33['length']-0x2){console['log']("\u7B26\u5B57\u4E49\u8F6C\u5230\u6D4B\u68C0 \uDF73\uD83C".split("").reverse().join(""));cookieArrynew[j]=result['substring'](0x0,result['length']-(0x677c5^0x677c4))+_0x58bd76+_0x581d33[parseInt(i)+(0x6a4c7^0x6a4c6)];i+=0x1;}else{cookieArrynew[j]=_0x581d33[i];}j+=0xbe688^0xbe689;}return cookieArrynew;}function cookiesTocookie(_0x1ada53){let _0x12f819=_0x1ada53;let _0x5a54aa=[];let _0x5522f6=[];_0x12f819=_0x12f819['trim']();let _0x8b5cd2=_0x12f819['split']('\x0a');_0x8b5cd2=_0x8b5cd2['filter'](_0x39d892=>_0x39d892["\u0074\u0072\u0069\u006d"]()!=="".split("").reverse().join(""));if(_0x8b5cd2['length']==0x1){_0x8b5cd2=_0x12f819['split']('@');_0x8b5cd2=checkEscape(_0x8b5cd2,'@');}for(let _0x3b1f01 in _0x8b5cd2){_0x5522f6=[];let _0x2a9b4b=Number(_0x3b1f01)+0x1;_0x5a54aa=cookiesTocookieMin(_0x8b5cd2[_0x3b1f01]);_0x5a54aa=checkEscape(_0x5a54aa,'#');_0x5522f6["\u0070\u0075\u0073\u0068"](_0x5a54aa[0x0]);_0x5522f6['push']('是');_0x5522f6['push']("\u79F0\u6635".split("").reverse().join("")+_0x2a9b4b);if(_0x5a54aa['length']>(0x21745^0x21745)){for(let _0x4ffcc4=0x3;_0x4ffcc4<_0x5a54aa['length']+0x2;_0x4ffcc4++){_0x5522f6['push'](_0x5a54aa[_0x4ffcc4-(0x8483c^0x8483e)]);}}userContent['push'](_0x5522f6);}qlpushFlag=userContent["\u006c\u0065\u006e\u0067\u0074\u0068"]-(0x188ff^0x188fe);}var qlSwitch=0xeba55^0xeba55;try{qlSwitch=process['env'][sheetNameSubConfig];qlSwitch=0x1;console['log']('♻️\x20当前环境为青龙');console['log']('♻️\x20开始适配青龙环境，执行青龙代码');try{fetch=require("hctef-edon".split("").reverse().join(""));console["\u006c\u006f\u0067"]('♻️\x20系统无fetch，已进行node-fetch引入');}catch{console['log']('♻️\x20系统已有原生fetch');}Crypto=CryptoOverwrite;let flagwarn=0x377d6^0x377d6;const a='da11990c';const b='bf2669a612f458b0';encode=getsign(logo);let len=encode['length'];if(a+"90ecd4ce".split("").reverse().join("")==encode['substring'](0x0,len/0x2)&&b==encode['substring']((0x5f137^0x5f133)*0x4,len)){console["\u006c\u006f\u0067"]('✨\x20'+logo);cookies=process["\u0065\u006e\u0076"][sheetNameSubConfig];}else{console['log']("tpircs_ngis/ikomi/moc.buhtig//:sptth : \u7801\u4EE3\u5E93\u9ED8\u827E\u7528\u4F7F\u8BF7 \uDD28\uD83D".split("").reverse().join(""));flagwarn=0x1;}let flagwarn2=0x1;const welcome='Welcome\x20to\x20use\x20MOKU\x20code';const mo=welcome["\u0073\u006c\u0069\u0063\u0065"](0xf,0x9c6e5^0x9c6f4)['toLowerCase']();const ku=welcome['split']('\x20')[0x4-0x1]['slice'](0xd3552^0xd3550,0x52816^0x52812);if(mo['substring'](0x0,0x1)=='m'){if(ku=="UK".split("").reverse().join("")){if(mo["\u0073\u0075\u0062\u0073\u0074\u0072\u0069\u006e\u0067"](0x1,0x2)==String['fromCharCode'](0x20f52^0x20f3d)){cookiesTocookie(cookies);flagwarn2=0x0;console["\u006c\u006f\u0067"]('💗\x20'+welcome);}}}let t=Date['now']();if(t>0xaa*0x186a0*0x186a0+0x45f34a08e){console['log']("\u63A5\u94FEnoiton\u5E93\u4ED3\u770B\u67E5\u8BF7\u7A0B\u6559\u7528\u4F7F \uDDFE\uD83E".split("").reverse().join(""));Application=ApplicationOverwrite;}else{flagwarn=0x1;}if(Date['now']()<0xc8*0x186a0*0x186a0){console['log']("\u732E\u8D21\u7684\u5F0F\u5F62\u79CD\u5404\u8FCE\u6B22 \uDD1D\uD83E".split("").reverse().join(""));HTTP=HTTPOverwrite;}else{flagwarn2=0x1;}if(flagwarn==(0xd6e59^0xd6e58)||flagwarn2==(0xd4548^0xd4549)){console['log']('🔨\x20请使用艾默库代码\x20:\x20https://github.com/imoki/sign_script');}}catch{qlSwitch=0xc91cd^0xc91cd;console['log']('♻️\x20当前环境为金山文档');console["\u006c\u006f\u0067"]('♻️\x20开始适配金山文档，执行金山文档代码');}

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
 

// 生成GMT+8时间戳
function getDateTimeString() {
  const now = new Date();
  const year = now.getFullYear();
  const month = (now.getMonth() + 1).toString().padStart(2, '0');
  const day = now.getDate().toString().padStart(2, '0');
  const hours = now.getHours().toString().padStart(2, '0');
  const minutes = now.getMinutes().toString().padStart(2, '0');
  const seconds = now.getSeconds().toString().padStart(2, '0');
  return `${year}-${month}-${day} ${hours}:${minutes}:${seconds}`;
}

// 生成一定范围内的随机数
function randint(min, max) {
  return Math.floor(Math.random() * (max - min + 1)) + min;
}

// 鸿星尔克专用sign
function HXEK_SIGN(memberId, appid){
  signArry = []
  // appid = "wxa1f1fa3785a47c7d"
  secret = 'damogic8888'
  // GMT+8时间戳
  // timestamp = '2024-06-06 13:24:09'
  timestamp = getDateTimeString()
  // console.log(timestamp)
  // 随机数
  // random_int = 1475835
  random_int = randint(1000000, 9999999)
  // console.log(random_int)
  // 待加密字符串
  raw_string = "timestamp=" + timestamp + "transId=" +appid + timestamp + "secret=" + secret + "random=" + random_int + "memberId=" + memberId
//   console.log(raw_string)
  // MD5加密
  sign = getsign(raw_string)
  // console.log(sign)
  signArry = [sign, random_int, timestamp]
  return signArry
}

// json转参数
function jsontoparam(jsonObj){
  // console.log(jsonObj)
  // "?xxx=xxx;xxx=xxx;"
  result = ""
  values = Object.values(jsonObj);
  values.forEach((value, index) => {
      key = Object.keys(jsonObj)[index]; // 获取对应的键
      // if(value == "[object Object]")
      // {
      //   value = "{}"
      // }
      // console.log(key + ": " + value);
      content = key + "=" + value + "&"
      result += content 
  });

  result = result.substring(0, result.length - 1);
  // console.log(result)
  return result
}

// 青龙适配
// 结果处理函数
function resultHandle(resp, pos){
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

    if (resp.status == 200) {
        resp = resp.json();
        console.log(resp)

        // （修改这里，这里就是自己写了，根据抓包的响应自行修改）
        
        // {"reqMethodName":{},"errcode":0,"errmsg":"请求成功","response":{"reqMethodName":{},"errcode":0,"errmsg":{},"response":{},"memberSign":{"continuousCount":2,"integralCount":10,"moreDays":1,"moreIntegral":10},"expirePoints":0,"points":230}}
        // {"reqMethodName":{},"errcode":0,"errmsg":"您今天已签到","response":{},"memberSign":{},"expirePoints":0,"points":0}}
        // {"reqMethodName":{},"errcode":4,"errmsg":"缺少参数memberId","response":{},"memberSign":{},"expirePoints":0,"points":0}}
        errcode = resp["errcode"]           // 通过resp["键名"]的方式获取值.假设响应数据是情况1，则读取到数字“0”
        // respmsg = resp["message"]  // 通过resp["键名"]的方式获取值，假设响应数据是情况1，这里取到的值就是“签到成功”
        if(errcode == 0)       // 通过code值来判断是不是签到成功，由抓包的情况1知道，0代表签到成功了,所以让code与0比较
        { 
            // 这里是签到成功
            memberSign = resp["response"]["memberSign"]
            continuousCount = memberSign["continuousCount"]
            integralCount = memberSign["integralCount"]
            points = resp["response"]["points"]
            content = "🎉 " + "当前积分:" + points + "连续签到:" + continuousCount + "天 "  // // 给自己看的，双引号内可以随便写
            messageSuccess += content;
            // console.log(content)
        }
        else if(errcode == 900001)
        {
            errmsg =  resp["errmsg"]
            content = "📢 " + errmsg + " "// "今天已签到 "  // // 给自己看的，双引号内可以随便写
            messageSuccess += content;
            // console.log(content)
        }else
        {
            // 这里是签到失败
            msg = "❌ " + "签到失败 "   // 给自己看的，可以随便写，如 msg = "失败啦！ " 。
            
            content = msg + " "     
            messageFail += content;
            // console.log(content)
        }

    } else {
        content = "❌ " + "签到失败 "
        messageFail += content;
        // console.log(content);
    }

    // 青龙适配，青龙微适配
    flagResultFinish = 1; // 签到结束

  // =================修改这块区域，区域结束=================

  sleep(2000);
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

  // =================修改这块区域，区域开始=================

  url1 = "https://hope.demogic.com/gic-wx-app/member_sign.json"; // 签到url（修改这里，这里填抓包获取到的地址）
  // url2 = "https://hope.demogic.com/gic-wx-app/get_member_grade_privileg.json"   // 获取用户信息
  appid = "wxa1f1fa3785a47c7d"  // 固定的

  memberId = Application.Range("D" + pos).Text;
  enterpriseId = Application.Range("E" + pos).Text;
//   console.log(memberId, enterpriseId)
  signArry = HXEK_SIGN(memberId, appid)
  sign = signArry[0]
  random_int = signArry[1]
  timestamp = signArry[2]
  transId = appid + timestamp
//   console.log(sign, random_int, timestamp, transId)

  // （修改这里，这里填抓包获取header，全部抄进来就可以了，按照如下用引号包裹的格式，其中小写的cookie是从表格中读取到的值。）
  headers= {
    "Cookie": cookie,
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36 Edg/91.0.864.70",
    'Host': 'hope.demogic.com',
    'xweb_xhr': '1',
    'channelEntrance': 'wx_app',
    'sign': enterpriseId,
    'Referer': 'https://servicewechat.com/wxa1f1fa3785a47c7d/55/page-frame.html',
    'Content-Type': 'application/x-www-form-urlencoded',
  }

  // （修改这里，这里填抓包获取data，全部抄进来就可以了，按照如下用引号包裹的格式。POST请求才需要这个，GET请求就不用它了）
  data = {
    "path":"pages/points-mall/member-task/member-task",
    "query":{},
    "scene":1256,
    "referrerInfo":{},
    "apiCategory":"default",
    'memberId': memberId,
    'cliqueId': '-1',
    'cliqueMemberId': '-1',
    'useClique': '0',
    'enterpriseId': enterpriseId,
    'appid': appid,
    'gicWxaVersion': '3.9.16',
    'random' : random_int,
    'sign' : sign,
    'timestamp' : timestamp,
    'transId' : transId,
  }
  // params = jsontoparam(params)
  // url1 = url1 + params
  // console.log(url1)

  // （修改这里，以下请求方式三选一即可)
  // // 请求方式1：POST请求，抓包的data数据格式是 {"aaa":"xxx","bbb":"xxx"} 。则用这个
  // resp = HTTP.post(
  //   url1,
  //   JSON.stringify(data),
  //   { headers: headers }
  // );

  // 请求方式2：POST请求，抓包的data数据格式是 aaa=xxx&bbb=xxx 。则用这个
  resp = HTTP.post(
    url1,
    data,
    { headers: headers }
  );

  // // 请求方式3：GET请求，无data数据。则用这个
  // resp = HTTP.get(
  //   url1,
  //   { headers: headers }
  // );
    
    // 青龙适配，青龙微适配
    if(qlSwitch != 1){  // 选择金山文档
        resultHandle(resp, pos)
    }

}