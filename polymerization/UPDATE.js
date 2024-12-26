/*
    脚本名称：UPDATE.js
    脚本兼容: airsript 1.0、airscript 2.0
    更新时间：20241226
    备注：更新脚本。用于自动生成表格，以及追加表格数据
          适配airsript 1.0版本及airscript 2.0版本
    其他：若想添加新内容，请搜索（修改这里），按照格式修改
*/

var confiWorkbook = 'CONFIG'  // 主配置表名称
var pushWorkbook  = 'PUSH' // 推送表的名称
var emailWorkbook = 'EMAIL' // 邮箱表的名称
var version = 1 // 版本类型，自动识别并适配。默认为airscript 1.0，否则为2.0（Beta）

// 表中激活的区域的行数和列数
var workbook = [] // 存储已存在表数组
var row = 0;
var col = 0;
var maxRow = 100; // 规定最大行
var maxCol = 22; // 规定最大列
var colNum = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V']

// CONFIG表内容
// 推送昵称(推送位置标识)选项：若“是”则推送“账户名称”，若账户名称为空则推送“单元格Ax”，这两种统称为位置标识。若“否”，则不推送位置标识
// 存放CONFIG表内容，标题+标题下内容
var configContent = [];
// 实际写入CONFIG表的值
// CONFIG表标题
configTitle = ['工作表的名称', '备注', '只推送失败消息（是/否）', '推送昵称（是/否）', '是否存活', '程序结束时间', '消息', '推送时间', '推送方式', '是否通知', '加入消息池', '推送优先级', '当日可推送次数', '当日剩余推送次数']
// 定义CONFIG表标题下内容的默认值
var configBodyDefault = ['xxx', '', '否', '是', '是', '', '', '', '@all', '是', '否', '0', '1', ''];
// CONFIG表标题下内容
var configBody = [
    { name: 'xxx', note: '', isAlive: '否',},
    { name: 'xxx', note: '', isAlive: '否',},
    { name: 'noteyoudao', note: '有道云笔记',},
    { name: 'tieba', note: '百度贴吧',},
    { name: 'xxx', note: '', isAlive: '否',},
    { name: 'xxx', note: '', isAlive: '否',},
    { name: 'xxx', note: '', isAlive: '否',},
    { name: 'xxx', note: '', isAlive: '否',},
    { name: 'smzdm', note: '什么值得买抽奖',},
    { name: 'toollu', note: '在线工具',},
    { name: 'xxx', note: '', isAlive: '否',},
    { name: 'xxx', note: '', isAlive: '否',},
    { name: 'xxx', note: '', isAlive: '否',},
    { name: 'xxx', note: '', isAlive: '否',},
    { name: 'ddmc_ddyt', note: '叮咚买菜-叮咚鱼塘',},
    { name: 'xxx', note: '', isAlive: '否',},
    { name: 'xxx', note: '', isAlive: '否',},
    { name: 'acfun', note: 'AcFun',},
    { name: 'xmly', note: '喜马拉雅',},
    { name: 'xxx', note: '', isAlive: '否',},
    { name: 'en', note: '希沃白板',},
    { name: 'xmc', note: '小木虫',},
    { name: 'quark', note: '夸克网盘',},
    { name: 'huluxia', note: '葫芦侠3楼',},
    { name: 'xxx', note: '', isAlive: '否',},
    { name: 'xxx', note: '', isAlive: '否',},
    { name: 'ztebbs', note: '中兴社区',},
    { name: 'mi', note: '小米商城',},
    { name: 'kanxue', note: '看雪论坛',},
    { name: 'bilibili', note: '哔哩哔哩',},
    { name: 'vivo', note: 'vivo社区',},
    { name: 'xxx', note: '', isAlive: '否',},
    { name: 'wps_daka', note: 'wps(打卡版）',},
    { name: 'golo', note: 'golo汽修大师',},
    { name: 'xxx', note: '', isAlive: '否',},
    { name: 'aliyun', note: '阿里云盘(自动更新token版)',},
    { name: 'chinadsl', note: '宽带技术网',},
    { name: 'xxx', note: '', isAlive: '否',},
    { name: 'xxx', note: '', isAlive: '否',},
    { name: 'ztemall', note: '中兴商城',},
    { name: 'wnflb', note: '万能福利吧',},
    { name: 'xxx', note: '', isAlive: '否',},
    { name: 'xxx', note: '', isAlive: '否',},
    { name: 'fwxs', note: '废文小说',},
    { name: 'hxek', note: '鸿星尔克',},
    { name: 'xxx', note: '', isAlive: '否',},
    { name: 'xxx', note: '', isAlive: '否',},
    { name: 'xxx', note: '', isAlive: '否',},
    { name: 'ddai', note: '钉钉AI',},
    { name: 'xxx', note: '', isAlive: '否',},
    { name: 'xxx', note: '', isAlive: '否',},
    { name: 'kyt', note: '科研通',},
    { name: 'parsdata', note: '伊朗域名注册优惠码',},
    { name: 'quarksave', note: '夸克订阅更新自动转存',},
    { name: 'games70', note: '70games',},
    { name: 'xxx', note: '', isAlive: '否',},
    { name: 'xxx', note: '', isAlive: '否',},
    { name: 'yhsh', note: '永辉生活',},
    { name: 'xpnc', note: '兴攀农场',},
    { name: 'bwcj', note: '霸王茶姬',},
    { name: 'syns', note: '所有女生',},
    { name: 'qcs', note: '屈臣氏',},
    { name: 'hdl', note: '海底捞',},
    { name: 'hzh', note: '华住会',},
    { name: 'eswxlt', note: '恩山无线论坛',},
    { name: 'steamtools', note: 'steamtools',},
    { name: 'xmdl', note: '熊猫代理',},
    { name: 'linkai', note: 'LinkAi',},
    { name: 'hfweather', note: '和风天气', pushPriority: '1',},
    { name: 'dml', note: '达美乐',},
    { name: 'ciba', note: '词霸每日一句',},

    // { name: '（修改这里）', note: '（修改这里）',},  // 添加新增内容
];


// 定义字段映射关系，标题和键的对应关系，用于动态个性化处理
var configTitleMapping = {
    '工作表的名称': 'name',
    '备注': 'note',
    '只推送失败消息（是/否）': 'pushFailureOnly',
    '推送昵称（是/否）': 'pushNickname',
    '是否存活': 'isAlive',
    '更新时间': 'updateTime',
    '消息': 'message',
    '推送时间': 'pushTime',
    '推送方式': 'pushMethod',
    '是否通知': 'notify',
    '加入消息池': 'addToMessagePool',
    '推送优先级': 'pushPriority',
    '当日可推送次数': 'dailyPushLimit',
    '当日剩余推送次数': 'remainingDailyPushes',
};

// PUSH表内容 		
var pushContent = [
  ['推送类型', '推送识别号(如：token、key)', '是否推送（是/否）'],
  ['bark', 'xxxxxxxx', '否'],
  ['pushplus', 'xxxxxxxx', '否'],
  ['ServerChan', 'xxxxxxxx', '否'],
  ['email', '若要邮箱发送，请配置EMAIL表', '否'],
  ['dingtalk', 'xxxxxxxx', '否'],
  ['discord', '请填入镜像webhook链接,自行处理Query参数', '否'],
  ['qywx', 'xxxxxxxx', '否'],
  ['xizhi', 'xxxxxxxx', '否'],
  ['jishida', 'xxxxxxxx', '否'],
  ['wxpusher', 'appToken|uid', '否'],
]

// email表内容
var emailContent = [
  ['SMTP服务器域名', '端口', '发送邮箱', '授权码'],
  ['smtp.qq.com', '465', 'xxxxxxxx@qq.com', 'xxxxxxxx']
]

// 分配置表内容
var subConfigContent = [
  ['cookie(默认20个)', '是否执行(是/否)', '账号名称(可不填写)'],
  ['xxxxxxxx1', '是', '昵称1'],
  ['xxxxxxxx2', '否', '昵称2']
]

// 定制化分配置表内容，叮咚买菜
var subConfigDdmc = [
  ['cookie(默认20个)', '是否执行(是/否)', '账号名称(可不填写)', 'seedId_ddgy', 'propsId_ddgy', 'seedId_ddyt', 'propsId_ddty'],
  ['xxxxxxxx1', '是', '昵称1', '填果园seedId', '填果园propsId', '填鱼塘seedId', '填鱼塘propsId'],
  ['xxxxxxxx2', '否', '昵称2', '填果园seedId', '填果园propsId', '填鱼塘seedId', '填鱼塘propsId']
]

// 定制化分配置表内容，WPS
var subConfigWps = [
  ['cookie(默认20个)', '是否执行(是/否)', '账号名称(可不填写)', '转存PPT(是/否)', '是否渠道1打卡(是/否)', '是否渠道2打卡(是/否)', 'Signature(渠道2)'],
  ['xxxxxxxx1', '是', '昵称1', '否', '是', '否' , 'xxxxxxxx' ,],
  ['xxxxxxxx2', '否', '昵称2', '否', '是', '否' , 'xxxxxxxx' ,]
]

// 定制化分配置表内容，小米商城
var subConfigMi = [
  ['cookie(默认20个)', '是否执行(是/否)', '账号名称(可不填写)', 'mishop-client-id'],
  ['xxxxxxxx1', '是', '昵称1', '100'],
  ['xxxxxxxx2', '否', '昵称2', '100']
]

// 定制化分配置表内容，golo汽修大师
var subConfigGolo = [
  ['cookie(默认20个)', '是否执行(是/否)', '账号名称(可不填写)', 'username', 'password'],
  ['xxxxxxxx1', '是', '昵称1', '此格填用户名', '此格填密码'],
  ['xxxxxxxx2', '否', '昵称2', '此格填用户名', '此格填密码']
]

// 定制化分配置表内容，阿里云盘(自动更新token)
var subConfigAliyunToken = [
  ['cookie(默认20个)', '是否执行(是/否)', '账号名称(可不填写)', '是否领取自动备份的奖励(是/否)', 'token登陆时间', '签到结果'],
  ['xxxxxxxx1', '是', '昵称1', '否', '无', '无'],
  ['xxxxxxxx2', '否', '昵称2', '否', '无', '无']
]

// 定制化分配置表内容，hxek
var subConfigHxek = [
  ['cookie(默认20个)', '是否执行(是/否)', '账号名称(可不填写)', "memberId", "enterpriseId"],
  ['xxxxxxxx1', '是', '昵称1', 'xxx', 'xxx'],
  ['xxxxxxxx2', '否', '昵称2', 'xxx', 'xxx']
]

// 定制化分配置表内容，parsdata
var subConfigParsdata = [
  ['cookie(默认20个)', '是否执行(是/否)', '账号名称(可不填写)', '优惠码'],
  ['xxxxxxxx1', '是', '昵称1', ''],
  ['xxxxxxxx2', '否', '昵称2', '']
]

// 定制化分配置表内容，雨云
var subConfigRainyun = [
  ['cookie(默认20个)', '是否执行(是/否)', '账号名称(可不填写)', '账号', '密码'],
  ['xxxxxxxx1', '是', '昵称1', '', ''],
  ['xxxxxxxx2', '否', '昵称2', '', '']
]

// 定制化分配置表内容，霸王茶姬
var subConfigBwcj = [
  ['cookie(默认20个)', '是否执行(是/否)', '账号名称(可不填写)', 'userid'],
  ['xxxxxxxx1', '是', '昵称1', ''],
  ['xxxxxxxx2', '否', '昵称2', '']
]

// 定制化分配置表内容，海底捞
var subConfigHdl = [
  ['cookie(默认20个)', '是否执行(是/否)', '账号名称(可不填写)', 'uid'],
  ['xxxxxxxx1', '是', '昵称1', ''],
  ['xxxxxxxx2', '否', '昵称2', '']
]

// 定制化分配置表内容，屈臣氏
var subConfigQcs = [
  ['cookie(默认20个)', '是否执行(是/否)', '账号名称(可不填写)', 'openId', 'unionId'],
  ['xxxxxxxx1', '是', '昵称1', '', ''],
  ['xxxxxxxx2', '否', '昵称2', '', '']
]

// 定制化分配置表内容，恩山无线论坛
var subConfigEswxlt = [
  ['cookie(默认20个)', '是否执行(是/否)', '账号名称(可不填写)', 'openId', 'unionId'],
  ['xxxxxxxx1', '是', '昵称1', '', ''],
  ['xxxxxxxx2', '否', '昵称2', '', '']
]

// 定制化分配置表内容，xmdl熊猫代理
var subConfigXmdl = [
  ['cookie(默认20个)', '是否执行(是/否)', '账号名称(可不填写)', 'username', 'password'],
  ['xxxxxxxx1', '是', '昵称1', '此格填用户名', '此格填密码'],
  ['xxxxxxxx2', '否', '昵称2', '此格填用户名', '此格填密码']
]

// 定制化分配置表内容，hfweather和风天气
var subConfigHfweather = [
  ['cookie(默认20个)', '是否执行(是/否)', '账号名称(可不填写)', '省份/城市/区/经纬度/ID（手动输入）', '今日天气预报', '实时天气预报', '当天生活指数', '天气灾害预警', '逐小时天气预报', '分钟级降水','其他模式（可选）','消息过滤器（可选）','高级配置（可选）','实际定位（自动生成）','地区ID（自动生成）','经纬度（自动生成）','一致性校验（自动生成）','冗余观测站（自动生成）'],
  ['xxxxxxxx1', '是', '昵称1', '朝阳区', '是', '否', '否', '否', '否', '否', '', '', '', '', '', '', '', ''],
  // ['xxxxxxxx2', '否', '昵称2', '是', '否', '否', '否', '否', '否', '', '', '', '', '', '', '', '']
]

// 定制化分配置表内容，ciba词霸每日一句
var subConfigCiba = [
  ['cookie(默认20个)', '是否执行(是/否)', '账号名称(可不填写)', '中文每日一句', '英文每日一句'],
  ['xxxxxxxx1', '是', '昵称1', '是', '是'],
  ['xxxxxxxx2', '否', '昵称2', '是', '是']
]

// 定制化表
var subConfig = {
  "ddmc"  : subConfigDdmc, 
  "wps"  : subConfigWps,
  "golo"  : subConfigGolo,
  "aliyun" : subConfigAliyunToken,
  "hxek" : subConfigHxek,
  "parsdata":subConfigParsdata,
  "rainyun":subConfigRainyun,
  "bwcj":subConfigBwcj,
  "hdl":subConfigHdl,
  "qcs":subConfigQcs,
  "eswxlt":subConfigEswxlt,
  "xmdl":subConfigXmdl,
  "hfweather":subConfigHfweather,
  "ciba":subConfigCiba,
}
// var mosaic = "xxxxxxxx" // 马赛克
// var strFail = "否"
// var strTrue = "是"

main()  // 入口

// CONFIG表添加超链接
function hypeLink(){
  // console.log("添加超链接")
  let workSheet= Application.Sheets(confiWorkbook) //配置表
  // let workUsedRowEnd = workSheet.UsedRange.RowEnd //用户使用表格的最后一行
  let rowcol = getRowCol() 
  let workUsedRowEnd = rowcol[0]  // 行
  // console.log(workUsedRowEnd)
  for(let row = 2; row <= workUsedRowEnd; row++){
    link_name=workSheet.Range("A" + row).Text
    if (link_name == "") {
      break; // 如果为空行，则提前结束读取
    }
    link_name ='=HYPERLINK("#'+link_name+'!$A$1","'+link_name+'")' //设置超链接
    // console.log(link_name)  // HYPERLINK("#PUSH!$A$1","PUSH")
    if(version == 1){
      // airscipt 1.0
      workSheet.Range("A" + row).Value =link_name 
    }else{
      // airscipt 2.0
      workSheet.Range("A" + row).Value2 =link_name
    }
    
  }
}

// 判断表格行列数，并记录目前已写入的表格行列数。目的是为了不覆盖原有数据，便于更新
function determineRowCol() {
  for (let i = 1; i < maxRow; i++) {
    let content = Application.Range("A" + i).Text
    if (content == "")  // 如果为空行，则提前结束读取
    {
      row = i - 1;  // 记录的是存在数据所在的行
      break;
    }
  }
  // 超过最大行了，认为row为0，从头开始
  let length = colNum.length
  for (let i = 1; i <= length; i++) {
    content = Application.Range(colNum[i - 1] + "1").Text
    if (content == "")  // 如果为空行，则提前结束读取
    {
      col = i - 1;  // 记录的是存在数据所在的行
      break;
    }
  }
  // 超过最大行了，认为col为0，从头开始
  // console.log("✨ 当前激活表已存在：" + row + "行，" + col + "列")
}

// 获取当前激活表的表的行列
function getRowCol() {
  let row = 0
  let col = 0
  for (let i = 1; i < maxRow; i++) {
    let content = Application.Range("A" + i).Text
    if (content == "")  // 如果为空行，则提前结束读取
    {
      row = i - 1;  // 记录的是存在数据所在的行
      break;
    }
  }
  // 超过最大行了，认为row为0，从头开始
  let length = colNum.length
  for (let i = 1; i <= length; i++) {
    content = Application.Range(colNum[i - 1] + "1").Text
    if (content == "")  // 如果为空行，则提前结束读取
    {
      col = i - 1;  // 记录的是存在数据所在的行
      break;
    }
  }
  // 超过最大行了，认为col为0，从头开始

  // console.log("✨ 当前激活表已存在：" + row + "行，" + col + "列")
  return [row, col]
}

// 激活工作表函数
function ActivateSheet(sheetName) {
  let flag = 0;
  try {
    let sheet = Application.Sheets.Item(sheetName)
    sheet.Activate()
    // console.log("🍾 激活工作表：" + sheet.Name)
    flag = 1;
  } catch {
    flag = 0;
    // console.log("📢 无法激活工作表，工作表可能不存在")
    console.log("🪄 创建工作表：" + sheetName)
    createSheet(sheetName)
  }
  return flag;
}

// 统一编辑表函数
function editConfigSheet(content) {
  determineRowCol();
  let lengthRow = content.length
  let lengthCol = content[0].length
  if (row == 0) { // 如果行数为0，认为是空表,开始写表头
    for (let i = 0; i < lengthCol; i++) {
      if(version == 1){
        // airscipt 1.0
        Application.Range(colNum[i] + 1).Value = content[0][i]
      }else{
        // airscript 2.0(Beta)
        Application.Range(colNum[i] + 1).Value2 = content[0][i]
      }
      
    }

    row += 1; // 让行数加1，代表写入了表头。
  }

  // 从已写入的行的后一行开始逐行写入数据
  // 先写行
  for (let i = 1 + row; i <= lengthRow; i++) {  // 从未写入区域开始写
    for (let j = 0; j < lengthCol; j++) {
      if(version == 1){
        // airscipt 1.0
        Application.Range(colNum[j] + i).Value = content[i - 1][j]
      }else{
        // airscript 2.0(Beta)
        Application.Range(colNum[j] + i).Value2 = content[i - 1][j]
      }
    }
  }
  // 再写列
  for (let j = col; j < lengthCol; j++) {
    for (let i = 1; i <= lengthRow; i++) {  // 从未写入区域开始写
      if(version == 1){
        // airscipt 1.0
        Application.Range(colNum[j] + i).Value = content[i - 1][j]
      }else{
        // airscript 2.0(Beta)
        Application.Range(colNum[j] + i).Value2 = content[i - 1][j]
      }
    }
  }
}

// 存储已存在的表
function storeWorkbook() {
  // 工作簿（Workbook）中所有工作表（Sheet）的集合,下面两种写法是一样的
  let sheets = Application.ActiveWorkbook.Sheets
  sheets = Application.Sheets

  // 打印所有工作表的名称
  for (let i = 1; i <= sheets.Count; i++) {
    workbook[i - 1] = (sheets.Item(i).Name)
    // console.log(workbook[i-1])
  }
}

// 判断表是否已存在
function workbookComp(name) {
  let flag = 0;
  let length = workbook.length
  for (let i = 0; i < length; i++) {
    if (workbook[i] == name) {
      flag = 1;
      console.log("✨ " + name + "表已存在")
      break
    }
  }
  return flag
}

// 创建表，若表已存在则不创建，直接写入数据
function createSheet(sheetname) {
  // const defaultName = Application.Sheets.DefaultNewSheetName
  // 工作表对象
  if (!workbookComp(sheetname)) {
    console.log("🪄 创建工作表：" + sheetname)
    try{
        Application.Sheets.Add(
        null,
        Application.ActiveSheet.Name,
        1,
        Application.Enum.XlSheetType.xlWorksheet,
        sheetname
      )
      
    }catch{
      // console.log("😶‍🌫️ 适配airscript 2.0版本")
      version = 2 // 设置版本为2.0
      let newSheet = Application.Sheets.Add(undefined, undefined, undefined, xlWorksheet)
      // let newSheet = Application.Worksheets.Add()
      newSheet.Name = sheetname
    }

  }
}

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

// 主函数执行流程
function main(){
  checkVesion() // 版本检测，以进行不同版本的适配

  // 动态写入CONFIG表数组数据操作
  // 加入标题
  configContent[0] = configTitle
  // 写入标题下内容
  for (let i = 0; i < configBody.length; i++) {
      let row = [];
      for (let j = 0; j < configContent[0].length; j++) {
          let fieldName = configContent[0][j];
          let fieldValue = configBody[i][configTitleMapping[fieldName]];
          if (fieldValue === undefined) {
              fieldValue = configBodyDefault[j]; // 如果字段不存在，使用默认值
          }
          row.push(fieldValue);
      }
      configContent.push(row);
  }

  storeWorkbook()
  // console.log("🪄 创建主分配表")
  // const sheet = Application.ActiveSheet // 激活当前表
  // sheet.Name = confiWorkbook  // 将当前工作表的名称改为 CONFIG
  createSheet(confiWorkbook)
  ActivateSheet(confiWorkbook)
  editConfigSheet(configContent)  // editConfig()

  // 加入跳转超链接
  hypeLink()

  // console.log("🪄 创建推送表")
  createSheet(pushWorkbook)
  ActivateSheet(pushWorkbook)
  editConfigSheet(pushContent)  // editPush()

  // console.log("🪄 创建邮箱表")
  createSheet(emailWorkbook)
  ActivateSheet(emailWorkbook)
  editConfigSheet(emailContent)

  let length = configContent.length - 1
  // console.log(length)
  console.log("🍳 正在检索分配置表进行创建")
  for (let i = 0; i < length; i++) {
    let workbook = ""
    let subworkbook = ""
    // console.log(configContent[i+1][4])
    if(configContent[i+1][4] == "是"){  // 存活的才生成表
      // 部分表合并，如ddmc_ddgy，则以_为分割，生成ddmc表
      workbook = configContent[i+1][0].split("_")[0] // 使用下划线作为分隔符，取第一个部分)
      ActivateSheet(workbook)  // 根据CONFIG表来生成
      editConfigSheet(subConfigContent)

      // 检查是否有定制化内容，有则生成
      try{
        subworkbook = subConfig[workbook]
        // console.log(subworkbook)
        if(subworkbook != undefined){
          ActivateSheet(workbook) // 激活分配置表
          editConfigSheet(subworkbook)
          // console.log("存在定制化内容")
        }
      }catch{
        // 无定制化内容
        // console.log("无定制化内容")
      }
    }
  }

  console.log("🎉 更新完成")

}