// UPDATE.js 更新脚本
// 20241026

var confiWorkbook = 'CONFIG'  // 主配置表名称
var pushWorkbook = 'PUSH' // 推送表的名称
var emailWorkbook = 'EMAIL' // 邮箱表的名称

var workbook = [] // 存储已存在表数组

// 表中激活的区域的行数和列数
var row = 0;
var col = 0;
var maxRow = 100; // 规定最大行
var maxCol = 16; // 规定最大列
var colNum = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q']

// CONFIG表内容
// 推送昵称(推送位置标识)选项：若“是”则推送“账户名称”，若账户名称为空则推送“单元格Ax”，这两种统称为位置标识。若“否”，则不推送位置标识
var configContent = [
  ['工作表的名称', '备注', '只推送失败消息（是/否）', '推送昵称（是/否）', '是否存活', '更新时间', '消息', '推送时间', '推送方式',  '是否通知', '加入消息池'],
  ['xxx', '', '否', '是', '否' , '' , '' , '', '@all' , '是' , '否' ,],
  ['xxx', '', '否', '是', '否' , '' , '' , '', '@all' , '是' , '否' ,],
  ['noteyoudao', '有道云笔记', '否', '是', '是' , '' , '' , '', '@all' , '是' , '否' ,],
  ['tieba', '百度贴吧', '否', '是', '是' , '' , '' , '', '@all' , '是' , '否' ,],
  ['xxx', '', '否', '是', '否' , '' , '' , '', '@all' , '是' , '否' ,],
  ['xxx', '', '否', '是', '否' , '' , '' , '', '@all' , '是' , '否' ,],
  ['xxx', '', '否', '是', '否' , '' , '' , '', '@all' , '是' , '否' ,],
  ['xxx', '', '否', '是', '否' , '' , '' , '', '@all' , '是' , '否' ,],
  ['smzdm', '什么值得买抽奖', '否', '是', '是' , '' , '' , '', '@all' , '是' , '否' ,],
  ['toollu', '在线工具', '否', '是', '是' , '' , '' , '', '@all' , '是' , '否' ,],
  ['xxx', '', '否', '是', '否' , '' , '' , '', '@all' , '是' , '否' ,],
  ['xxx', '', '否', '是', '否' , '' , '' , '', '@all' , '是' , '否' ,],
  ['xxx', '', '否', '是', '否' , '' , '' , '', '@all' , '是' , '否' ,],
  ['xxx', '', '否', '是', '否' , '' , '' , '', '@all' , '是' , '否' ,],
  ['ddmc_ddyt', '叮咚买菜-叮咚鱼塘', '否', '是', '是' , '' , '' , '', '@all' , '是' , '否' ,],
  ['xxx', '', '否', '是', '否' , '' , '' , '', '@all' , '是' , '否' ,],
  ['xxx', '', '否', '是', '否' , '' , '' , '', '@all' , '是' , '否' ,],
  ['acfun', 'AcFun', '否', '是', '是' , '' , '' , '', '@all' , '是' , '否' ,],
  ['xmly', '喜马拉雅', '否', '是', '是' , '' , '' , '', '@all' , '是' , '否' ,],
  ['xxx', '', '否', '是', '否' , '' , '' , '', '@all' , '是' , '否' ,],
  ['en', '希沃白板', '否', '是', '是' , '' , '' , '', '@all' , '是' , '否' ,],
  ['xmc', '小木虫', '否', '是', '是' , '' , '' , '', '@all' , '是' , '否' ,],
  ['quark', '夸克网盘', '否', '是', '是' , '' , '' , '', '@all' , '是' , '否' ,],
  ['huluxia', '葫芦侠3楼', '否', '是', '是' , '' , '' , '', '@all' , '是' , '否' ,],
  ['xxx', '', '否', '是', '否' , '' , '' , '', '@all' , '是' , '否' ,],
  ['xxx', '', '否', '是', '否' , '' , '' , '', '@all' , '是' , '否' ,],
  ['ztebbs', '中兴社区', '否', '是', '是' , '' , '' , '', '@all' , '是' , '否' ,],
  ['mi', '小米商城', '否', '是', '是' , '' , '' , '', '@all' , '是' , '否' ,],
  ['kanxue', '看雪论坛', '否', '是', '是' , '' , '' , '', '@all' , '是' , '否' ,],
  ['bilibili', '哔哩哔哩', '否', '是', '是' , '' , '' , '', '@all' , '是' , '否' ,],
  ['vivo', 'vivo社区', '否', '是', '是' , '' , '' , '', '@all' , '是' , '否' ,],
  ['xxx', '', '否', '是', '否' , '' , '' , '', '@all' , '是' , '否' ,],
  ['wps_daka', 'wps(打卡版）', '否', '是', '是' , '' , '' , '', '@all' , '是' , '否' ,],
  ['golo', 'golo汽修大师', '否', '是', '是' , '' , '' , '', '@all' , '是' , '否' ,],
  ['xxx', '', '否', '是', '否' , '' , '' , '', '@all' , '是' , '否' ,],
  ['aliyun', '阿里云盘(自动更新token版)', '否', '是', '是' , '' , '' , '', '@all' , '是' , '否' ,],
  ['chinadsl', '宽带技术网', '否', '是', '是' , '' , '' , '', '@all' , '是' , '否' ,],
  ['xxx', '', '否', '是', '否' , '' , '' , '', '@all' , '是' , '否' ,],
  ['xxx', '', '否', '是', '否' , '' , '' , '', '@all' , '是' , '否' ,],
  ['ztemall', '中兴商城', '否', '是', '是' , '' , '' , '', '@all' , '是' , '否' ,],
  ['wnflb', '万能福利吧', '否', '是', '是' , '' , '' , '', '@all' , '是' , '否' ,],
  ['xxx', '', '否', '是', '否' , '' , '' , '', '@all' , '是' , '否' ,],
  ['xxx', '', '否', '是', '否' , '' , '' , '', '@all' , '是' , '否' ,],
  ['fwxs', '废文小说', '否', '是', '是' , '' , '' , '', '@all' , '是' , '否' ,],
  ['hxek', '鸿星尔克', '否', '是', '是' , '' , '' , '', '@all' , '是' , '否' ,],
  ['xxx', '', '否', '是', '否' , '' , '' , '', '@all' , '是' , '否' ,],
  ['xxx', '', '否', '是', '否' , '' , '' , '', '@all' , '是' , '否' ,],
  ['xxx', '', '否', '是', '否' , '' , '' , '', '@all' , '是' , '否' ,],
  ['ddai', '钉钉AI', '否', '是', '是' , '' , '' , '', '@all' , '是' , '否' ,],
  ['xxx', '', '否', '是', '否' , '' , '' , '', '@all' , '是' , '否' ,],
  ['xxx', '', '否', '是', '否' , '' , '' , '', '@all' , '是' , '否' ,],
  ['kyt', '科研通', '否', '是', '是' , '' , '' , '', '@all' , '是' , '否' ,],
  ['parsdata', '伊朗域名注册优惠码', '否', '是', '是' , '' , '' , '', '@all' , '是' , '否' ,],
  ['quarksave', '夸克订阅更新自动转存', '否', '是', '是' , '' , '' , '', '@all' , '是' , '否' ,],
  ['games70', '70games', '否', '是', '是' , '' , '' , '', '@all' , '是' , '否' ,],
  ['xxx', '', '否', '是', '否' , '' , '' , '', '@all' , '是' , '否' ,],
  ['rainyun', '雨云', '否', '是', '是' , '' , '' , '', '@all' , '是' , '否' ,],
  ['yhsh', '永辉生活', '否', '是', '是' , '' , '' , '', '@all' , '是' , '否' ,],
  ['xpnc', '兴攀农场', '否', '是', '是' , '' , '' , '', '@all' , '是' , '否' ,],
  ['bwcj', '霸王茶姬', '否', '是', '是' , '' , '' , '', '@all' , '是' , '否' ,],
  ['syns', '所有女生', '否', '是', '是' , '' , '' , '', '@all' , '是' , '否' ,],
  ['qcs', '屈臣氏', '否', '是', '是' , '' , '' , '', '@all' , '是' , '否' ,],
  ['hdl', '海底捞', '否', '是', '是' , '' , '' , '', '@all' , '是' , '否' ,],
  ['hzh', '华住会', '否', '是', '是' , '' , '' , '', '@all' , '是' , '否' ,],
  ['eswxlt', '恩山无线论坛', '否', '是', '是' , '' , '' , '', '@all' , '是' , '否' ,],
]

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
}
// var mosaic = "xxxxxxxx" // 马赛克
// var strFail = "否"
// var strTrue = "是"

// 主函数执行流程
storeWorkbook()
console.log("🥚 创建主分配表")
// const sheet = Application.ActiveSheet // 激活当前表
// sheet.Name = confiWorkbook  // 将当前工作表的名称改为 CONFIG
createSheet(confiWorkbook)
ActivateSheet(confiWorkbook)
editConfigSheet(configContent)  // editConfig()

console.log("🥚 创建推送表")
createSheet(pushWorkbook)
ActivateSheet(pushWorkbook)
editConfigSheet(pushContent)  // editPush()

console.log("🥚 创建邮箱表")
createSheet(emailWorkbook)
ActivateSheet(emailWorkbook)
editConfigSheet(emailContent)

// createSubConfig()

function count2DArray(arr) {
  let count = 0;
  for (let i = 0; i < arr.length; i++) {
      for (let j = 0; j < arr[i].length; j++) {
          count++;
      }
  }
  return count;
}

let length = configContent.length - 1
// console.log(length)
console.log("🍳 正在检索分配表，并进行创建")
for (let i = 0; i < length; i++) {
  let workbook = ""
  let subworkbook = ""
  // console.log(configContent[i+1][4])
  if(configContent[i+1][4] == "是"){  // 存活的才生成表
    // 部分表合并，如ddmc_ddgy，则以_为分割，生成ddmc表
    workbook = configContent[i+1][0].split("_")[0] // 使用下划线作为分隔符，取第一个部分)
    ActivateSheet(workbook)  // 根据CONFIG表来生成
    editConfigSheet(subConfigContent)   // editSubConfig()

    // 检查是否有定制化内容，有则生成
    try{
      subworkbook = subConfig[workbook]
      // console.log(subworkbook)
      if(subworkbook != undefined){
        ActivateSheet(workbook) // 激活阿里云盘分配置表
        editConfigSheet(subworkbook)  // editSubConfigCustomized(subConfigAliyundrive)
        // console.log("存在定制化内容")
      }
    }catch{
      // 无定制化内容
      // console.log("无定制化内容")
    }
  }
}

// 加入跳转超链接
hypeLink()

// CONFIG表添加超链接
function hypeLink(){
  let workSheet= Application.Sheets(confiWorkbook) //配置表
  let workUsedRowEnd = workSheet.UsedRange.RowEnd //用户使用表格的最后一行

  for(let row = 2; row <= workUsedRowEnd; row++){
    link_name=workSheet.Range("A" + row).Text
    if (link_name == "") {
      break; // 如果为空行，则提前结束读取
    }

    link_name ='=HYPERLINK("#'+link_name+'!$A$1","'+link_name+'")' //设置超链接
    //console.log(link_name)  // HYPERLINK("#PUSH!$A$1","PUSH")
    workSheet.Range("A" + row).Value =link_name 
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

  console.log("✨ 当前激活表已存在：" + row + "行，" + col + "列")
}

// 激活工作表函数
function ActivateSheet(sheetName) {
  let flag = 0;
  try {
    // 激活工作表
    let sheet = Application.Sheets.Item(sheetName)
    sheet.Activate()
    console.log("🥚 激活工作表：" + sheet.Name)
    flag = 1;
  } catch {
    flag = 0;
    console.log("📢 无法激活工作表，工作表可能不存在")
    console.log("🎉 创建此表：" + sheetName)
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
      Application.Range(colNum[i] + 1).Value = content[0][i]
    }

    row += 1; // 让行数加1，代表写入了表头。
  }

  // 从已写入的行的后一行开始逐行写入数据
  // 先写行
  for (let i = 1 + row; i <= lengthRow; i++) {  // 从未写入区域开始写
    for (let j = 0; j < lengthCol; j++) {
      Application.Range(colNum[j] + i).Value = content[i - 1][j]
    }
  }
  // 再写列
  for (let j = col; j < lengthCol; j++) {
    for (let i = 1; i <= lengthRow; i++) {  // 从未写入区域开始写
      Application.Range(colNum[j] + i).Value = content[i - 1][j]
    }
  }
}

// 编辑主分配表(已弃用)
function editConfig() {
  determineRowCol();
  let lengthRow = configContent.length
  let lengthCol = configContent[0].length
  if (row == 0) { // 如果行数为0，认为是空表,开始写表头
    for (let i = 0; i < lengthCol; i++) {
      Application.Range(colNum[i] + 1).Value = configContent[0][i]
    }

    row += 1; // 让行数加1，代表写入了表头。
  }

  // 从已写入的行的后一行开始逐行写入数据
  // 先写行
  for (let i = 1 + row; i <= lengthRow; i++) {  // 从未写入区域开始写
    for (let j = 0; j < lengthCol; j++) {
      Application.Range(colNum[j] + i).Value = configContent[i - 1][j]
    }
  }
  // 再写列
  for (let j = col; j < lengthCol; j++) {
    for (let i = 1; i <= lengthRow; i++) {  // 从未写入区域开始写
      Application.Range(colNum[j] + i).Value = configContent[i - 1][j]
    }
  }
}

// 编辑推送表(已弃用)
function editPush() {
  determineRowCol();
  let lengthRow = pushContent.length
  let lengthCol = pushContent[0].length
  if (row == 0) { // 如果行数为0，认为是空表,开始写表头
    for (let i = 0; i < lengthCol; i++) {
      Application.Range(colNum[i] + 1).Value = pushContent[0][i]
    }

    row += 1; // 让行数加1，代表写入了表头。
  }

  // 从已写入的行的后一行开始逐行写入数据
  // 先写行
  for (let i = 1 + row; i <= lengthRow; i++) {  // 从未写入区域开始写
    for (let j = 0; j < lengthCol; j++) {
      Application.Range(colNum[j] + i).Value = pushContent[i - 1][j]
    }
  }
  // 再写列
  for (let j = col; j < lengthCol; j++) {
    for (let i = 1; i <= lengthRow; i++) {  // 从未写入区域开始写
      Application.Range(colNum[j] + i).Value = pushContent[i - 1][j]
    }
  }
}

// 编辑分配置表(已弃用)
function editSubConfig() {
  determineRowCol();
  let lengthRow = subConfigContent.length
  let lengthCol = subConfigContent[0].length
  if (row == 0) { // 如果行数为0，认为是空表,开始写表头
    for (let i = 0; i < lengthCol; i++) {
      Application.Range(colNum[i] + 1).Value = subConfigContent[0][i]
    }

    row += 1; // 让行数加1，代表写入了表头。
  }

  // 从已写入的行的后一行开始逐行写入数据
  // 先写行
  for (let i = 1 + row; i <= lengthRow; i++) {  // 从未写入区域开始写
    for (let j = 0; j < lengthCol; j++) {
      Application.Range(colNum[j] + i).Value = subConfigContent[i - 1][j]
    }
  }
  // 再写列
  for (let j = col; j < lengthCol; j++) {
    for (let i = 1; i <= lengthRow; i++) {  // 从未写入区域开始写
      Application.Range(colNum[j] + i).Value = subConfigContent[i - 1][j]
    }
  }
}

// 编辑定制化分配置表(已弃用)
function editSubConfigCustomized(content) {
  determineRowCol();
  let lengthRow = content.length
  let lengthCol = content[0].length
  if (row == 0) { // 如果行数为0，认为是空表,开始写表头
    for (let i = 0; i < lengthCol; i++) {
      Application.Range(colNum[i] + 1).Value = content[0][i]
    }

    row += 1; // 让行数加1，代表写入了表头。
  }

  // 从已写入的行的后一行开始逐行写入数据
  // 先写行
  for (let i = 1 + row; i <= lengthRow; i++) {  // 从未写入区域开始写
    for (let j = 0; j < lengthCol; j++) {
      Application.Range(colNum[j] + i).Value = content[i - 1][j]
    }
  }
  // 再写列
  for (let j = col; j < lengthCol; j++) {
    for (let i = 1; i <= lengthRow; i++) {  // 从未写入区域开始写
      Application.Range(colNum[j] + i).Value = content[i - 1][j]
    }
  }
}


// 创建分配置表(已弃用)
function createSubConfig() {
  let length = subConfigWorkbook.length
  for (let i = 0; i < length; i++) {
    console.log("🥚 创建分配置表：" + subConfigWorkbook[i])
    createSheet(subConfigWorkbook[i])
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
function createSheet(name) {
  // const defaultName = Application.Sheets.DefaultNewSheetName
  // 工作表对象
  if (!workbookComp(name)) {
    Application.Sheets.Add(
      null,
      Application.ActiveSheet.Name,
      1,
      Application.Enum.XlSheetType.xlWorksheet,
      name
    )
  }
}