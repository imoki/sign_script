var confiWorkbook = 'CONFIG'  // 主配置表名称
var pushWorkbook = 'PUSH' // 推送表的名称
// 分配置表名称
var subConfigWorkbook=['aliyundrive_multiuser','52pojie','noteyoudao','wps','tieba'];
var workbook = [] // 存储已存在表数组

// 表中激活的区域的行数和列数
var row = 0;
var col = 0;
var maxRow = 100; // 规定最大行
var maxCol = 16; // 规定最大列
var colNum = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q']

// CONFIG表内容
var configContent=[
  ['工作表的名称','备注','只推送失败消息（是/否）','推送昵称（是/否）'],
  ['aliyundrive_multiuser','阿里云盘（多用户版）','否','否'],
  ['52pojie','吾爱破解','否','否'],
  ['noteyoudao','有道云笔记','否','否'],
  ['tieba','百度贴吧','否','否'],
  ['wps_light','wps(轻量版)','否','否'],
  ['wps_client','wps(客户端版','否','是'],
  ['wps_docer','wps(稻壳版）','否','否'],
]

// PUSH表内容 		
var pushContent=[
  ['推送类型','推送识别号(如：token、key)','是否推送（是/否）'],
  ['bark', 'xxxxxxxx', '否'],
  ['pushplus', 'xxxxxxxx', '否'],
  ['ServerChan', 'xxxxxxxx', '否'],
]

// 分配置表内容
var subConfigContent = [
  ['cookie(默认20个)','是否执行(是/否)','账号名称(可不填写)'],
  ['xxxxxxxx1', '是', '昵称1'],
  ['xxxxxxxx2', '否', '昵称2']
]

var mosaic = "xxxxxxxx" // 马赛克
var strFail = "否"
var strTrue = "是"

// 主函数执行流程
storeWorkbook()
console.log("创建主分配表")
// const sheet = Application.ActiveSheet // 激活当前表
// sheet.Name = confiWorkbook  // 将当前工作表的名称改为 CONFIG
createSheet(confiWorkbook)
ActivateSheet(confiWorkbook)
editConfig()

console.log("创建推送表")
createSheet(pushWorkbook)
ActivateSheet(pushWorkbook)
editPush()


createSubConfig()
let length = subConfigWorkbook.length
for(let i = 0; i < length; i++){
  ActivateSheet(subConfigWorkbook[i])
  editSubConfig()
}

// 判断表格行列数，并记录目前已写入的表格行列数。目的是为了不覆盖原有数据，便于更新
function determineRowCol(){
  for (let i = 1; i < maxRow; i++){
    let content = Application.Range("A" + i).Text
    if(content == "")  // 如果为空行，则提前结束读取
    {
      row = i - 1;  // 记录的是存在数据所在的行
      break;
    }  
  }
  // 超过最大行了，认为row为0，从头开始
  let length = colNum.length
  for (let i = 1; i <= length; i++){
    content = Application.Range(colNum[i-1] + "1").Text
    if(content == "")  // 如果为空行，则提前结束读取
    {
      col = i - 1;  // 记录的是存在数据所在的行
      break;
    }  
  }
  // 超过最大行了，认为col为0，从头开始

  console.log("当前激活表已存在：" + row + "行，" + col + "列")
}

// 激活工作表函数
function ActivateSheet(sheetName){
  let flag = 0;
  try{
    // 激活工作表
    let sheet = Application.Sheets.Item(sheetName)
    sheet.Activate()
    console.log("激活工作表：" + sheet.Name)
    flag = 1;
  }catch{
    flag = 0;
    console.log("无法激活工作表，工作表可能不存在")
  }
  return flag;
}


// 编辑主分配表
function editConfig(){
  determineRowCol();
  let lengthRow = configContent.length
  let lengthCol = configContent[0].length
  if(row == 0){ // 如果行数为0，认为是空表,开始写表头
    for(let i = 0; i< lengthCol; i++){
      Application.Range(colNum[i] + 1).Value = configContent[0][i]
    }

    row += 1; // 让行数加1，代表写入了表头。
  }

  // 从已写入的行的后一行开始逐行写入数据
  // 先写行
  for(let i = 1 + row; i <= lengthRow; i++){  // 从未写入区域开始写
    for(let j = 0; j < lengthCol; j++){
      Application.Range(colNum[j] + i).Value = configContent[i-1][j]
    }
  }
  // 再写列
  for(let j = col; j < lengthCol; j++){
    for(let i = 1; i <= lengthRow; i++){  // 从未写入区域开始写
      Application.Range(colNum[j] + i).Value = configContent[i-1][j]
    }
  }
}

// 编辑推送表
function editPush(){
  determineRowCol();
  let lengthRow = pushContent.length
  let lengthCol = pushContent[0].length
  if(row == 0){ // 如果行数为0，认为是空表,开始写表头
    for(let i = 0; i< lengthCol; i++){
      Application.Range(colNum[i] + 1).Value = pushContent[0][i]
    }

    row += 1; // 让行数加1，代表写入了表头。
  }

  // 从已写入的行的后一行开始逐行写入数据
  // 先写行
  for(let i = 1 + row; i <= lengthRow; i++){  // 从未写入区域开始写
    for(let j = 0; j < lengthCol; j++){
      Application.Range(colNum[j] + i).Value = pushContent[i-1][j]
    }
  }
  // 再写列
  for(let j = col; j < lengthCol; j++){
    for(let i = 1; i <= lengthRow; i++){  // 从未写入区域开始写
      Application.Range(colNum[j] + i).Value = pushContent[i-1][j]
    }
  }
}

// 编辑分配置表
function editSubConfig(){
  determineRowCol();
  let lengthRow = subConfigContent.length
  let lengthCol = subConfigContent[0].length
  if(row == 0){ // 如果行数为0，认为是空表,开始写表头
    for(let i = 0; i< lengthCol; i++){
      Application.Range(colNum[i] + 1).Value = subConfigContent[0][i]
    }

    row += 1; // 让行数加1，代表写入了表头。
  }

  // 从已写入的行的后一行开始逐行写入数据
  // 先写行
  for(let i = 1 + row; i <= lengthRow; i++){  // 从未写入区域开始写
    for(let j = 0; j < lengthCol; j++){
      Application.Range(colNum[j] + i).Value = subConfigContent[i-1][j]
    }
  }
  // 再写列
  for(let j = col; j < lengthCol; j++){
    for(let i = 1; i <= lengthRow; i++){  // 从未写入区域开始写
      Application.Range(colNum[j] + i).Value = subConfigContent[i-1][j]
    }
  }
}


// 创建分配置表
function createSubConfig(){
  let length = subConfigWorkbook.length
  for(let i = 0; i < length; i++){
    console.log("创建分配置表：" + subConfigWorkbook[i])
    createSheet(subConfigWorkbook[i])
  }
}

// 存储已存在的表
function storeWorkbook(){
  // 工作簿（Workbook）中所有工作表（Sheet）的集合,下面两种写法是一样的
  let sheets = Application.ActiveWorkbook.Sheets
  sheets = Application.Sheets

  // 打印所有工作表的名称
  for (let i = 1; i <= sheets.Count; i++) {
      workbook[i-1] = (sheets.Item(i).Name)
      // console.log(workbook[i-1])
  }
}

// 判断表是否已存在
function workbookComp(name){
  let flag = 0;
  let length = workbook.length
  for(let i = 0; i < length; i++){
    if(workbook[i] == name){
      flag = 1;
      console.log(name + "表已存在")
      break
    }
  }
  return flag
}

// 创建表，若表已存在则不创建，直接写入数据
function createSheet(name){
  // const defaultName = Application.Sheets.DefaultNewSheetName
  // 工作表对象
  if(!workbookComp(name)){
    Application.Sheets.Add(
        null,
        Application.ActiveSheet.Name,
        1,
        Application.Enum.XlSheetType.xlWorksheet,
        name
    )
  }
}
