/*======================
  基本設定
========================*/
const XLSX = require('xlsx-style')

const { genTemplate, genData } = require('./libs/generate')
const { xlsxToJson, getRange, merge } = require('./libs/func')

// 宣告變數
const display = ['no', 'company', 'job', 'name']
const config = {
  header: {
    value: '第70屆金鐘獎頒獎典禮工作會議',
    position: 'E2:E2'
  },
  location: {
    value: '會議室座位表',
    position: 'J4:J4'
  },
  door1: {
    value: '門',
    position: 'B37:C38'
  },
  door2: {
    value: '門',
    position: 'V37:W38'
  },
  lectern: {
    value: '講台',
    position: 'I37:N38'
  }
}

/*======================
  執行
========================*/

// 讀取 Excel 匯出 worksheets
const workbook = XLSX.readFile('./excels/conference.xlsx', { cellStyles: true })
const worksheet = workbook.SheetNames
const _seats = workbook.Sheets[worksheet[0]]
const _students = workbook.Sheets[worksheet[1]]

// 處理 worksheets（過濾 _seats to array；轉換 _students to JSON）
const keys = Object.keys(_seats)
const seats = keys.filter(key => {
  return (key.charAt(0) !== '!' && _seats[key].v)
})
const { content } = xlsxToJson(_students)

// 處理使用者需求並合併為 worksheet object
const template = genTemplate(config, seats, display.length - 1)
const data = genData(content, display)
const output = merge(template, data)

// 範圍
const ref = getRange(output)

// 生成 workbook
const wb = {
  SheetNames: ['sheet_1'],
  Sheets: {
    'sheet_1': {
      '!ref': ref, // 若無範圍，則只會產出空表而已
      ...output
    }
  }
}

// 生成 Excel 文件
XLSX.writeFile(wb, './excels/output.xlsx')