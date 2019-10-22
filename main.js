// 引入套件及函式
const XLSX = require('xlsx-style')

const generateTemplate = require('./libs/genTemplate')
const generateData = require('./libs/genData')
const { xlsxToJson, getRange, merge } = require('./libs/func')

// 宣告相關變數
const display = ['no', 'office', 'title', 'name']
const config = {
  header: {
    value: '第3期司法院暨所屬會計人員會計業務專業講習',
    position: 'E2:E2'
  },
  location: {
    value: '401教室座位表',
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

// 讀取 xlsx 匯出 worksheets
const workbook = XLSX.readFile('./excels/conference.xlsx')
const worksheet = workbook.SheetNames
const _seats = workbook.Sheets[worksheet[0]]
const _students = workbook.Sheets[worksheet[1]]

// 處理 worksheets（過濾 _seats；轉換 _students）
const keys = Object.keys(_seats)
const seats = keys.filter(key => {
  return (key.charAt(0) !== '!' && _seats[key].v)
})
const { content } = xlsxToJson(_students)

// 轉換成 worksheet 特定結構
const template = generateTemplate(config, seats, display.length - 1)
const data = generateData(content, display)
const output = merge(template, data)

// 範圍
const ref = getRange(output)

// 生成 workbook
const wb = {
  SheetNames: ['sheet_1'],
  Sheets: {
    'sheet_1': {
      '!ref': ref,
      ...output
    }
  }
}

// 生成 xlsx 檔案
XLSX.writeFile(wb, './excels/output.xlsx')