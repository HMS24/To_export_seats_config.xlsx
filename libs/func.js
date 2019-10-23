// 取得 cell 行列如 A1 => {A,1}
const getColAndRow = (cell) => {
  // col 位數
  let index = cell.split('').findIndex(value => {
    return !isNaN(value)
  })

  // 如 A21 中的 A
  const col = cell.substr(0, index)
  // 如 A21 中的 21
  const row = parseInt(cell.substr(index))

  return { col, row }
}
// 解析 xlsx 返回 {headers, data}
const xlsxToJson = (worksheet) => {
  const _headers = {}
  const content = []
  const keys = Object.keys(worksheet)

  keys
    // 過濾非'!'開頭的 keys
    .filter(k => k.charAt(0) !== '!')
    .forEach(k => {
      const { col, row } = getColAndRow(k)
      // 當前 cell 的值
      const value = worksheet[k].v

      // 第一列為 headers
      if (row === 1) {
        _headers[col] = value
        return
      }

      // 解析 data
      const index = row - 2

      if (!content[index]) {
        content[index] = {}
      }

      content[index][_headers[col]] = value
    })

  // 
  const headers = Object.values(_headers)

  return { headers, content }
}
// 畫外框線，如 A1:C3 的正方形外框
const drawFieldsBorder = (start, end) => {

  const startCR = getColAndRow(start) // {A, 1}
  const endCR = getColAndRow(end) // {C, 3}
  const startCol = startCR.col // A
  const startRow = startCR.row // 1
  const endCol = endCR.col // C
  const endRow = endCR.row // 3
  const startChar = start.charCodeAt() // 65
  const endChar = end.charCodeAt() //67
  const result = {}

  /* 外框選項 
  {
    A:'left', 
    C:'right', 
    '1':'top', 
    '3':'bottom'
  }
  */
  const opts = {
    [startCol]: 'left',
    [endCol]: 'right',
    [startRow]: 'top',
    [endRow]: 'bottom'
  }

  for (let c = startChar; c <= endChar; c++) {
    for (let r = startRow; r <= endRow; r++) {
      const col = String.fromCharCode(c)
      const row = r
      const cell = col + row
      const border = {}

      if (opts[col]) {
        border[opts[col]] = { style: 'medium' }
      }

      if (opts[row]) {
        border[opts[row]] = { style: 'medium' }
      }

      // 如僅有單一欄則需畫左邊框線
      if (startCol === endCol) {
        border['left'] = { style: 'medium' }
      }

      // 無邊框 cells 就不需要 s 屬性
      if (Object.keys(border).length === 0) {
        result[cell] = { v: '' }
      } else {
        result[cell] = { s: { border }, v: '' }
      }

    }
  }
  return result
}
// 取得 worksheet 範圍 返回如 'A1:G8'
const getRange = (worksheet) => {
  const cellsArr = Object.keys(worksheet)

  // 行
  const cols = cellsArr.map(cell => {
    return getColAndRow(cell).col
  }).sort()

  const startCol = cols[0]
  const endCol = cols[cols.length - 1]

  // 列
  const rows = cellsArr.map(cell => {
    return getColAndRow(cell).row
  }).sort((a, b) => a - b)

  const startRow = rows[0]
  const endRow = rows[rows.length - 1]

  return `${startCol + startRow}:${endCol + endRow}`
}
// 合併 template 及 data
const merge = (template, data) => {
  const dataCells = Object.keys(data)
  const error = []

  dataCells.forEach(cell => {

    if (!template[cell]) {
      error.push(data[cell])
    } else {
      template[cell].v = data[cell].v
    }

  })

  if (error.length) return error
  return template
}

module.exports = {
  getColAndRow,
  xlsxToJson,
  drawFieldsBorder,
  getRange,
  merge
}