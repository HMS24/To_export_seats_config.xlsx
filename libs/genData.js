const { getColAndRow } = require('./func')

// 解析 json 生成 cells object
const generateData = (dataOfJson, display) => {

  const data = dataOfJson
    .map(item => {
      const { col, row } = getColAndRow(item.position)

      // 返回依 display 順序顯示的欄位
      return display.map((key, index) => {
        const position = col + (row + index)

        // 返回 worksheet 的結構
        return { [position]: { v: item[key] } }
      })
    })
    // 降維
    .reduce((prev, curr) => {
      return prev.concat(curr)
    })
    .reduce((prev, curr) => {
      return Object.assign({}, prev, curr)
    })

  return data
}

module.exports = generateData