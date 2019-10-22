const { getColAndRow, drawFieldsBorder } = require('./func')

// 生成 template
const generateTemplate = (config, seats, cellsPerSeat) => {
  // 畫出配置
  const c = Object.keys(config)
    .map(key => {
      const item = config[key]
      const start = item.position.split(':')[0]
      const end = item.position.split(':')[1]

      // 門及講台的 cell 畫框線及賦值
      if (key === 'door1' || key === 'door2' || key === 'lectern') {
        const fieldsHasBorder = drawFieldsBorder(start, end)
        fieldsHasBorder[start].v = item.value
        fieldsHasBorder[start].s['font'] = { sz: '16', bold: true }
        return fieldsHasBorder
      }

      return {
        [start]: {
          v: item.value,
          s: {
            font: {
              sz: '24',
              bold: true
            }
          }
        }
      }

    })
    .reduce((prev, curr) => {
      return Object.assign({}, prev, curr)
    })

  // 畫出座位
  const s = seats
    .map(start => {
      const { col, row } = getColAndRow(start)
      const end = col + (row + cellsPerSeat)

      return drawFieldsBorder(start, end)
    })
    .reduce((prev, curr) => {
      return Object.assign({}, prev, curr)
    })

  // 返回合併後的配置及座位
  return Object.assign({}, c, s)
}

module.exports = generateTemplate