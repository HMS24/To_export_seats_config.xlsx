const { getColAndRow, drawFieldsBorder } = require('./func')

/* 解析 JSON 生成 data 如
{
  P8: { v: 1 },
  P9: { v: '喬傑立娛樂' },
  P10: { v: '歌手' },
  P11: { v: '孫協志' },
  T6: { v: 2 },
  T7: { v: '三立藝能' },
  T8: { v: '演員' },
  T9: { v: '紀言愷' },
}
*/

const genData = (json, display) => {

  const data = json
    .map(item => {
      const { col, row } = getColAndRow(item.position)

      // 返回依 display 順序顯示的欄位
      return display.map((key, index) => {
        const position = col + (row + index)

        // 返回 worksheet object
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

/* 生成 template 如
{
  E2: { v: '第70屆金鐘獎頒獎典禮工作會議', s: { font: [Object] } },
  J4: { v: '會議室座位表', s: { font: [Object] } },
  B37: { s: { border: [Object], font: [Object] }, v: '門' },
}
*/

const genTemplate = (config, seats, cellsPerSeat) => {

  // 畫出配置
  const c = Object.keys(config)
    .map(key => {
      const item = config[key]
      const start = item.position.split(':')[0]
      const end = item.position.split(':')[1]

      // 門及講台的儲存格畫框線
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

module.exports = {
  genTemplate,
  genData
}