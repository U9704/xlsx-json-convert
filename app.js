const express = require('express')
const fs = require('fs')
const json2xls = require('json2xls')
const xlsx = require('node-xlsx')

const app = express()

app.listen(3054, () => {
  console.log('服务启动')
})

/**
 * txt转xlsx
 */
app.get('/txtToXlsx', (req, res) => {
  const filePath = req.query.inputPath
  const outputFilePath = req.query.outputPath
  if (!filePath || !outputFilePath) {
    res.json('操作失败，参数inputPath、outputPath必传')
    return
  }
  const jsonData = fs.readFileSync(filePath, 'utf-8')
  const lines = jsonData.split(/\r?\n/)
  const xlsxData = []
  for (let i = 0; i < lines.length; i++) {
    const obj = JSON.parse(lines[i])
    xlsxData.push(obj)
  }
  let xls = json2xls(xlsxData)
  fs.writeFileSync(outputFilePath, xls, 'binary')
  res.json('操作成功')
})

/**
 * xlsx转txt
 */
app.get('/xlsxToTxt', async (req, res) => {
  const inputPath = req.query.inputPath
  const outputPath = req.query.outputPath
  if (!inputPath || !outputPath) {
    res.json('操作失败，参数inputPath、outputPath必传')
    return
  }
  const sheets = xlsx.parse(inputPath)
  const sheetData = sheets[0].data
  const objKeys = sheetData[0] // 取出表头
  for (let i = 1; i < sheetData.length; i++) {
    const dataItem = sheetData[i]
    const obj = {}
    for (let j = 0; j < objKeys.length; j++) {
      obj[objKeys[j]] = dataItem[j] ? dataItem[j] + '' : ''
    }
    await fs.appendFileSync(outputPath, JSON.stringify(obj) + '\r\n')
  }
  res.json('操作成功')
})
