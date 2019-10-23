# Generate Excel
學習於 Node.js 使用套件 xlsx-style 處理 Excel 文件。依據使用者需求匯出座位表，以方便清楚查看會議的配置。

## Features
1. 使用者可以匯出 Excel 座位表

## Preview Pages
<img src="./public/imgs/output.jpg" width="400px" target="_blank">

## Environment and package used
* [Node.js](https://nodejs.org/en/) v10.15.0

## Package used
* [xlsx-style](https://www.npmjs.com/package/xlsx-style) v0.8.13

## Installation and usage
#### 複製專案
```git
git clone https://github.com/HuangMinShi/export_xlsx.git
```

#### 切換專案
```git
cd export_xlsx
```

#### 安裝套件
```npm
npm i xlsx-style
```
 
#### 啟動伺服器
```npm
npm run dev
```
#### 生成 Excel
所產生的 Excel 檔案為 public/excels/output.xlsx

## Sth need to improve
- 使用者於出席者清單中指定錯誤的座位欄位時，所匯出 Excel 將無法開啟。