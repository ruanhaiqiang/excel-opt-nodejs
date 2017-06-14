# excel-opt-nodejs

## 说明
该module主要用于excel解析


## Install

```
npm install excel-opt-nodejs
```

## Usage

```
var excel-opt = require("excel-opt-nodejs");
var ExcelOpt = new excel-opt();
```

### loadExcel

```
var workbook = ExcelOpt.readFile("./test.xlsx");
```