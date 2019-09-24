# excel-email-to-json

# Usage
```
npm install excel-email-to-json
```

```
const ExcelEmailToJson = require('excel-email-to-json);

/**
* @params {filename}  读取Excel文件的文件名，绝对路径
* @return ExcelEmailToJson类的实例
/
const data = new ExcelEmailToJson(filename);
data.result // 返回所有的email的数组

/**
* @params {filename} // 写入的文件名，绝对路径
* @params {row} // 拆分文件的行数，默认为0，不拆分
/
data.writeFile(filename, row);

```