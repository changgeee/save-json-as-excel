# JSONTOEXCEL
> ## 使用说明
+ ### 提供exportXLS(data:array,[fileName:string,header:object])和exportCSV(data:array,[fileName:string,header:array])两个方法；
+ ##### exportXLS中header参数用于指定表头和需要生成到excel中的字段,缺省时无表头，遍历添加每条数据中所有字段;
+ ##### exportCSV中header参数用于指定生成到excel中字段，缺省时遍历添加每天数据中所有字段;
+ ### example

```javascript
  var data = [{
          name: 'nico', age: '18' 
        }, {
          name: 'changge', age: '20'
      }]
  var xlsheader = {
        name: '姓名',
        age: '年龄'
      }
  var csvhearder=['name','age']
  JSONTOEXCEL.exportXLS(data, 'test.xls', header)
  JSONTOEXCEL.exportCSV(data, 'test.csv')
```

 