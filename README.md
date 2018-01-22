# JSONTOEXCEL
> ## 使用说明
+ 用于浏览器端将json格式数据转换为 伪xls/csv 格式文件导出；

> ## API

+ JSONTOEXCEL.exportXLS(dataSource:Array\<any\>, columns:Array<{title:string,dataIndex?:string,computed?:function}>,fileName:string)

+ JSONTOEXCEL.exportCSV(dataSource:Array\<any\>, columns:Array<{dataIndex?:string,computed?:function}>,fileName:string)


> ## example
```javascript
  var dataSource = [{
        name: '胡彦斌',
        age: 32,
        address: '西湖区湖底公园1号'
      }, {
        name: '胡彦祖',
        age: 42,
        address: '西湖区湖底公园1号'
      }],

      columns = columns = [{
        title: '姓名',
        dataIndex: 'name'
      }, {
        title: '年龄',
        dataIndex: 'age',
      }, {
        title: '住址',
        dataIndex: 'address'
      }, {
        title: '出生年份',
        computed: function (item) {
          return 2017 - item.age
        }
      }]

  JSONTOEXCEL.exportXLS(dataSource, columns, 'test.xls')

  JSONTOEXCEL.exportCSV(dataSource, columns, 'test.csv')
```

> ## change log

+ v 0.1.0
  + feature 支持计算属性 computed；
  + update api优化；