# save-json-as-excel   

### save json as xls or csv in browser , and download . (在浏览器端将json转成xls/csv并下载)

> ## example
```javascript

  var dataSource = [{
        name: '海绵宝宝',
        age: 31,
        address: '贝壳街124号'
      }, {
        name: '蟹老板',
        age: 75,
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
      }];

  saveJsonAsExcel.saveAsXls(dataSource, columns, 'test.xls');
  saveJsonAsExcel.saveAsCsv(dataSource, columns, 'test.csv');

```
