# save-json-as-excel

> save json as xls or csv in browser , and download . (在浏览器端将 json 转成 xls/csv 并下载)

#### install

```sh
  npm install save-json-as-excel --save
```

#### example

```javascript
var dataSource = [
    {
      name: '海绵宝宝',
      age: 31,
      address: '贝壳街124号',
    },
    {
      name: '蟹老板',
      age: 75,
      address: '船锚路3541号',
    },
  ],
  columns = (columns = [
    {
      title: '姓名',
      dataIndex: 'name',
    },
    {
      title: '年龄',
      dataIndex: 'age',
    },
    {
      title: '住址',
      dataIndex: 'address',
    },
    {
      title: '出生年份',
      computed: function(item) {
        return 2017 - item.age;
      },
    },
  ]);

saveJsonAsExcel.saveAsXls(dataSource, columns, 'test.xls');
saveJsonAsExcel.saveAsCsv(dataSource, columns, 'test.csv');
```
