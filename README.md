# excel组件
基于[excelize](https://github.com/xuri/excelize)组件,封装excel文件读写操作.
## How to use
> go get -u github.com/gongqin1991/excel

### 一个简单的例子：
```
w := excel.NewWriter("abc.xlsx", "自定义工作表名")
w.WriteHeader([]string{"列1", "列2"})
w.WriteColumns([]string{"1", "2"}, 2)
w.SaveTo()
if err := w.Err(); err != nil {
    panic(err)
}
```

