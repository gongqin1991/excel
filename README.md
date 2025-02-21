# excel组件
基于[excelize](https://github.com/xuri/excelize)组件,封装excel文件读写操作.
## How to use
> go get -u github.com/gongqin1991/excel

### 一个简单的例子：
```
w := excel.NewWriter("abc.xlsx", "自定义工作表名")
w.WriteHeader([]string{"列1", "列2"})
w.WriteRow([]string{"1", "2"}, excel.ROW+1)
w.SaveTo()
if err := w.Err(); err != nil {
    panic(err)
}
```

### 多sheet写入：
#### 方法1：
```
w := excel.NewWriter("abc.xlsx", "自定义工作表名")
w.WriteHeader([]string{"列1", "列2"})
w.WriteRow([]string{"1", "2"}, excel.ROW+1)

w1 := excel.OpenWriter(w, "自定义工作表名1")
w1.WriteHeader([]string{"列1", "列2"})
w1.WriteRow([]string{"1", "2"}, excel.ROW+1)

w.SaveTo()
if err := w.Err(); err != nil {
    panic(err)
}
```
#### 方法2：
```
wr := excel.NewWriter2("abc.xlsx")

w1 := excel.OpenWriter(wr, "自定义工作表名")
w1.WriteHeader([]string{"列1", "列2"})
w1.WriteRow([]string{"1", "2"}, excel.ROW+1)

w2 := excel.OpenWriter(wr, "自定义工作表名1")
w2.WriteHeader([]string{"列1", "列2"})
w2.WriteRow([]string{"1", "2"}, excel.ROW+1)

wr.SaveTo()
if err := w.Err(); err != nil {
    panic(err)
}
```

