# excel
单页面导出

##### usage
```go
	p, err := NewPage() //新建页面
	if err != nil {
		panic(err)
	}
	defer p.Close()
	p.SetTitle("name", "phone") //设置标题
	p.PushRow("li", "134", "29") //添加一行行的内容
	_ = p.Save("ppxxx") //保存文件

