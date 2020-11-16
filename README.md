# doExcel

## 基于python编写的简单excel文件数据梳理整理工具

## 使用方法

> 1.将要处理汇总的xls文件放入项目根目录
> 2.修改success.py文件中的文件名字为自己文件名称
> 3.注意格式和提供的模板样例文件尽量一致
> 4.执行success.py文件（注意python环境和对应依赖包的导入）

```python
if __name__ == '__main__':
-   filename = "20201112.xls"
   #修改为自己的文件名
+   filename = "[自己的文件名]"
    extract(filename)
    write_excel(list_data,filename)
    print('更新文件成功')
```

## 专门给你写的，瞅啥呢？还不搞起来~
