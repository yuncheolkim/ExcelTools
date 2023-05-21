F:\tools\excel.exe -format=f -inDir D:\Config -outDir D:\Data1\


- 第一行注释
- 第二行变量名
- 第三行类型,支持类型为:int,long,float,string,json
- 第一列必须为id列
- 第一列为 # 字符则跳过这行
- sheet名字test,temp开头则跳过读取
- format=f,不格式化,去掉这个选项则格式化