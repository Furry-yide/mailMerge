# mailMerge

通过读取指定的 Excel 以及 Word 文件，将 Excel 中的标题行下的字段替换到 Word 文件中。这对于不会使用 WPS 或者 Office 中自带的邮件合并的人群，有更相对方便的操作

# 运行

<h2>安装依赖</h2>

```bash
pip install -r requirements.txt
```

<h2>运行</h2>

```bash
python3 main.py
```

运行成功后，会在当前目录生成一个 `build` 文件夹，
当测试的文件成功生成在 `build` 文件夹内，那么就证明你已经部署成功了

# 如何使用

## 前言

在开始使用前，先了解一下该程序的文件树结构：

```bash
mailMerge/
├── build/
├── data/
│    ├── dataBase.xlsx
│    └── template.docx
├── main.py
└── README.md
```

- `build` 是已经生成好的文件
- `dataBase.xlsx` 是数据源文件，其中包含了大量的信息
- `template.docx` 是模板文件，其中包含了大量的占位符，占位符的格式为 `{字段名}`
- `main.py` 是程序入口文件，运行该文件即可完成文件的生成
- `README.md` 是本 README 文件

> **注意！:**\
> `template.docx`中的占位符与`dataBase.xlsx`中的字段名必须保持一致，否则替换不会成功

## 如何使用

将 `dataBase.xlsx` 与 `template.docx` 替换成你已经准备好的数据源文件以及模版文件，然后运行 `main.py`

打开 `build` 文件夹，查看已生成的文件

## 进阶使用

如果你并不想在 `data` 文件夹中寻找你的数据源文件，那么你可以在 `main.py` 中修改 `readExcel()` 的传参，将路径替换成你想要的数据源文件目录

操作如下:

```python
# readExcel可以传参 dataBasePath="data/dataBase.xlsx" 数据源路径
headerData , studentData = readExcel(dataBasePath="data/dataBase.xlsx") # 读取Excel文件
```

修改模版路径也同理:

```python
# buildWord可以传参 templatePath="data/template.docx" 模版路径
buildWord(excelHeader=headerData, data=studentData, templatePath="data/template.docx" ) # 生成Word文件
```

其中，在 `buildWord()` 变量里还预留了一些参数可以传入，用于控制一些其他的配置:
| 参数名称 | 作用 | 默认值 | 备注 |
|---|---|---|---|
|buildPath |当文件生成后所在的文件夹|build/|
|buildName |生成后的文件名|file|
|templatePath |模版所在的路径|data/template.docx|
|type |是否生成多个文件|False|是否生成多个文件,当 False 将生成一个文件,当 True 为生成多个文件|

# 其他
如果有字段无法替换，请检查你的数据源文件与模版文件中的字段是否一致