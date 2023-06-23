# ExcelTools

入口文件:`ExcelTools.py`

主要功能：将excel文件中的原始订单信息进行统计，并转化成订单-颜色-尺码维度的数量统计报表文件

## 项目信息
语言及版本： python3， 建议使用python3.11
使用的库： customtkinter, openpyxl, pyinstaller


## 打包：windows打包exe

### 1. 安装pyinstaller
```bash
pip install pyinstaller
```

### 2. 查看customtkinter库路径: <CustomTkinter Location>
```bash
pip show customtkinter
```
A Location will be shown, for example: c:\users\<user_name>\appdata\local\programs\python\python310\lib\site-packages

Then add the library folder like this: --add-data "C:/Users/<user_name>/AppData/Local/Programs/Python/Python310/Lib/site-packages/customtkinter;customtkinter/"

### 3. 打包
```bash
pyinstaller --noconfirm --onedir --windowed --add-data "<CustomTkinter Location>/customtkinter;customtkinter/"  "<Path to Python Script>"
```
<Path to Python Script>: 项目的入口文件路径，如：E:\code\python\ExcelTools.py

完整示例：
```bash
pyinstaller --noconfirm --onedir --windowed --add-data "C:\Users\Administrator\AppData\Local\Programs\Python\Python311\Lib\site-packages\customtkinter;customtkinter\" "E:\code\python\ExcelTools.py"
```

## 4. 打包后的exe文件在dist目录下

> 参考资料：
> [customtkinter官方打包说明](https://customtkinter.tomschimansky.com/documentation/packaging)
> [customtkinter Documentation](https://customtkinter.tomschimansky.com/documentation/)