
# windows打包exe

## 1. 安装pyinstaller
```bash
pip install pyinstaller
```

## 2. 查看customtkinter库路径
```bash
pip show customtkinter
```
A Location will be shown, for example: c:\users\<user_name>\appdata\local\programs\python\python310\lib\site-packages

Then add the library folder like this: --add-data "C:/Users/<user_name>/AppData/Local/Programs/Python/Python310/Lib/site-packages/customtkinter;customtkinter/"

## 3. 打包
```bash
pyinstaller --noconfirm --onedir --windowed --add-data "" "项目入口脚本路径"
```


pyinstaller --noconfirm --onedir --windowed --add-data "C:\Users\Administrator\AppData\Local\Programs\Python\Python311\Lib\site-packages\customtkinter;customtkinter\" "E:\code\python\ExcelTools.py"