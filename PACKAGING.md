# 打包说明

## 已生成内容

- `dist/会议信息录入助手/`：可直接运行的程序目录。
- `release/会议信息录入助手-便携版.zip`：便携版压缩包，解压后运行 `会议信息录入助手.exe`。
- `installer/会议信息录入助手.iss`：Inno Setup 安装包脚本。

## 重新生成 exe

```powershell
python -m PyInstaller --noconfirm --clean --windowed --name "会议信息录入助手" --hidden-import win32com.client --hidden-import pythoncom --hidden-import pywintypes --exclude-module numpy --exclude-module pandas --exclude-module PIL --exclude-module lxml --exclude-module matplotlib --exclude-module scipy meeting_assistant.py
Copy-Item -LiteralPath '会议安排.xlsx' -Destination 'dist\会议信息录入助手\会议安排.xlsx' -Force
Copy-Item -LiteralPath 'README.md' -Destination 'dist\会议信息录入助手\README.md' -Force
```

## 生成安装包

安装 Inno Setup 6 后，右键打开 `installer/会议信息录入助手.iss` 并点击 Compile，即可生成 `会议信息录入助手-安装包.exe`。

也可以用命令行：

```powershell
& "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" "installer\会议信息录入助手.iss"
```

## 使用说明

安装版和便携版均不要求对方安装 Python。若需要实时写入已打开的 Excel，需要对方电脑安装 Microsoft Excel；如果 Excel 未打开，程序也可以直接后台写入 `.xlsx` 文件。
