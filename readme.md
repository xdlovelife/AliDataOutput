# 数据提取自动化工具
![27b0358e-b93d-4b96-b0d3-5e7cbf3f2758](https://github.com/user-attachments/assets/d2fdccd5-60f1-4278-abfb-462f0cb964fe)

## 项目简介

本项目是一个基于 Selenium 的自动化工具，用于从指定网站提取数据并保存到 Excel 文件中。该工具支持输入多个搜索项，并能够处理搜索结果。用户可以通过 Excel 文件输入要搜索的内容，程序将自动执行搜索并提取相关信息。

## 环境要求

- Python 3.x
- 安装以下 Python 库：
  - `selenium`
  - `pandas`
  - `openpyxl`
  - `requests`
  - `tkinter`
  - `xlrd`
  - `win32com`
  - `psutil`

您可以使用以下命令安装所需库：

pip install selenium pandas openpyxl requests xlrd pywin32 psutil



## 使用步骤

1. **配置文件**：
   - 在项目根目录下，创建一个名为 `app_config.json` 的配置文件，包含以下内容：
     ```json
     {
         "excel_path": "path/to/your/excel/file.xlsx",
         "account": "your_account",
         "password": "your_password",
         "driver_path": "path/to/your/chromedriver.exe"
     }
     ```
   - 确保 `excel_path` 指向您要读取和写入的 Excel 文件。

2. **下载 ChromeDriver**：
   - 确保您已安装 Chrome 浏览器，并下载与您的 Chrome 版本匹配的 ChromeDriver。
   - 将 ChromeDriver 放置在配置文件中指定的路径。

3. **运行程序**：
   - 在命令行中导航到项目目录并运行：
     ```bash
     python main.py
     ```

4. **输入搜索项**：
   - 程序启动后，您可以在 Excel 文件中输入要搜索的项（D 列），程序将自动读取这些项并进行搜索。

5. **处理验证码**：
   - 在某些情况下，程序可能会遇到验证码。此时，程序会暂停并提示您手动输入验证码。
   - 请确保在输入验证码后，程序能够继续执行。

## 注意事项

- **验证码处理**：
  - 当程序检测到需要输入验证码时，会暂停并提示用户手动输入验证码。
  - 确保在输入验证码后，程序能够继续执行。您可能需要在代码中添加适当的等待时间，以确保验证码输入后页面能够正常加载。

- **网络连接**：
  - 确保您的网络连接稳定，以避免在搜索过程中出现超时或加载失败的情况。

- **元素定位**：
  - 如果程序在定位某些元素时失败，请检查网页结构是否发生变化，并相应地更新 XPath。

- **异常处理**：
  - 程序中包含了基本的异常处理机制，确保在出现错误时能够记录日志并继续执行。

## 常见问题

- **程序无法启动**：
  - 确保您已正确安装所有依赖库，并且 Python 环境配置正确。

- **搜索结果未加载**：
  - 检查网络连接，并确保搜索页面能够正常访问。

- **验证码输入后程序未继续**：
  - 确保在输入验证码后，程序能够继续执行。您可能需要在代码中添加适当的等待时间。

## 贡献

欢迎任何形式的贡献！如果您发现了错误或有改进建议，请提交问题或拉取请求。

## 许可证

本项目采用 MIT 许可证，详细信息请参见 LICENSE 文件。
