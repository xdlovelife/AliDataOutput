import PyInstaller.__main__
import os
import sys

def build():
    # 确保图标文件存在
    if not os.path.exists('xdlovelife.ico'):
        print("错误: 未找到图标文件 xdlovelife.ico")
        return

    # 打包参数
    params = [
        'test.py',                          # 主程序
        '--name=阿里巴巴数据处理工具',       # 程序名称
        '--onefile',                        # 打包成单个文件
        '--windowed',                       # 无控制台窗口
        '--icon=xdlovelife.ico',           # 程序图标
        '--add-data=xdlovelife.ico;.',     # 添加图标文件
        '--hidden-import=pandas',           # 添加隐藏导入
        '--hidden-import=pandas._libs.tslibs.base',
        '--hidden-import=pandas._libs.tslibs.timedeltas',
        '--hidden-import=pandas._libs.tslibs.timestamps',
        '--hidden-import=pandas._libs.tslibs.offsets',
        '--hidden-import=pandas._libs.tslibs.parsing',
        '--hidden-import=pandas._libs.tslibs.fields',
        '--hidden-import=pandas._libs.tslibs.conversion',
        '--hidden-import=pandas._libs.tslibs.nattype',
        '--hidden-import=pandas._libs.tslibs.np_datetime',
        '--hidden-import=pandas._libs.tslibs.strptime',
        '--hidden-import=pandas._libs.tslibs.period',
        '--hidden-import=pandas._libs.tslibs.ccalendar',
        '--hidden-import=pandas._libs.tslibs.vectorized',
        '--collect-all=selenium',           # 收集所有相关文件
        '--collect-all=pandas',
        '--collect-all=openpyxl',
        '--collect-all=xlwt',
        '--collect-all=webdriver_manager',
        '--clean',                          # 清理临时文件
        '--noconfirm',                      # 不确认覆盖
    ]

    # 执行打包
    PyInstaller.__main__.run(params)

if __name__ == "__main__":
    build() 