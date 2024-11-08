import os
import shutil

def optimize_dist():
    """优化dist文件夹大小"""
    dist_path = "dist/阿里巴巴数据邮箱处理工具"
    
    # 要删除的不必要文件
    unnecessary_files = [
        '_internal/tcl',
        '_internal/tk',
        '_internal/test',
        '_internal/pandas/tests',
        '_internal/numpy/tests',
    ]
    
    for file_path in unnecessary_files:
        full_path = os.path.join(dist_path, file_path)
        if os.path.exists(full_path):
            if os.path.isdir(full_path):
                shutil.rmtree(full_path)
            else:
                os.remove(full_path)
            print(f"已删除: {file_path}")

if __name__ == "__main__":
    optimize_dist() 