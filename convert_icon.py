from PIL import Image
import os

def create_icon():
    try:
        # 支持的图片格式
        image_formats = ['.png', '.jpg', '.jpeg']
        
        # 查找可用的图片
        for fmt in image_formats:
            image_name = 'xdlovelife' + fmt
            if os.path.exists(image_name):
                # 打开图片
                img = Image.open(image_name)
                
                # 调整大小为标准图标尺寸
                icon_sizes = [(16,16), (32,32), (48,48), (64,64), (128,128)]
                img.save('xdlovelife.ico', format='ICO', 
                        sizes=icon_sizes)
                
                print(f"图标已成功创建：xdlovelife.ico")
                return True
                
        print("未找到可用的源图片文件")
        return False
        
    except Exception as e:
        print(f"创建图标时发生错误: {str(e)}")
        return False

if __name__ == "__main__":
    create_icon() 