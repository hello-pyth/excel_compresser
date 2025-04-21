from spire.xls import *
from spire.xls.common import *
import os
from PIL import Image
import io
import tempfile
import shutil
import uuid
from pathlib import Path
import win32com.client
import time
import sys
import pythoncom

def get_image_data_direct(pic):
    """直接通过二进制方式获取图片数据"""
    try:
        # 尝试直接访问底层数据
        img_data = pic.Picture.Data
        if img_data and len(img_data) > 0:
            return img_data
    except:
        pass
    
    # 如果上面的方法失败，使用备用方法
    try:
        return pic.Picture.Raw
    except:
        pass
    
    return None

def extract_images_with_win32com(input_file, output_dir):
    """使用win32com提取图片"""
    print("尝试使用win32com提取图片...")
    
    try:
        # 初始化COM
        pythoncom.CoInitialize()
        
        # 创建Excel应用
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        
        # 打开文件
        workbook = excel.Workbooks.Open(os.path.abspath(input_file))
        
        image_paths = []
        
        # 遍历所有工作表
        for i in range(1, workbook.Sheets.Count + 1):
            sheet = workbook.Sheets(i)
            print(f"处理工作表: {sheet.Name}")
            
            try:
                # 尝试通过Shape对象获取图片
                if sheet.Shapes.Count > 0:
                    for j in range(1, sheet.Shapes.Count + 1):
                        try:
                            shape = sheet.Shapes(j)
                            
                            # 只处理图片
                            if shape.Type == 13:  # 13 = msoPicture
                                # 生成唯一文件名
                                temp_filename = f"excel_img_{uuid.uuid4()}"
                                img_path = os.path.join(output_dir, f"{temp_filename}.png")
                                
                                # 使用CopyPicture方法
                                shape.CopyPicture()
                                # 获取剪贴板内容
                                import win32clipboard
                                win32clipboard.OpenClipboard()
                                if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_DIB):
                                    img_data = win32clipboard.GetClipboardData(win32clipboard.CF_DIB)
                                    win32clipboard.CloseClipboard()
                                    
                                    # 保存图片
                                    with open(img_path, 'wb') as f:
                                        f.write(img_data)
                                    
                                    if os.path.exists(img_path) and os.path.getsize(img_path) > 0:
                                        print(f"通过剪贴板导出图片: {img_path}")
                                        image_info = {
                                            'path': img_path,
                                            'sheet': sheet.Name,
                                            'left': shape.Left,
                                            'top': shape.Top,
                                            'width': shape.Width,
                                            'height': shape.Height
                                        }
                                        image_paths.append(image_info)
                                else:
                                    win32clipboard.CloseClipboard()
                                    print(f"剪贴板中没有图片数据")
                        except Exception as e:
                            print(f"导出Shape {j} 时出错: {str(e)}")
                            continue
            except:
                print(f"工作表 {sheet.Name} 没有Shapes集合")
        
        # 关闭文件
        workbook.Close(False)
        excel.Quit()
        
        return image_paths
    
    except Exception as e:
        print(f"使用win32com提取图片时出错: {str(e)}")
        return []
    finally:
        # 释放COM
        pythoncom.CoUninitialize()

def optimize_image(input_path, output_path, max_size_kb=300):
    """优化图片大小，确保不超过指定大小"""
    try:
        with Image.open(input_path) as img:
            # 计算合适的尺寸
            max_width = 800
            max_height = 600
            width, height = img.size
            
            # 如果图片太大，按比例缩小
            if width > max_width or height > max_height:
                ratio = min(max_width/width, max_height/height)
                new_size = (int(width*ratio), int(height*ratio))
                img = img.resize(new_size, Image.Resampling.LANCZOS)
            
            # 保存优化后的图片
            quality = 70
            if img.mode in ('RGBA', 'LA') or (img.mode == 'P' and 'transparency' in img.info):
                # PNG格式
                img.save(output_path, format='PNG', optimize=True, compression_level=9)
            else:
                # JPEG格式
                img = img.convert('RGB')
                while quality > 10:
                    buffer = io.BytesIO()
                    img.save(buffer, format='JPEG', quality=quality, optimize=True, progressive=True)
                    if buffer.tell() <= max_size_kb * 1024 or quality <= 10:
                        img.save(output_path, format='JPEG', quality=quality, optimize=True, progressive=True)
                        break
                    quality -= 10
            
            return True
    except Exception as e:
        print(f"优化图片时出错: {str(e)}")
        return False

def compress_excel_file(input_file, output_file, target_size_ratio=0.5):
    """使用多种方法压缩Excel文件中的图片"""
    # 创建临时目录
    temp_dir = tempfile.mkdtemp()
    compressed_dir = os.path.join(temp_dir, "compressed")
    os.makedirs(compressed_dir, exist_ok=True)
    
    print(f"临时目录: {temp_dir}")
    
    try:
        # 获取原始文件大小
        original_size = os.path.getsize(input_file)
        target_size = original_size * target_size_ratio
        print(f"目标文件大小: {target_size / (1024*1024):.2f} MB")
        
        # 先尝试使用win32com提取图片
        images = extract_images_with_win32com(input_file, temp_dir)
        
        if not images:
            print("无法使用win32com提取图片，尝试备用方法...")
            # 备用方法：使用Spire.XLS
            workbook = Workbook()
            workbook.LoadFromFile(input_file)
            
            for sheet_index in range(workbook.Worksheets.Count):
                sheet = workbook.Worksheets[sheet_index]
                pic_count = sheet.Pictures.Count
                
                for i in range(pic_count):
                    try:
                        pic = sheet.Pictures[i]
                        temp_filename = f"excel_img_{uuid.uuid4()}"
                        img_path = os.path.join(temp_dir, f"{temp_filename}.png")
                        
                        # 尝试直接保存图片
                        pic.Picture.Save(img_path)
                        
                        if os.path.exists(img_path) and os.path.getsize(img_path) > 0:
                            image_info = {
                                'path': img_path,
                                'sheet': sheet.Name,
                                'left': pic.LeftColumnOffset,
                                'top': pic.TopRowOffset,
                                'width': pic.Width,
                                'height': pic.Height
                            }
                            images.append(image_info)
                    except Exception as e:
                        print(f"使用Spire.XLS提取图片时出错: {str(e)}")
                        continue
        
        print(f"成功提取 {len(images)} 个图片")
        
        # 压缩图片
        compressed_images = []
        for img_info in images:
            input_path = img_info['path']
            filename = os.path.basename(input_path)
            output_path = os.path.join(compressed_dir, f"compressed_{filename}")
            
            if optimize_image(input_path, output_path):
                compressed_info = img_info.copy()
                compressed_info['compressed_path'] = output_path
                compressed_images.append(compressed_info)
                
                # 显示压缩前后的大小
                original_size_kb = os.path.getsize(input_path) / 1024
                compressed_size_kb = os.path.getsize(output_path) / 1024
                print(f"图片压缩: {original_size_kb:.1f}KB -> {compressed_size_kb:.1f}KB ({compressed_size_kb/original_size_kb:.1%})")
        
        # 复制原始文件到输出位置进行修改
        shutil.copy2(input_file, output_file)
        
        # 使用win32com替换图片
        pythoncom.CoInitialize()
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            
            workbook = excel.Workbooks.Open(os.path.abspath(output_file))
            
            # 替换图片
            for img_info in compressed_images:
                sheet_name = img_info['sheet']
                sheet = workbook.Sheets(sheet_name)
                
                # 删除原图
                for shape in sheet.Shapes:
                    if abs(shape.Left - img_info['left']) < 5 and abs(shape.Top - img_info['top']) < 5:
                        shape.Delete()
                        break
                
                # 插入压缩后的图片
                sheet.Shapes.AddPicture(
                    Filename=img_info['compressed_path'],
                    LinkToFile=False,
                    SaveWithDocument=True,
                    Left=img_info['left'],
                    Top=img_info['top'],
                    Width=img_info['width'],
                    Height=img_info['height']
                )
            
            # 保存文件
            workbook.Save()
            workbook.Close()
            excel.Quit()
            
        except Exception as e:
            print(f"替换图片时出错: {str(e)}")
        finally:
            pythoncom.CoUninitialize()
        
        # 检查最终文件大小
        final_size = os.path.getsize(output_file)
        compression_ratio = final_size / original_size
        print(f"压缩比率: {compression_ratio:.2%}")
        
        if compression_ratio > target_size_ratio:
            print("警告：未达到目标压缩比率")
            
            # 如果压缩比率不理想，可以尝试更激进的压缩设置
            if compression_ratio > 0.8 and len(compressed_images) > 0:
                print("正在使用更激进的压缩设置重试...")
                # 递归调用，使用更激进的设置
                return compress_excel_file(input_file, output_file, target_size_ratio * 0.8)
        
    except Exception as e:
        print(f"主程序出错: {str(e)}")
        raise
    finally:
        try:
            shutil.rmtree(temp_dir)
            print("临时文件已清理")
        except:
            print(f"临时文件夹未能删除: {temp_dir}")

def analyze_file_size(file_path):
    """分析文件大小"""
    size_bytes = os.path.getsize(file_path)
    size_mb = size_bytes / (1024 * 1024)
    return size_mb

if __name__ == "__main__":
    print("请确保已安装必要的库：")
    print("pip install Spire.XLS")
    print("pip install Pillow")
    print("pip install pywin32")
    
    input_file = '美团刷单报销.xlsx'
    output_file = 'compressed_美团刷单报销2.xlsx'

    print(f"原始文件大小: {analyze_file_size(input_file):.2f} MB")

    try:
        compress_excel_file(input_file, output_file, target_size_ratio=0.5)
        print(f"压缩后文件大小: {analyze_file_size(output_file):.2f} MB")
        print("处理完成！")
    except Exception as e:
        print(f"处理过程中出错: {str(e)}")