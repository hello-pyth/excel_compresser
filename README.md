# Excel 图片压缩工具

该项目旨在从 Excel 文件中提取图片并进行压缩，以减小文件大小。使用了 `win32com` 和 `Spire.XLS` 库来处理 Excel 文件，使用 `Pillow` 库来优化图片。

## 功能

- 从 Excel 文件中提取图片
- 优化图片大小，确保不超过指定的大小
- 支持 JPEG 和 PNG 格式
- 生成压缩后的 Excel 文件

## 依赖

请确保已安装以下库：

```bash
pip install Spire.XLS
pip install Pillow
pip install pywin32
```

## 使用方法

1. 将待处理的 Excel 文件放在项目目录中。
2. 修改 `Excel_image_compresser.py` 文件中的 `input_file` 和 `output_file` 变量，指定输入和输出文件名。
3. 运行脚本：

```bash
python Excel_image_compresser.py
```

4. 脚本将输出原始文件大小和压缩后的文件大小。

## 函数说明

- `get_image_data_direct(pic)`: 直接通过二进制方式获取图片数据。
- `extract_images_with_win32com(input_file, output_dir)`: 使用 `win32com` 提取图片。
- `optimize_image(input_path, output_path, max_size_kb=300)`: 优化图片大小，确保不超过指定大小。
- `compress_excel_file(input_file, output_file, target_size_ratio=0.5)`: 使用多种方法压缩 Excel 文件中的图片。
- `analyze_file_size(file_path)`: 分析文件大小。

## 注意事项

- 请确保 Excel 文件中包含图片。
- 该工具在 Windows 环境下运行良好。

