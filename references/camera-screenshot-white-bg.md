# Excel Camera 截图：白色背景处理

## 问题

使用 `win32com` 的 `CopyPicture` → 粘贴到临时工作表 → `ImageGrab.grabclipboard()` 截取的 PNG 图片背景是透明的（透明像素对应 Excel 表无填充色单元格、行间空区域），导致表中文字在白底浏览时看不清。

## 根因

`CopyPicture` 以图片格式复制时保留填充色（如表头蓝底），但白色/无色单元格的透明通道会如实传递。`PIL.ImageGrab.grabclipboard()` 拿到 RGBA 模式的图像，透明(alpha=0) 区域在白色背景上等同于白色，但在纯 PNG 展示时显示为透明。

## 修复

保存前将 RGBA 粘贴到纯白 RGB 背景上：

```python
from PIL import ImageGrab, Image

img = ImageGrab.grabclipboard()
if img.mode in ('RGBA', 'LA') or (img.mode == 'P' and 'transparency' in img.info):
    background = Image.new('RGB', img.size, (255, 255, 255))
    if img.mode == 'P':
        img = img.convert('RGBA')
    background.paste(img, mask=img.split()[-1] if img.mode == 'RGBA' else None)
    img = background
elif img.mode != 'RGB':
    img = img.convert('RGB')
img.save(output_path, 'PNG')
```

## 完整脚本参考

每次文件打开/关闭间最好重建 Excel 实例（`win32com.client.Dispatch("Excel.Application")`），防止 Paste 操作在第 N 个文件后失败。

`D:\\hermes_files\\temp\\camera_screenshots_v3.py` 是 12 个文件截图的生产脚本模板。

## 批量截图可靠性坑（2026-06-22）

CopyPicture → Paste → Clipboard 链路在连续调用（12+次）时会随机失败，报错「Microsoft Excel 无法粘贴数据」。稳定模式：每个文件用独立 Python 进程 + 独立 Excel.Application，每次执行后 `taskkill /f /im EXCEL.EXE` 再跑下一个。模板脚本：`D:\\hermes_files\\temp\\camera_one.py` + `D:\\hermes_files\\temp\\batch_camera.py`。
