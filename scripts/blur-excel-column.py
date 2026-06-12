#!/usr/bin/env python3
"""
Blur a specific column in an Excel Camera screenshot.

Usage:
    python blur-excel-column.py <xlsx_path> <screenshot_path> <column_letter> [output_path]

Example:
    python blur-excel-column.py report.xlsx report.png C report_blurred.png
    # Blurs column C (用户名) in the screenshot

Dependencies: pywin32, Pillow, openpyxl

注意：本脚本仅 Windows + Microsoft Excel 环境可用（依赖 win32com）。
如不传 capture_range 参数，默认截图范围为 A1:N15，超出此范围需手动指定。
"""
import sys, os


def get_column_ratio(xlsx_path: str, capture_range: str, column_letter: str) -> tuple:
    """
    Get the x-position ratio of a column within the capture range.
    Returns (x_start_ratio, x_end_ratio).
    """
    # 延迟导入 win32com，让无 Excel 环境的用户收到友好提示而非崩溃
    try:
        import win32com.client as win32
    except ImportError:
        print('错误：需要 pywin32 库。请运行: pip install pywin32')
        sys.exit(1)
    except AttributeError:
        print('错误：win32com 初始化失败。请确认已安装 Microsoft Excel。')
        sys.exit(1)

    excel = None
    wb = None
    try:
        excel = win32.gencache.EnsureDispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(xlsx_path)
        ws = wb.ActiveSheet
        cap = ws.Range(capture_range)
        col = ws.Columns(column_letter)
        x_start = (col.Left - cap.Left) / cap.Width
        x_end = (col.Left + col.Width - cap.Left) / cap.Width
        return (x_start, x_end)
    finally:
        if wb:
            wb.Close(SaveChanges=False)
        if excel:
            excel.Quit()


def blur_column(screenshot_path: str, output_path: str,
                x_start_ratio: float, x_end_ratio: float,
                radius: float = 12):
    """Apply GaussianBlur to a vertical column region in the screenshot."""
    try:
        from PIL import Image, ImageFilter
    except ImportError:
        print('错误：需要 Pillow 库。请运行: pip install Pillow')
        sys.exit(1)

    img = Image.open(screenshot_path)
    w, h = img.size
    x1 = int(w * x_start_ratio)
    x2 = int(w * x_end_ratio)
    region = img.crop((x1, 0, x2, h))
    blurred = region.filter(ImageFilter.GaussianBlur(radius=radius))
    result = img.copy()
    result.paste(blurred, (x1, 0))
    result.save(output_path)
    return (w, h, x1, x2)


def main():
    if len(sys.argv) < 4:
        print(__doc__)
        sys.exit(1)

    xlsx_path = sys.argv[1]
    screenshot_path = sys.argv[2]
    column_letter = sys.argv[3].upper()
    output_path = sys.argv[4] if len(sys.argv) > 4 else (
        os.path.splitext(screenshot_path)[0] + "_blurred.png"
    )
    capture_range = sys.argv[5] if len(sys.argv) > 5 else "A1:N15"

    print(f"Calculating column {column_letter} position in {capture_range}...")
    x_start, x_end = get_column_ratio(xlsx_path, capture_range, column_letter)
    print(f"Column {column_letter}: x_ratio={x_start:.4f} to {x_end:.4f}")

    print(f"Blurring {screenshot_path} -> {output_path}...")
    w, h, px1, px2 = blur_column(screenshot_path, output_path, x_start, x_end)
    print(f"Image: {w}x{h}, blurred x=[{px1}, {px2}] ({px2-px1}px wide)")
    print("Done.")


if __name__ == "__main__":
    main()
