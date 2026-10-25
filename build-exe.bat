@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ============================================
echo   正在生成应用程序图标...
echo ============================================
python -c "
import os
from PIL import Image, ImageDraw, ImageFont

sizes = [(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)]
images = []
font_path = os.path.join(os.environ.get('WINDIR', r'C:\Windows'), 'Fonts', 'segmdl2.ttf')

for w, h in sizes:
    img = Image.new('RGBA', (w, h), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    radius = int(w * 0.22)
    draw.rounded_rectangle([0, 0, w-1, h-1], radius=radius, fill='#3B82F6')
    try:
        font_size = int(w * 0.55)
        font = ImageFont.truetype(font_path, font_size)
        draw.text((w / 2, h / 2), '\uE8B9', font=font, fill='#FFFFFF', anchor='mm')
    except Exception:
        pass
    images.append(img)

images[0].save('app.ico', format='ICO', sizes=[(img.width, img.height) for img in images], append_images=images[1:])
"

echo.
echo ============================================
echo   正在打包为单文件 exe (直接输出到项目根目录)...
echo ============================================
pyinstaller --noconfirm --onefile --windowed --name "公众号素材整理" --icon "app.ico" --collect-all docx --collect-all pptx --hidden-import lxml --hidden-import bs4 --distpath "." "weixin-gui.py"

if exist "dist" rmdir /s /q "dist"

echo.
echo 打包完成！已生成：公众号素材整理.exe
pause
