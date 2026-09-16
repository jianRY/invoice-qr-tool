import os
from PIL import Image

png = r"C:\Users\toxuj\WorkBuddy\发票处理\generated-images\A_clean_modern_app_icon_design_2026-08-14T08-54-45.png"
ico = r"C:\Users\toxuj\WorkBuddy\发票处理\app_icon.ico"

img = Image.open(png).convert("RGBA")
# 裁剪为居中 960x960，去掉右下角 AI 水印
w, h = img.size
left = (w - 960) // 2
top = (h - 960) // 2
right = left + 960
bottom = top + 960
img = img.crop((left, top, right, bottom))
# 缩放并生成多尺寸 ICO
sizes = [(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)]
img.save(ico, format="ICO", sizes=sizes)
print(f"Saved ICO: {ico} from {img.size}")
