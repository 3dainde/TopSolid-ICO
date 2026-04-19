from PIL import Image, ImageDraw, ImageFont
import shutil, os

sizes = [16, 24, 32, 48, 64, 128, 256]
imgs = []

for s in sizes:
    img = Image.new('RGBA', (s, s), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    # Full red rounded-square background for visibility
    r = max(2, int(s * 0.15))
    draw.rounded_rectangle([0, 0, s-1, s-1], radius=r, fill=(220, 50, 20, 255))

    # White bold "T" letter centered
    font_size = int(s * 0.7)
    font = None
    for fname in ['segoeuib.ttf', 'C:/Windows/Fonts/segoeuib.ttf', 'C:/Windows/Fonts/arialbd.ttf']:
        try:
            font = ImageFont.truetype(fname, font_size)
            break
        except:
            continue
    if font is None:
        font = ImageFont.load_default()

    bbox = draw.textbbox((0, 0), 'T', font=font)
    tw = bbox[2] - bbox[0]
    th = bbox[3] - bbox[1]
    x = (s - tw) // 2
    y = (s - th) // 2 - int(s * 0.04)
    draw.text((x, y), 'T', fill=(255, 255, 255, 255), font=font)

    imgs.append(img)

ico_path = r'D:\TopSolid_Icones\TopICO\app.ico'
# Save all sizes properly
icon_sizes = [(s, s) for s in sizes]
imgs[-1].save(ico_path, format='ICO', append_images=imgs[:-1])
print(f'Icon: {ico_path} ({os.path.getsize(ico_path)} bytes)')

dst = r'D:\TopSolid_Icones\BreakIcons\app.ico'
shutil.copy2(ico_path, dst)
print(f'Copied: {dst}')
