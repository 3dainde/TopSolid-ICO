from PIL import Image, ImageDraw, ImageFont
import shutil, os

sizes = [16, 24, 32, 48, 64, 128, 256]
imgs = []

for s in sizes:
    img = Image.new('RGBA', (s, s), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    # Gray background rounded rectangle
    r = max(1, int(s * 0.1))
    draw.rounded_rectangle([0, 0, s-1, s-1], radius=r, fill=(120, 120, 120, 255))

    # Three small dark squares arranged in an L-shape pattern
    pad = int(s * 0.2)
    sq = int(s * 0.25)
    gap = int(s * 0.06)

    # Top-left square
    x1, y1 = pad, pad
    draw.rectangle([x1, y1, x1+sq, y1+sq], fill=(50, 50, 50, 255))

    # Top-right square
    x2, y2 = x1 + sq + gap, pad
    draw.rectangle([x2, y2, x2+sq, y2+sq], fill=(50, 50, 50, 255))

    # Bottom-left square
    x3, y3 = pad, y1 + sq + gap
    draw.rectangle([x3, y3, x3+sq, y3+sq], fill=(50, 50, 50, 255))

    imgs.append(img)

ico_path = r'D:\TopSolid_Icones\TopICO\app.ico'
imgs[-1].save(ico_path, format='ICO', append_images=imgs[:-1])
print(f'Icon: {ico_path} ({os.path.getsize(ico_path)} bytes)')

dst = r'D:\TopSolid_Icones\BreakIcons\app.ico'
shutil.copy2(ico_path, dst)
print(f'Copied: {dst}')
