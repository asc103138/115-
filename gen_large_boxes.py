import os
from PIL import Image, ImageDraw, ImageFont

FONT_PATH = "C:/Windows/Fonts/kaiu.ttf"

# High-resolution large handwriting box
# Box: W_BOX = 130, H = 130
# Zhuyin column: W_ZY = 54
# Total width: 184 px, height: 130 px
W_BOX = 130
W_ZY = 54
H = 130

f_zy = ImageFont.truetype(FONT_PATH, 36)
f_tone = ImageFont.truetype(FONT_PATH, 30)
f_hz = ImageFont.truetype(FONT_PATH, 82)

os.makedirs("exam_boxes_large", exist_ok=True)

def generate_large_box_img(char, zhuyin_str, mode, out_path):
    img = Image.new("RGB", (W_BOX + W_ZY, H), (255, 255, 255))
    d = ImageDraw.Draw(img)
    # Outer border and vertical divider (3px solid black)
    d.rectangle([2, 2, W_BOX, H - 3], outline=(0, 0, 0), width=3)
    d.rectangle([W_BOX, 2, W_BOX + W_ZY, H - 3], outline=(0, 0, 0), width=3)
    
    tones = ['ˊ', 'ˇ', 'ˋ', '˙']
    tone = ''
    body = []
    for c in zhuyin_str:
        if c in tones:
            tone = c
        else:
            body.append(c)
            
    if mode == "char_blank":
        # Blank square on left for student to write character.
        # Zhuyin symbols centered in the right column.
        n = len(body)
        step = (H - 24) / max(n, 1)
        for i, c in enumerate(body):
            y = 10 + i * step
            d.text((W_BOX + 6, y), c, fill=(0, 0, 0), font=f_zy)
        if tone:
            if tone == '˙':
                # Light tone goes above the first symbol
                d.text((W_BOX + 14, 2), tone, fill=(0, 0, 0), font=f_tone)
            else:
                # 2nd, 3rd, 4th tones go to the right of the middle/last symbol
                tone_y = 10 + (n - 0.7) * step - 14
                d.text((W_BOX + 32, tone_y), tone, fill=(0, 0, 0), font=f_tone)
    elif mode == "zhuyin_blank":
        # Character printed in square on left.
        # Zhuyin blank on right for student to write zhuyin!
        bbox = d.textbbox((0, 0), char, font=f_hz)
        cw, ch = bbox[2] - bbox[0], bbox[3] - bbox[1]
        x = (W_BOX - cw) / 2
        y = (H - ch) / 2 - 10
        d.text((x, y), char, fill=(0, 0, 0), font=f_hz)
        
    img.save(out_path)
    return out_path

items_p1 = [
    ("彷", "ㄈㄤˇ", "char_blank"),
    ("鷹", "ㄧㄥ", "zhuyin_blank"),
    ("衝", "ㄔㄨㄥ", "char_blank"),
    ("竟", "ㄐㄧㄥˋ", "zhuyin_blank"),
    ("訝", "ㄧㄚˋ", "char_blank"),
    ("恐", "ㄎㄨㄥˇ", "zhuyin_blank"),
    ("樹", "ㄕㄨˋ", "char_blank"),
    ("蕉", "ㄐㄧㄠ", "zhuyin_blank"),
    ("狀", "ㄓㄨㄤˋ", "char_blank"),
    ("隊", "ㄉㄨㄟˋ", "zhuyin_blank"),
    ("寧", "ㄋㄧㄥˊ", "char_blank"),
    ("控", "ㄎㄨㄥˋ", "zhuyin_blank"),
    ("制", "ㄓˋ", "char_blank"),
    ("碑", "ㄅㄟ", "zhuyin_blank"),
    ("巡", "ㄒㄩㄣˊ", "char_blank"),
    ("甜", "ㄊㄧㄢˊ", "zhuyin_blank"),
    ("航", "ㄏㄤˊ", "zhuyin_blank"),
    ("糖", "ㄊㄤˊ", "char_blank"),
    ("榮", "ㄖㄨㄥˊ", "zhuyin_blank"),
    ("衛", "ㄨㄟˋ", "char_blank"),
    ("灣", "ㄨㄢ", "zhuyin_blank"),
    ("嶼", "ㄩˇ", "char_blank"),
    ("翔", "ㄒㄧㄤˊ", "char_blank"),
    ("研", "ㄧㄢˊ", "zhuyin_blank"),
    ("造", "ㄗㄠˋ", "char_blank"),
    ("篇", "ㄆㄧㄢ", "zhuyin_blank"),
    ("盤", "ㄆㄢˊ", "char_blank"),
    ("耀", "ㄧㄠˋ", "zhuyin_blank"),
]

box_paths = {}
for i, (ch, zy, mode) in enumerate(items_p1):
    path = f"exam_boxes_large/box_{i}_{ch}_{mode}.png"
    generate_large_box_img(ch, zy, mode, path)
    box_paths[(ch, mode)] = path

print(f"Generated {len(box_paths)} large boxes.")
