# Генерирует иконку приложения build/icon.ico (+ icon.png 512 для просмотра) средствами Pillow.
# Собственный рисунок: перо (кончик пера с прорезью) на скруглённом квадрате фирменного цвета.
# Запуск: python build/make_icon.py
from pathlib import Path
from PIL import Image, ImageDraw

S = 1024                      # рисуем крупно, потом уменьшаем — так края сглаживаются
GREEN = (31, 111, 92, 255)    # #1f6f5c
WHITE = (255, 255, 255, 255)
OUT = Path(__file__).parent


def draw() -> Image.Image:
    img = Image.new('RGBA', (S, S), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle((32, 32, S - 32, S - 32), radius=200, fill=GREEN)

    # Кончик пера рисуем вертикально (остриём вниз) на отдельном слое и поворачиваем на 45°
    nib = Image.new('RGBA', (S, S), (0, 0, 0, 0))
    n = ImageDraw.Draw(nib)
    cx = S // 2
    top, shoulder, tip = 170, 520, 900
    half = 175
    n.polygon([(cx - half, top), (cx + half, top), (cx + half, shoulder), (cx, tip), (cx - half, shoulder)], fill=WHITE)
    n.ellipse((cx - half, top - half // 2, cx + half, top + half // 2), fill=WHITE)  # скругление верха
    hole_y, r = 520, 52
    n.ellipse((cx - r, hole_y - r, cx + r, hole_y + r), fill=GREEN)               # отверстие
    n.rectangle((cx - 13, hole_y, cx + 13, tip - 40), fill=GREEN)                 # прорезь до острия
    nib = nib.rotate(-45, resample=Image.BICUBIC, center=(cx, S // 2 + 30))
    img.alpha_composite(nib, (0, -20))
    return img


if __name__ == '__main__':
    big = draw()
    big.resize((512, 512), Image.LANCZOS).save(OUT / 'icon.png')
    sizes = [(16, 16), (20, 20), (24, 24), (32, 32), (40, 40), (48, 48), (64, 64), (128, 128), (256, 256)]
    big.resize((256, 256), Image.LANCZOS).save(OUT / 'icon.ico', sizes=sizes)
    print('ok:', OUT / 'icon.ico')
