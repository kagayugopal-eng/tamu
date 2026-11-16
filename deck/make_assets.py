"""Generate background images and crop the logo from the reference screenshot."""
import sys
from PIL import Image, ImageDraw, ImageFilter

ref = sys.argv[1]
W, H = 1920, 1080


def triangles(base):
    ov = Image.new("RGBA", (W, H), (255, 255, 255, 0))
    d = ImageDraw.Draw(ov)
    d.polygon([(0, 380), (300, 900), (0, 1080)], fill=(120, 200, 235, 40))
    d.polygon([(0, 620), (560, 1080), (0, 1080)], fill=(214, 64, 159, 28))
    d.polygon([(1100, 0), (1920, 0), (1920, 300)], fill=(150, 120, 245, 24))
    d.polygon([(1500, 1080), (1920, 700), (1920, 1080)], fill=(120, 200, 235, 30))
    return Image.alpha_composite(base, ov)


def gradient_line(img, y, th):
    stops = [(0, (34, 178, 216)), (0.55, (123, 92, 240)), (1, (214, 64, 159))]
    d = ImageDraw.Draw(img)
    for x in range(W):
        t = x / (W - 1)
        for (a, ca), (b, cb) in zip(stops, stops[1:]):
            if a <= t <= b:
                k = (t - a) / (b - a)
                c = tuple(int(ca[i] + (cb[i] - ca[i]) * k) for i in range(3))
                break
        d.line([(x, y), (x, y + th)], fill=c + (255,))


base = Image.new("RGBA", (W, H), (255, 255, 255, 255))
cover = triangles(base)
cover.convert("RGB").save("cover_bg.png")
content = triangles(base)
gradient_line(content, int(0.88 / 5.625 * H), 6)
content.convert("RGB").save("content_bg.png")

im = Image.open(ref).convert("RGB")
print("ref size", im.size)
# logo is at top-right of the reference screenshot
sx = im.size[0] / 1341
box = tuple(int(v * sx) for v in (1150, 12, 1300, 72))
logo = im.crop(box).convert("RGBA")
px = logo.load()
for y in range(logo.size[1]):
    for x in range(logo.size[0]):
        r, g, b, a = px[x, y]
        m = min(r, g, b)
        if m > 235:
            px[x, y] = (r, g, b, 0)
        elif m > 200:
            px[x, y] = (r, g, b, int((235 - m) / 35 * 255))
logo.save("gopal_logo.png")
print("logo", logo.size)
