"""
Generates loxone_icon.ico — Loxone green rounded square + bold white checkmark.
Run once:  py make_icon.py
"""
from PIL import Image, ImageDraw, ImageFont
import os


LOXONE_GREEN = (91, 163, 25)    # #5BA319
DARK_GREEN   = (58, 110, 10)    # darker outline / shadow
WHITE        = (255, 255, 255)


def draw_icon(size: int) -> Image.Image:
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d   = ImageDraw.Draw(img)

    pad = max(1, size // 20)
    r   = size // 5          # corner radius

    # ── Rounded square background ──────────────────────────────
    d.rounded_rectangle(
        [pad, pad, size - pad - 1, size - pad - 1],
        radius=r,
        fill=LOXONE_GREEN,
    )

    # ── Subtle inner shadow (darker bottom-right rim) ──────────
    d.rounded_rectangle(
        [pad, pad, size - pad - 1, size - pad - 1],
        radius=r,
        outline=DARK_GREEN,
        width=max(1, size // 24),
    )

    # ── Bold white checkmark ────────────────────────────────────
    # The tick is drawn as a thick polyline with two segments:
    #   A  (left)  → B (pivot, lower-centre)  → C  (right, top)
    s  = size
    ax = s * 0.18;  ay = s * 0.52
    bx = s * 0.41;  by = s * 0.72
    cx = s * 0.80;  cy = s * 0.28

    lw = max(2, size // 7)   # line width — bold

    # Draw the tick twice (once slightly offset) for a clean look
    d.line([(ax, ay), (bx, by), (cx, cy)],
           fill=WHITE, width=lw, joint="curve")

    return img


def main():
    sizes  = [16, 24, 32, 48, 64, 128, 256]
    frames = [draw_icon(s) for s in sizes]

    out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "loxone_icon.ico")
    frames[0].save(
        out,
        format="ICO",
        sizes=[(s, s) for s in sizes],
        append_images=frames[1:],
    )
    print(f"Icon saved: {out}")

    # also save a 256-px PNG preview
    prev = os.path.join(os.path.dirname(out), "loxone_icon_preview.png")
    frames[-1].convert("RGBA").save(prev)
    print(f"Preview:    {prev}")


if __name__ == "__main__":
    main()
