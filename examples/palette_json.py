from __future__ import annotations

import json
import sys

from potxkit import PotxTemplate


def main() -> int:
    if len(sys.argv) != 3:
        print("Usage: palette_json.py palette.json output.potx")
        return 1

    palette_path, output_path = sys.argv[1:3]
    palette = json.loads(open(palette_path, "r", encoding="utf-8").read())

    tpl = PotxTemplate.new()
    tpl.theme.colors.set_dark1(palette["dark1"])
    tpl.theme.colors.set_light1(palette["light1"])
    tpl.theme.colors.set_dark2(palette["dark2"])
    tpl.theme.colors.set_light2(palette["light2"])
    tpl.theme.colors.set_accent(1, palette["accent1"])
    tpl.theme.colors.set_accent(2, palette["accent2"])
    tpl.theme.colors.set_accent(3, palette["accent3"])
    tpl.theme.colors.set_accent(4, palette["accent4"])
    tpl.theme.colors.set_accent(5, palette["accent5"])
    tpl.theme.colors.set_accent(6, palette["accent6"])
    tpl.theme.colors.set_hyperlink(palette["hlink"])
    tpl.theme.colors.set_followed_hyperlink(palette["folHlink"])
    tpl.theme.fonts.set_major(palette["majorFont"])
    tpl.theme.fonts.set_minor(palette["minorFont"])

    tpl.save(output_path)
    print(f"Wrote {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
