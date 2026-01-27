from __future__ import annotations

from pathlib import Path

from potxkit import PotxTemplate

PALETTE = {
    "dark1": "#FFFFFF",
    "light1": "#0B0B0E",
    "dark2": "#2C2C34",
    "light2": "#E9ECF2",
    "accent1": "#1F6BFF",
    "accent2": "#E0328C",
    "accent3": "#F6A225",
    "accent4": "#6B3AF6",
    "accent5": "#38D3FF",
    "accent6": "#FF4D6D",
    "hlink": "#1F6BFF",
    "folHlink": "#C0186B",
}


def main() -> None:
    tpl = PotxTemplate.new()
    tpl.theme.colors.set_dark1(PALETTE["dark1"])
    tpl.theme.colors.set_light1(PALETTE["light1"])
    tpl.theme.colors.set_dark2(PALETTE["dark2"])
    tpl.theme.colors.set_light2(PALETTE["light2"])
    tpl.theme.colors.set_accent(1, PALETTE["accent1"])
    tpl.theme.colors.set_accent(2, PALETTE["accent2"])
    tpl.theme.colors.set_accent(3, PALETTE["accent3"])
    tpl.theme.colors.set_accent(4, PALETTE["accent4"])
    tpl.theme.colors.set_accent(5, PALETTE["accent5"])
    tpl.theme.colors.set_accent(6, PALETTE["accent6"])
    tpl.theme.colors.set_hyperlink(PALETTE["hlink"])
    tpl.theme.colors.set_followed_hyperlink(PALETTE["folHlink"])

    tpl.theme.fonts.set_major("Aptos Display")
    tpl.theme.fonts.set_minor("Aptos")

    out_path = Path("templates/default-dark.potx")
    out_path.parent.mkdir(parents=True, exist_ok=True)
    tpl.save(str(out_path))


if __name__ == "__main__":
    main()
