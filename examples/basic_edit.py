from __future__ import annotations

import sys

from potxkit import PotxTemplate


def main() -> int:
    if len(sys.argv) != 3:
        print("Usage: basic_edit.py input.potx output.potx")
        return 1

    input_path, output_path = sys.argv[1:3]
    tpl = PotxTemplate.open(input_path)

    tpl.theme.colors.set_accent(1, "#1F6BFF")
    tpl.theme.colors.set_accent(2, "#E0328C")
    tpl.theme.colors.set_hyperlink("#1F6BFF")
    tpl.theme.fonts.set_major("Aptos Display")
    tpl.theme.fonts.set_minor("Aptos")

    tpl.save(output_path)
    print(f"Wrote {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
