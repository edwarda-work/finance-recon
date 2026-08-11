from pathlib import Path
import re
import tomllib


config = tomllib.loads(Path("/Users/fido-edward/.codex/config.toml").read_text())
color_like = re.compile(r"(?:#[0-9a-fA-F]{3,8}|rgb\(|hsl\(|^[0-9A-Fa-f]{6,8}$)")


def walk(value, path=""):
    if isinstance(value, dict):
        for key, child in value.items():
            current = f"{path}.{key}" if path else key
            walk(child, current)
        return
    print("KEY", path)
    leaf = path.rsplit(".", 1)[-1].lower()
    if any(token in leaf for token in ("color", "colour", "theme", "palette", "accent", "background", "foreground")):
        print("PALETTE_VALUE", path, repr(value))
    elif isinstance(value, str) and color_like.search(value):
        print("PALETTE_VALUE", path, repr(value))


walk(config)
