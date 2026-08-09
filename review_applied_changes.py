import json
import pathlib

root = pathlib.Path(__file__).parent
old_root = root / "content-backup-before-editorial-20260806"
new_root = root / "content"


def strings(obj, path=""):
    if isinstance(obj, dict):
        for k, v in obj.items():
            yield from strings(v, f"{path}/{k}")
    elif isinstance(obj, list):
        for i, v in enumerate(obj):
            yield from strings(v, f"{path}/{i}")
    elif isinstance(obj, str):
        yield path, obj


for new_path in sorted(new_root.rglob("*.json")):
    old_path = old_root / new_path.relative_to(new_root)
    if not old_path.exists():
        continue
    old = dict(strings(json.loads(old_path.read_text(encoding="utf-8"))))
    new = dict(strings(json.loads(new_path.read_text(encoding="utf-8"))))
    for key in sorted(set(old) & set(new)):
        if old[key] != new[key]:
            print(f"{new_path.relative_to(new_root)} {key}\nOLD: {old[key]}\nNEW: {new[key]}\n")
