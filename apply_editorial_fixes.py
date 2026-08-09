import json
import pathlib

from build_editorial_report import CONTENT, corrected


def rewrite(value):
    changes = 0
    if isinstance(value, dict):
        result = {}
        for key, item in value.items():
            new_item, count = rewrite(item)
            result[key] = new_item
            changes += count
        return result, changes
    if isinstance(value, list):
        result = []
        for item in value:
            new_item, count = rewrite(item)
            result.append(new_item)
            changes += count
        return result, changes
    if isinstance(value, str):
        new_value = corrected(value)
        return new_value, int(new_value != value)
    return value, 0


total = 0
for path in sorted(CONTENT.rglob("*.json")):
    data = json.loads(path.read_text(encoding="utf-8"))
    revised, count = rewrite(data)
    if count:
        path.write_text(
            json.dumps(revised, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        print(f"{path.relative_to(CONTENT)}: {count}")
        total += count

print(f"TOTAL={total}")
