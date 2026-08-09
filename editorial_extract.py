import json
import pathlib
import sys

root = pathlib.Path("masa-hadorot/content")
pattern = sys.argv[1] if len(sys.argv) > 1 else "part-b/chapter-*.json"

for path in sorted(root.glob(pattern)):
    data = json.loads(path.read_text(encoding="utf-8"))
    print(f"\n### {path.as_posix()} | {data.get('title', '')}")
    print(f"SUBTITLE | {data.get('subtitle', '')}")
    for index, slide in enumerate(data.get("slides", []), 1):
        print(f"S{index} | {slide.get('title', '')} | {slide.get('narration', '')}")
    for index, question in enumerate(data.get("quiz", []), 1):
        options = " / ".join(question.get("options", []))
        print(f"Q{index} | {question.get('q', '')} | {options} | {question.get('explain', '')}")
