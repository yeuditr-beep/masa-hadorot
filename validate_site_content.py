import json
import pathlib

ROOT = pathlib.Path(__file__).parent
CONTENT = ROOT / "content"

files = sorted(CONTENT.rglob("*.json"))
data_by_file = {}
errors = []

for path in files:
    try:
        data_by_file[path] = json.loads(path.read_text(encoding="utf-8"))
    except Exception as exc:
        errors.append(f"JSON {path}: {exc}")

figures_path = CONTENT / "figures.json"
figure_ids = {f.get("id") for f in data_by_file.get(figures_path, {}).get("figures", [])}

chapters = 0
slides = 0
questions = 0
for path, data in data_by_file.items():
    if "chapter" not in path.name:
        continue
    chapters += 1
    slides += len(data.get("slides", []))
    questions += len(data.get("quiz", []))
    for slide in data.get("slides", []):
        image = slide.get("image")
        if image and not (ROOT / image).exists():
            errors.append(f"Missing image {path.name}: {image}")
        for figure_id in slide.get("figures", []):
            if figure_id not in figure_ids:
                errors.append(f"Missing figure {path.name}: {figure_id}")
    for idx, q in enumerate(data.get("quiz", []), 1):
        options = q.get("options", [])
        answer = q.get("answer")
        if not isinstance(answer, int) or answer < 0 or answer >= len(options):
            errors.append(f"Invalid answer {path.name} Q{idx}: {answer}")

print({
    "json_files": len(files),
    "chapters": chapters,
    "slides": slides,
    "questions": questions,
    "figures": len(figure_ids),
    "errors": len(errors),
})
for error in errors:
    print(error)
raise SystemExit(1 if errors else 0)
