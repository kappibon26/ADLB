"""Inserta dashboard/data.json en la plantilla y genera dashboard/index.html."""
import json, pathlib
root = pathlib.Path(__file__).resolve().parent.parent / "dashboard"
data = json.loads((root / "data.json").read_text(encoding="utf-8"))
html = (root / "app.template.html").read_text(encoding="utf-8")
blob = json.dumps(data, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
(root / "index.html").write_text(html.replace("__DATA__", blob), encoding="utf-8")
print("ok", len(data["rows"]), "filas")
