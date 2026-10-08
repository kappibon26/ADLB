"""Inserta dashboard/data.json y la plantilla Excel (dashboard/plantilla.xlsx) en la página y genera dashboard/index.html."""
import base64, json, pathlib
root = pathlib.Path(__file__).resolve().parent.parent / "dashboard"
data = json.loads((root / "data.json").read_text(encoding="utf-8"))
html = (root / "app.template.html").read_text(encoding="utf-8")
blob = json.dumps(data, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
tpl = base64.b64encode((root / "plantilla.xlsx").read_bytes()).decode()
(root / "index.html").write_text(html.replace("__DATA__", blob).replace("__TPL__", tpl), encoding="utf-8")
print("ok", len(data["rows"]), "filas")
