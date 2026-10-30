# -*- coding: utf-8 -*-
"""Better Mind fusion - version exe Windows (Flask local) et Web (Render).

Mode local (exe) : ouvre automatiquement le navigateur.
Mode hébergé : variable BM_HOSTED=1 puis  gunicorn better_mind_fusion:app (ou app:app selon le nom du fichier)
Le travail de fusion est fait par fusion_core.py (à garder dans le même dossier).
"""
import io
import os
import socket
import sys
import threading
import webbrowser
from xml.sax.saxutils import escape

# --- CORRECTIF POUR PYINSTALLER ---
if getattr(sys, 'frozen', False):
    base_dir = sys._MEIPASS
    if base_dir not in sys.path:
        sys.path.insert(0, base_dir)
else:
    base_dir = os.path.dirname(os.path.abspath(__file__))
    if base_dir not in sys.path:
        sys.path.insert(0, base_dir)
# ----------------------------------

if sys.stdout is None:
    sys.stdout = open(os.devnull, "w")
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w")

from flask import Flask, render_template_string, request, send_file
from fusion_core import fusionner

LOCAL = os.environ.get("BM_HOSTED") != "1"

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 300 * 1024 * 1024  # 300 Mo

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Better Mind fusion</title>
<style>
  body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background:#f4f7f6; margin:0; padding:24px 14px; display:flex; justify-content:center; }
  .card { background:#fff; padding:28px; border-radius:14px; box-shadow:0 4px 15px rgba(0,0,0,.06); max-width:520px; width:100%; }
  h1 { color:#2c3e50; font-size:24px; margin:0 0 6px; text-align:center; }
  .sub { color:#7f8c8d; text-align:center; margin:0 0 20px; font-size:14px; }
  .drop { display:block; padding:22px 12px; border:2px dashed #3498db; border-radius:10px; background:#fdfefe; text-align:center; cursor:pointer; color:#2c3e50; font-weight:600; }
  .drop small { display:block; font-weight:400; color:#7f8c8d; margin-top:6px; }
  ul { list-style:none; padding:0; margin:14px 0; }
  li { display:flex; align-items:center; gap:6px; padding:8px 10px; border:1px solid #e5e8e8; border-radius:8px; margin-bottom:6px; font-size:14px; }
  li .n { color:#95a5a6; width:22px; }
  li .name { flex:1; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
  li button { width:auto; padding:4px 9px; font-size:14px; background:#ecf0f1; color:#2c3e50; border-radius:6px; }
  button { background:#2ecc71; color:#fff; border:none; padding:13px 24px; font-size:16px; border-radius:8px; cursor:pointer; width:100%; font-weight:700; }
  button:hover { background:#27ae60; }
  button:disabled { background:#95a5a6; cursor:wait; }
  .msg { padding:14px; border-radius:8px; margin-top:16px; font-size:14px; }
  .err { background:#fdecea; color:#c0392b; border:1px solid #c0392b; }
  .ok { background:#e8f8f5; color:#1e8449; border:1px solid #27ae60; }
  .quit { background:none; color:#7f8c8d; font-weight:400; font-size:13px; margin-top:18px; padding:6px; text-decoration:underline; }
  .quit:hover { background:none; color:#c0392b; }
  .info { color:#95a5a6; font-size:12px; margin-top:14px; text-align:center; }
</style>
</head>
<body>
<div class="card">
  <h1>🧠 Better Mind fusion</h1>
  <p class="sub">Fusionnez vos fichiers Word, Excel et images en un seul PDF.</p>

  <form id="f" method="POST" action="/fusionner" enctype="multipart/form-data">
    <label class="drop" for="files">📂 Cliquez ici pour choisir vos fichiers
      <small>Word (.docx) · Excel (.xlsx) · Images (.jpg, .png)</small></label>
    <input id="files" type="file" name="files" accept=".docx,.xlsx,.jpg,.jpeg,.png,.webp" multiple hidden>
    <ul id="list"></ul>
    <button id="go" type="submit">⚡ Fusionner en PDF</button>
  </form>

  {% if error %}<div class="msg err">{{ error|safe }}</div>{% endif %}
  <div id="status" class="msg ok" style="display:none"></div>

  <div class="info">Vos fichiers sont traités sans être conservés. Le PDF se télécharge dans votre dossier « Téléchargements ».</div>
  {% if local %}
  <form method="POST" action="/quitter" style="text-align:center">
    <button class="quit" type="submit">Fermer Better Mind fusion</button>
  </form>
  {% endif %}
</div>

<script>
  const input = document.getElementById('files');
  const list = document.getElementById('list');
  const form = document.getElementById('f');
  const go = document.getElementById('go');
  const status = document.getElementById('status');
  let files = [];

  function sync() {
    const dt = new DataTransfer();
    files.forEach(f => dt.items.add(f));
    input.files = dt.files;
  }
  function render() {
    list.innerHTML = '';
    files.forEach((f, i) => {
      const li = document.createElement('li');
      li.innerHTML = '<span class="n">' + (i + 1) + '.</span><span class="name"></span>';
      li.querySelector('.name').textContent = f.name;
      [['▲', -1], ['▼', 1]].forEach(([label, d]) => {
        const b = document.createElement('button'); b.type = 'button'; b.textContent = label;
        b.onclick = () => { const j = i + d; if (j < 0 || j >= files.length) return;
          [files[i], files[j]] = [files[j], files[i]]; sync(); render(); };
        li.appendChild(b);
      });
      const x = document.createElement('button'); x.type = 'button'; x.textContent = '✕';
      x.onclick = () => { files.splice(i, 1); sync(); render(); };
      li.appendChild(x);
      list.appendChild(li);
    });
  }
  input.addEventListener('change', () => {
    files = files.concat(Array.from(input.files));
    sync(); render();
  });
  form.addEventListener('submit', (e) => {
    if (files.length === 0) { e.preventDefault(); alert('Choisissez au moins un fichier.'); return; }
    go.disabled = true; go.textContent = '⏳ Fusion en cours…';
    status.style.display = 'block';
    status.textContent = 'Fusion en cours, le PDF va se télécharger dans quelques secondes.';
    setTimeout(() => { go.disabled = false; go.textContent = '⚡ Fusionner en PDF';
      status.textContent = '✅ Terminé : vérifiez vos téléchargements.'; }, 4000);
  });
</script>
</body>
</html>
"""

def _page(**kw):
    return render_template_string(HTML_TEMPLATE, local=LOCAL, **kw)

@app.route("/")
def index():
    return _page()

@app.route("/fusionner", methods=["POST"])
def fusionner_route():
    uploaded = [f for f in request.files.getlist("files") if f and f.filename]
    if not uploaded:
        return _page(error="Aucun fichier sélectionné.")

    pdf, errors = fusionner([(f.filename, f.read()) for f in uploaded])
    if pdf is None:
        return _page(error="Aucun fichier n'a pu être traité :<br>"
                           + "<br>".join(escape(e) for e in errors))

    return send_file(io.BytesIO(pdf), mimetype="application/pdf", as_attachment=True,
                     download_name="Better_Mind_fusion.pdf")

@app.route("/quitter", methods=["POST"])
def quitter():
    if not LOCAL:
        return "", 404
    threading.Timer(0.8, lambda: os._exit(0)).start()
    return ("<html><body style='font-family:Segoe UI,sans-serif;text-align:center;padding:60px'>"
            "<h2>Better Mind fusion est fermé.</h2><p>Vous pouvez fermer cet onglet.</p></body></html>")

def _free_port(preferred=5000):
    for port in (preferred, 0):
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        try:
            s.bind(("127.0.0.1", port))
            return s.getsockname()[1]
        except OSError:
            continue
        finally:
            s.close()
    return preferred

if __name__ == "__main__":
    if LOCAL:
        port = _free_port(5000)
        url = f"http://127.0.0.1:{port}"
        threading.Timer(1.2, lambda: webbrowser.open(url)).start()
        app.run(host="127.0.0.1", port=port, debug=False)
    else:
        # Mode Cloud (Render) : écoute sur le port assigné dynamiquement
        port = int(os.environ.get("PORT", 5000))
        app.run(host="0.0.0.0", port=port)
