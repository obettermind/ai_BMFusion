# -*- coding: utf-8 -*-
"""fusion_core.py - moteur de fusion pour Better Mind fusion.

Convertit des fichiers Word (.docx), Excel (.xlsx) et images en PDF,
puis les fusionne dans l'ordre reçu.

Conversion Office -> PDF, dans cet ordre :
  1. Microsoft Word / Excel via COM (Windows, si Office est installé)
  2. LibreOffice en ligne de commande (Windows, Linux, mode hébergé)

Dépendances : pip install pypdf pillow pywin32
(pywin32 est facultatif : seulement pour utiliser Word/Excel sous Windows)
"""
import io
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from PIL import Image, ImageOps
from pypdf import PdfReader, PdfWriter

IMG_EXT = {".jpg", ".jpeg", ".png", ".webp"}
WORD_EXT = {".docx"}
EXCEL_EXT = {".xlsx"}
NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)


# ---------------------------------------------------------------- images
def _image_vers_pdf(contenu: bytes) -> bytes:
    img = Image.open(io.BytesIO(contenu))
    img = ImageOps.exif_transpose(img)
    if img.mode in ("RGBA", "LA", "P"):
        fond = Image.new("RGB", img.size, "white")
        rgba = img.convert("RGBA")
        fond.paste(rgba, mask=rgba.split()[-1])
        img = fond
    else:
        img = img.convert("RGB")
    sortie = io.BytesIO()
    img.save(sortie, format="PDF", resolution=150.0)
    return sortie.getvalue()


# ------------------------------------------------------- Office via COM
def _office_com(src: str, dst: str, ext: str) -> None:
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    app = None
    try:
        if ext in WORD_EXT:
            app = win32com.client.DispatchEx("Word.Application")
            app.Visible = False
            app.DisplayAlerts = 0
            doc = app.Documents.Open(src, ReadOnly=True)
            try:
                doc.ExportAsFixedFormat(dst, 17)  # 17 = wdExportFormatPDF
            finally:
                doc.Close(False)
        else:
            app = win32com.client.DispatchEx("Excel.Application")
            app.Visible = False
            app.DisplayAlerts = False
            classeur = app.Workbooks.Open(src, ReadOnly=True)
            try:
                classeur.ExportAsFixedFormat(0, dst)  # 0 = xlTypePDF
            finally:
                classeur.Close(False)
    finally:
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


# ----------------------------------------------------------- LibreOffice
def _trouver_soffice():
    chemin = shutil.which("soffice") or shutil.which("libreoffice")
    if chemin:
        return chemin
    for c in (r"C:\Program Files\LibreOffice\program\soffice.exe",
              r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"):
        if os.path.exists(c):
            return c
    return None


def _libreoffice(src: str, dst: str) -> None:
    exe = _trouver_soffice()
    if not exe:
        raise RuntimeError("LibreOffice est introuvable")
    dossier = os.path.dirname(dst)
    profil = tempfile.mkdtemp(prefix="lo_profil_")
    try:
        subprocess.run(
            [exe, "-env:UserInstallation=" + Path(profil).as_uri(),
             "--headless", "--convert-to", "pdf", "--outdir", dossier, src],
            check=True, timeout=180, creationflags=NO_WINDOW,
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        )
    finally:
        shutil.rmtree(profil, ignore_errors=True)
    produit = os.path.join(dossier, Path(src).stem + ".pdf")
    if not os.path.exists(produit):
        raise RuntimeError("LibreOffice n'a pas produit de PDF")
    if os.path.abspath(produit) != os.path.abspath(dst):
        os.replace(produit, dst)


def _office_vers_pdf(contenu: bytes, ext: str) -> bytes:
    with tempfile.TemporaryDirectory(prefix="bm_fusion_") as tmp:
        src = os.path.abspath(os.path.join(tmp, "source" + ext))
        dst = os.path.abspath(os.path.join(tmp, "source.pdf"))
        with open(src, "wb") as f:
            f.write(contenu)

        erreurs = []
        essais = []
        if sys.platform == "win32":
            essais.append(("Microsoft Office", lambda: _office_com(src, dst, ext)))
        essais.append(("LibreOffice", lambda: _libreoffice(src, dst)))

        for nom, action in essais:
            try:
                action()
                if os.path.exists(dst):
                    with open(dst, "rb") as f:
                        return f.read()
            except Exception as e:
                erreurs.append(f"{nom} : {e}")
        raise RuntimeError("; ".join(erreurs) or "conversion impossible")


# --------------------------------------------------------------- fusion
def fusionner(fichiers):
    """fichiers : liste de (nom_fichier, contenu_bytes).

    Retourne (pdf_bytes ou None, liste_des_erreurs).
    """
    writer = PdfWriter()
    erreurs = []
    ajoutes = 0

    for nom, contenu in fichiers:
        ext = Path(nom).suffix.lower()
        try:
            if ext in IMG_EXT:
                pdf = _image_vers_pdf(contenu)
            elif ext in WORD_EXT or ext in EXCEL_EXT:
                pdf = _office_vers_pdf(contenu, ext)
            elif ext == ".pdf":
                pdf = contenu
            else:
                erreurs.append(f"{nom} : format non pris en charge")
                continue

            lecteur = PdfReader(io.BytesIO(pdf))
            for page in lecteur.pages:
                writer.add_page(page)
            ajoutes += 1
        except Exception as e:
            erreurs.append(f"{nom} : {e}")

    if ajoutes == 0:
        return None, erreurs or ["Aucun fichier traité."]

    sortie = io.BytesIO()
    writer.write(sortie)
    return sortie.getvalue(), erreurs
