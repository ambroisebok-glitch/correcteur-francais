"""Noyau du correcteur de français : lecture des documents, détection des fautes
(LanguageTool), application des corrections. Aucune dépendance à Flask ni à Streamlit."""
import bisect
import io
import json
import os
import re
import shutil
import subprocess
import tempfile
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
from xml.sax.saxutils import escape

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.text.run import Run
from pypdf import PdfReader

LT_URL = os.environ.get("LT_URL", "https://api.languagetool.org/v2/check")
PUBLIC_LT = "languagetool.org" in LT_URL
CHUNK = int(os.environ.get("LT_CHUNK", "6000"))      # caractères envoyés par requête
LT_USERNAME = os.environ.get("LT_USERNAME", "")      # option : abonnement LanguageTool
LT_APIKEY = os.environ.get("LT_APIKEY", "")


# --------------------------------------------------------------------------- lecture des documents

def para_runs(par):
    """Tous les segments de texte d'un paragraphe, y compris ceux des liens hypertextes."""
    return [Run(r, par) for r in par._p.iter(qn("w:r"))]


def para_text(par):
    return "".join(r.text for r in para_runs(par))


def iter_block(parent_elm, parent):
    """Paragraphes dans l'ordre du document, tableaux compris."""
    for child in parent_elm.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, parent)
        elif child.tag == qn("w:tbl"):
            table = Table(child, parent)
            for row in table.rows:
                seen = []
                for cell in row.cells:
                    if any(cell._tc is s for s in seen):
                        continue
                    seen.append(cell._tc)
                    yield from iter_block(cell._tc, cell)


def pdf_paragraphs(path):
    reader = PdfReader(path)
    if reader.is_encrypted:
        try:
            reader.decrypt("")
        except Exception:
            raise ValueError("Ce PDF est protégé par un mot de passe.")
    paras = []
    for page in reader.pages:
        text = page.extract_text() or ""
        lines = [l.rstrip() for l in text.split("\n")]
        longest = max([len(l) for l in lines] + [1])
        buf = ""
        for line in lines:
            if not line.strip():
                if buf:
                    paras.append(buf)
                    buf = ""
                continue
            if buf.endswith("-") and line[:1].islower():
                buf = buf[:-1] + line.strip()          # mot coupé en fin de ligne
            elif buf and re.search(r"[.!?…»]$", buf) and len(buf) < 0.75 * longest:
                paras.append(buf)
                buf = line.strip()
            else:
                buf = (buf + " " + line.strip()).strip()
        if buf:
            paras.append(buf)
    if not "".join(paras).strip():
        raise ValueError("Aucun texte trouvé dans ce PDF (document scanné ?).")
    return paras


def load_paragraphs(path, ext):
    """Renvoie (document_docx_ou_None, paragraphes_docx, textes)."""
    if ext == ".docx":
        doc = Document(path)
        pars = list(iter_block(doc.element.body, doc))
        return doc, pars, [para_text(p) for p in pars]
    if ext == ".pdf":
        return None, None, pdf_paragraphs(path)
    with open(path, "rb") as f:
        raw = f.read()
    try:
        text = raw.decode("utf-8-sig")
    except UnicodeDecodeError:
        text = raw.decode("cp1252", errors="replace")
    return None, None, [p.strip() for p in re.split(r"\n\s*\n", text.replace("\r\n", "\n")) if p.strip()]


# --------------------------------------------------------------------------- correction (LanguageTool)

_LT_LOCK = threading.Lock()
_LT_LAST = [0.0, 0]    # moment et taille du dernier envoi


def lt_check(text):
    """Un seul envoi à la fois vers le service public, en respectant ses limites de débit
    (20 requêtes et 75 000 caractères par minute), même quand plusieurs visiteurs corrigent en même temps."""
    if not PUBLIC_LT:
        return _lt_post(text)
    with _LT_LOCK:
        gap = max(3.2, _LT_LAST[1] / 75000 * 60)
        wait = gap - (time.time() - _LT_LAST[0])
        if wait > 0:
            time.sleep(wait)
        try:
            return _lt_post(text)
        finally:
            _LT_LAST[0], _LT_LAST[1] = time.time(), len(text)


def _lt_post(text):
    params = {"text": text, "language": "fr"}
    if LT_USERNAME and LT_APIKEY:
        params.update(username=LT_USERNAME, apiKey=LT_APIKEY)
    data = urllib.parse.urlencode(params).encode()
    req = urllib.request.Request(LT_URL, data=data, headers={"Accept": "application/json"})
    for attempt in range(2):
        try:
            with urllib.request.urlopen(req, timeout=90) as r:
                return json.load(r)["matches"]
        except urllib.error.HTTPError as e:
            if e.code == 429 and attempt == 0:
                time.sleep(25)
                continue
            if e.code == 429:
                raise RuntimeError("Limite du service gratuit atteinte. Réessaie dans une minute "
                                   "ou utilise un serveur LanguageTool local (voir LISEZ-MOI).")
            raise RuntimeError("Le service de correction a répondu avec une erreur (%d)." % e.code)
        except urllib.error.URLError:
            raise RuntimeError("Impossible de joindre le service de correction. Vérifie ta connexion internet "
                               "ou l'adresse LT_URL.")


def check_texts(texts, progress):
    chunks, cur, size = [], [], 0
    for i, t in enumerate(texts):
        if not t.strip():
            continue
        if cur and size + len(t) + 2 > CHUNK:
            chunks.append(cur)
            cur, size = [], 0
        cur.append(i)
        size += len(t) + 2
    if cur:
        chunks.append(cur)

    found = []
    for ci, idxs in enumerate(chunks):
        text, starts = "", []
        for i in idxs:
            starts.append(len(text))
            text += texts[i] + "\n\n"
        for m in lt_check(text):
            o, l = m["offset"], m["length"]
            k = bisect.bisect_right(starts, o) - 1
            if k < 0:
                continue
            pid, local = idxs[k], o - starts[k]
            if local + l > len(texts[pid]):
                continue
            found.append({
                "pid": pid, "start": local, "end": local + l,
                "original": texts[pid][local:local + l],
                "message": m.get("message", ""),
                "rule": m.get("rule", {}).get("id", ""),
                "category": m.get("rule", {}).get("category", {}).get("name", ""),
                "suggestions": [r["value"] for r in m.get("replacements", [])][:5],
            })
        progress((ci + 1) / len(chunks))

    found.sort(key=lambda x: (x["pid"], x["start"]))
    issues, prev_pid, prev_end = [], -1, -1
    for it in found:                      # on écarte les chevauchements
        if it["pid"] == prev_pid and it["start"] < prev_end:
            continue
        prev_pid, prev_end = it["pid"], it["end"]
        it["id"] = len(issues)
        issues.append(it)
    return issues


# --------------------------------------------------------------------------- application des corrections

def apply_edits_to_paragraph(par, edits):
    """edits : liste de (début, fin, remplacement) en positions du texte d'origine.
    Les corrections sont écrites dans les segments existants pour garder la mise en forme."""
    spans, pos = [], 0
    for r in para_runs(par):
        t = r.text
        spans.append((r, pos, pos + len(t)))
        pos += len(t)
    for s, e, repl in sorted(edits, reverse=True):
        placed = False
        for r, a, b in spans:
            if b <= a:
                continue
            if r._r.xpath(".//w:drawing|.//w:pict"):
                continue
            if e > s and (b <= s or a >= e):
                continue
            if e == s and not (a <= s <= b):
                continue
            t = r.text
            ls, le = max(s, a) - a, min(e, b) - a
            new = t[:ls] + ("" if placed else repl) + t[le:]
            placed = True
            if new != t:
                r.text = new
            if e == s:
                break


def selected_edits(texts, issues, decisions):
    by_par = {}
    for it in issues:
        repl = decisions.get(str(it["id"]))
        if repl is None:
            continue
        t = texts[it["pid"]]
        if t[it["start"]:it["end"]] != it["original"]:
            continue
        by_par.setdefault(it["pid"], []).append((it["start"], it["end"], repl))
    return by_par


def clean_for_pdf(s):
    s = s.replace(" ", " ").replace(" ", " ").replace("\t", "    ")
    return escape(s).replace("\n", "<br/>")


def texts_to_pdf(texts, out, title):
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.platypus import Paragraph as RLParagraph, SimpleDocTemplate
    style = ParagraphStyle("corps", fontName="Times-Roman", fontSize=11.5, leading=16, spaceAfter=8)
    doc = SimpleDocTemplate(out, pagesize=A4, title=title, leftMargin=2.2 * cm, rightMargin=2.2 * cm,
                            topMargin=2.2 * cm, bottomMargin=2.2 * cm)
    doc.build([RLParagraph(clean_for_pdf(t), style) for t in texts if t.strip()])


def texts_to_docx(texts, out):
    d = Document()
    for t in texts:
        if t.strip():
            d.add_paragraph(t)
    d.save(out)


def soffice_to_pdf(docx_path, outdir):
    exe = shutil.which("soffice") or shutil.which("libreoffice")
    if not exe:
        return None
    try:
        subprocess.run([exe, "--headless", "--convert-to", "pdf", "--outdir", outdir, docx_path],
                       check=True, capture_output=True, timeout=180)
    except Exception:
        return None
    pdf = os.path.splitext(docx_path)[0] + ".pdf"
    return pdf if os.path.exists(pdf) else None


# --------------------------------------------------------------------------- fonctions pour l'interface

def analyze_bytes(data, ext, progress=lambda p: None):
    """Renvoie (textes_des_paragraphes, signalements)."""
    with tempfile.TemporaryDirectory() as d:
        path = os.path.join(d, "source" + ext)
        with open(path, "wb") as f:
            f.write(data)
        _, _, texts = load_paragraphs(path, ext)
    if not "".join(texts).strip():
        raise ValueError("Le document ne contient aucun texte.")
    return texts, check_texts(texts, progress)


def _apply(texts, issues, decisions):
    by_par = selected_edits(texts, issues, decisions)
    corrected = list(texts)
    for pid, edits in by_par.items():
        t = texts[pid]
        for s, e, r in sorted(edits, reverse=True):
            t = t[:s] + r + t[e:]
        corrected[pid] = t
    return by_par, corrected


def build_docx(data, ext, texts, issues, decisions):
    """decisions : {str(id) : texte de remplacement}. Renvoie le .docx corrigé (octets)."""
    by_par, corrected = _apply(texts, issues, decisions)
    out = io.BytesIO()
    if ext == ".docx":
        with tempfile.TemporaryDirectory() as d:
            path = os.path.join(d, "source.docx")
            with open(path, "wb") as f:
                f.write(data)
            doc, pars, _ = load_paragraphs(path, ".docx")
            for pid, edits in by_par.items():
                apply_edits_to_paragraph(pars[pid], edits)
            doc.save(out)
    else:
        texts_to_docx(corrected, out)
    return out.getvalue()


def build_pdf(data, ext, texts, issues, decisions, title="document"):
    """PDF corrigé (octets). Pour un .docx, LibreOffice est utilisé s'il est installé (mise en page conservée)."""
    _, corrected = _apply(texts, issues, decisions)
    if ext == ".docx" and (shutil.which("soffice") or shutil.which("libreoffice")):
        with tempfile.TemporaryDirectory() as d:
            p = os.path.join(d, "corrige.docx")
            with open(p, "wb") as f:
                f.write(build_docx(data, ext, texts, issues, decisions))
            pdf = soffice_to_pdf(p, d)
            if pdf:
                with open(pdf, "rb") as f:
                    return f.read()
    out = io.BytesIO()
    texts_to_pdf(corrected, out, title)
    return out.getvalue()
