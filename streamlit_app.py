"""Correcteur de français + Atelier de français (version Streamlit).

    pip install -r requirements.txt
    streamlit run streamlit_app.py
"""
import hashlib
import os

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(page_title="Correcteur de français", layout="wide")

# Les « secrets » de Streamlit (LT_URL, LT_USERNAME, LT_APIKEY) servent de réglages pour le moteur de correction.
try:
    for _cle in ("LT_URL", "LT_USERNAME", "LT_APIKEY"):
        if _cle in st.secrets and _cle not in os.environ:
            os.environ[_cle] = str(st.secrets[_cle])
except Exception:
    pass

import core  # noqa: E402  (après les réglages ci-dessus)

ICI = os.path.dirname(os.path.abspath(__file__))
MIME_DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"


def vis(s):
    return s.replace(" ", "␣").replace(" ", "␣").replace(" ", "␣") if s.strip() == "" else s


def contexte(texts, it):
    t = texts[it["pid"]]
    a, b = max(0, it["start"] - 35), min(len(t), it["end"] + 35)
    extrait = t[a:it["start"]] + "[" + t[it["start"]:it["end"]] + "]" + t[it["end"]:b]
    return (("… " if a > 0 else "") + extrait + (" …" if b < len(t) else "")).replace("\n", " ")


def page_correction():
    st.title("Correcteur de français")
    st.write("Dépose un document Word, PDF ou texte. Les fautes d'orthographe, de grammaire, d'accord et de "
             "conjugaison sont repérées. Tu choisis les corrections à appliquer, puis tu télécharges le résultat.")
    up = st.file_uploader("Document Word (.docx), PDF ou texte (.txt)", type=["docx", "pdf", "txt"])
    if up is None:
        st.caption("Les fichiers .doc (ancien format Word) ne sont pas pris en charge : enregistre-les en .docx.")
        return

    data = up.getvalue()
    ext = os.path.splitext(up.name)[1].lower()
    base = os.path.splitext(up.name)[0] or "document"
    fid = hashlib.sha256(data).hexdigest()

    if st.session_state.get("fid") != fid:
        barre = st.progress(0.0, text="Analyse du document…")
        try:
            texts, issues = core.analyze_bytes(
                data, ext, lambda p: barre.progress(min(1.0, max(0.0, p)), text="Analyse du document…"))
        except Exception as e:
            barre.empty()
            st.error(str(e) or "Erreur inattendue pendant l'analyse.")
            return
        barre.empty()
        st.session_state["fid"] = fid
        st.session_state["texts"] = texts
        st.session_state["issues"] = issues
        st.session_state["pdf"] = None
    texts, issues = st.session_state["texts"], st.session_state["issues"]

    if ext == ".pdf":
        st.warning("Le PDF est converti en texte : la mise en page d'origine (colonnes, images, tableaux) "
                   "n'est pas conservée dans le document corrigé.")
    if not issues:
        st.success("Aucune faute détectée.")
        return

    st.subheader("Corrections proposées")
    st.caption("Coche « Appliquer » pour garder une correction. Tu peux aussi modifier le texte de la colonne "
               "« Correction ». La faute est entre crochets dans la colonne « Contexte ».")
    mode = st.radio("Au départ", ["Appliquer toutes les propositions", "Ne rien appliquer"], horizontal=True)
    accepter = mode.startswith("Appliquer")

    lignes = []
    for it in issues:
        sug = it["suggestions"]
        lignes.append({
            "Appliquer": bool(sug) and accepter,
            "Faute": vis(it["original"]),
            "Correction": sug[0] if sug else "",
            "Explication": it["message"],
            "Contexte": contexte(texts, it),
            "Autres propositions": ", ".join(sug[1:]),
        })
    edite = st.data_editor(
        pd.DataFrame(lignes), hide_index=True, key="ed-%s-%s" % (fid, accepter),
        disabled=["Faute", "Explication", "Contexte", "Autres propositions"],
        column_config={
            "Appliquer": st.column_config.CheckboxColumn("Appliquer"),
            "Correction": st.column_config.TextColumn("Correction"),
        },
    )

    decisions = {}
    for it, (_, ligne) in zip(issues, edite.iterrows()):
        if bool(ligne["Appliquer"]):
            corr = ligne["Correction"]
            decisions[str(it["id"])] = "" if pd.isna(corr) else str(corr)
    n = len(decisions)
    st.write("%d correction%s appliquée%s sur %d signalement%s."
             % (n, "s" if n > 1 else "", "s" if n > 1 else "", len(issues), "s" if len(issues) > 1 else ""))

    c1, c2 = st.columns(2)
    try:
        docx_bytes = core.build_docx(data, ext, texts, issues, decisions)
        c1.download_button("Télécharger en Word", docx_bytes, file_name=base + "-corrigé.docx", mime=MIME_DOCX)
    except Exception as e:
        c1.error("Impossible de créer le fichier Word : %s" % e)

    signature = (fid, tuple(sorted(decisions.items())))
    if c2.button("Préparer le PDF"):
        try:
            st.session_state["pdf"] = (signature, core.build_pdf(data, ext, texts, issues, decisions, base))
        except Exception as e:
            c2.error("Impossible de créer le PDF : %s" % e)
    pdf = st.session_state.get("pdf")
    if pdf and pdf[0] == signature:
        c2.download_button("Télécharger en PDF", pdf[1], file_name=base + "-corrigé.pdf", mime="application/pdf")
    elif pdf:
        c2.info("Les corrections ont changé : prépare le PDF à nouveau.")


def page_atelier():
    chemin = os.path.join(ICI, "atelier-francais.html")
    with open(chemin, encoding="utf-8") as f:
        html = f.read()
    st.download_button("Télécharger l'atelier (fonctionne sans connexion, à ouvrir sur le téléphone)",
                       html, file_name="atelier-francais.html", mime="text/html")
    components.html(html, height=1100, scrolling=True)


onglet_corr, onglet_atelier = st.tabs(["Corriger un document", "Atelier de français"])
with onglet_corr:
    page_correction()
    st.divider()
    st.caption("Correction orthographique et grammaticale par [LanguageTool](https://languagetool.org). "
               "Le texte du document est envoyé à ce service pour être corrigé. "
               "Le document n'est pas conservé par cette application.")
with onglet_atelier:
    page_atelier()
