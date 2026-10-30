# -*- coding: utf-8 -*-
import gradio as gr
from fusion_core import fusionner
import os

def interface_fusion(fichiers):
    if not fichiers:
        return None, "Veuillez sélectionner au moins un fichier."
    
    # Préparation des fichiers au format attendu par fusion_core.py : list of (filename, bytes)
    fichiers_fusion = []
    for f in fichiers:
        # f est un objet temporaire Gradio contenant le chemin .name
        nom_origine = os.path.basename(f.name)
        with open(f.name, "rb") as file_bytes:
            contenu = file_bytes.read()
        fichiers_fusion.append((nom_origine, contenu))
        
    # Appel de votre fonction de fusion d'origine
    pdf_output, erreurs = fusionner(fichiers_fusion)
    
    if pdf_output is None:
        msg_erreur = "Erreur lors du traitement :\n" + "\n".join(erreurs)
        return gr.update(visible=False), gr.update(value=msg_erreur, visible=True)
    
    # Sauvegarde temporaire du PDF généré pour que Gradio puisse le proposer au téléchargement
    chemin_pdf = "Better_Mind_fusion.pdf"
    with open(chemin_pdf, "wb") as f_out:
        f_out.write(pdf_output)
        
    msg_status = "✅ Fusion réussie !"
    if erreurs:
        msg_status += "\n\n⚠️ Certaines erreurs mineures sont survenues :\n" + "\n".join(erreurs)
        
    return gr.update(value=chemin_pdf, visible=True), gr.update(value=msg_status, visible=True)

# Création d'une interface graphique moderne en quelques lignes
with gr.Blocks(title="Better Mind fusion") as demo:
    gr.Markdown("# 🧠 Better Mind fusion")
    gr.Markdown("Fusionnez vos fichiers Word (.docx), Excel (.xlsx), PDF ou images (.jpg, .png) en un seul document PDF.")
    
    with gr.Row():
        fichiers_input = gr.File(
            file_count="multiple", 
            label="Déposez ou sélectionnez vos fichiers ici",
            file_types=[".docx", ".xlsx", ".pdf", ".jpg", ".jpeg", ".png", ".webp"]
        )
    
    bouton_fusion = gr.Button("⚡ Fusionner en PDF", variant="primary")
    
    # Éléments de sortie cachés par défaut
    status_output = gr.Textbox(label="Statut", visible=False, lines=3)
    file_output = gr.File(label="Télécharger votre PDF fusionné", visible=False)
    
    # Action au clic
    bouton_fusion.click(
        fn=interface_fusion,
        inputs=fichiers_input,
        outputs=[file_output, status_output]
    )

if __name__ == "__main__":
    demo.launch()
