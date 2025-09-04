import streamlit as st
import os
import io
from cv_functions import *

# Configuration de la page
st.set_page_config(
    page_title="Générateur de CV Clinkast - IA Enhanced",
    page_icon="🤖",
    layout="wide"
)

def main():
    st.title("🤖 Générateur de CV Clinkast - Analyse IA des Missions")
    st.markdown("*Optimisation intelligente du dossier de compétences selon la mission*")
    st.markdown("---")
    
    st.subheader("📁 Documents d'entrée")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("**📋 Description de la mission**")
        mission_file = st.file_uploader(
            "Chargez la description de mission",
            type=['pdf', 'txt'],
            help="Fichier PDF ou TXT contenant la description détaillée de la mission"
        )
        
        mission_text = None
        if mission_file:
            # Validation du fichier mission
            if valider_fichier_upload(mission_file, ['pdf', 'txt'], 10):
                if mission_file.type == "application/pdf":
                    mission_text = lire_fichier_pdf(mission_file)
                else:
                    mission_text = lire_fichier_txt(mission_file)
                
                if mission_text:
                    # Nettoyer le texte de la mission
                    mission_text = nettoyer_texte_mission(mission_text)
                    st.success(f"✅ Mission chargée ({len(mission_text.split())} mots)")
                    with st.expander("Aperçu du contenu", expanded=False):
                        st.text(mission_text[:500] + "..." if len(mission_text) > 500 else mission_text)
    
    with col2:
        st.markdown("**👤 Dossier de compétences actuel**")
        cv_file = st.file_uploader(
            "Chargez le CV/Dossier de compétences",
            type=['pdf', 'docx'],
            help="Document PDF ou Word contenant le dossier de compétences à adapter"
        )
        
        cv_text = None
        if cv_file:
            # Validation du fichier CV
            if valider_fichier_upload(cv_file, ['pdf', 'docx'], 10):
                if cv_file.type == "application/pdf":
                    cv_text = lire_fichier_pdf(cv_file)
                else:
                    cv_text = lire_fichier_word(cv_file)
                
                if cv_text:
                    st.success(f"✅ Dossier chargé ({len(cv_text.split())} mots)")
                    with st.expander("Aperçu du contenu", expanded=False):
                        st.text(cv_text[:500] + "..." if len(cv_text) > 500 else cv_text)
    
    with col3:
        st.markdown("**📄 Template Word**")
        template_file = st.file_uploader(
            "Chargez votre template Word",
            type=['docx'],
            help="Template Word avec les placeholders Clinkast"
        )
        
        if template_file:
            st.success("✅ Template chargé")
            with st.expander("Placeholders supportés", expanded=False):
                st.markdown("""
                - `{{nom_consultant}}`
                - `{{titre_du_poste}}`
                - `{{points_forts}}`
                - `{{niveaux_intervention}}`
                - `{{tableau_formation}}`
                - `{{tableau_connaissances}}`
                - `{{tableau_hobbies}}`
                - `{{tableau_experiences}}`
                """)
    
    st.subheader("⚙️ Configuration")
    
    col_config1, col_config2 = st.columns([2, 1])
    
    with col_config1:
        nom_fichier = st.text_input(
            "Nom du fichier de sortie",
            value="CV_optimise_mission.docx"
        )
        
        # Vérification de la configuration OpenAI
        test_client = configurer_openai()
        if not test_client:
            st.error("⚠️ Impossible de configurer OpenAI. Vérifiez votre clé API.")
            st.info("💡 Ajoutez votre clé OpenAI dans les secrets de Streamlit pour utiliser l'analyse IA.")
            return
        else:
            st.success("✅ Configuration OpenAI détectée")
    
    with col_config2:
        st.info("""
        **Analyse automatique :**
        - Détection du domaine
        - Optimisation des compétences  
        - Reformulation intelligente
        - Priorisation des expériences
        """)
    
    # Bouton de génération
    if st.button("🚀 Analyser et Générer", type="primary", disabled=not all([mission_file, cv_file, template_file])):
        
        if not all([mission_file, cv_file, template_file]):
            st.warning("⚠️ Veuillez charger tous les documents requis")
            return
        
        # Étape 1: Lecture des documents (si pas déjà fait)
        if not mission_text or not cv_text:
            with st.spinner("📖 Lecture des documents..."):
                if mission_file.type == "application/pdf":
                    mission_content = lire_fichier_pdf(mission_file)
                else:
                    mission_content = lire_fichier_txt(mission_file)
                
                # Nettoyer le texte de mission
                mission_content = nettoyer_texte_mission(mission_content)
                
                if cv_file.type == "application/pdf":
                    cv_content = lire_fichier_pdf(cv_file)
                else:
                    cv_content = lire_fichier_word(cv_file)
        else:
            mission_content = mission_text
            cv_content = cv_text
        
        if not mission_content or not cv_content:
            st.error("❌ Erreur lors de la lecture des documents")
            return
        
        # Étape 2: Analyse du domaine et suggestions
        with st.spinner("🔍 Analyse du domaine d'activité..."):
            domaine_detecte = detecter_domaine_mission(mission_content)
            categories_suggerees = extraire_categories_connaissances_par_domaine(cv_content, domaine_detecte)
        
        st.info(f"📋 **Domaine détecté:** {domaine_detecte}")
        
        if categories_suggerees:
            with st.expander("🎯 Catégories de compétences suggérées", expanded=False):
                for categorie, mots_cles in categories_suggerees.items():
                    if mots_cles:
                        st.markdown(f"**{categorie}:** {', '.join(mots_cles[:5])}")
        
        # Étape 3: Analyse IA
        st.info("🤖 Analyse intelligente avec OpenAI en cours...")
        
        donnees_optimisees = appeler_openai_pour_optimisation(mission_content, cv_content)
        
        if not donnees_optimisees:
            st.error("❌ Erreur lors de l'analyse IA")
            return
        
        # Vérification du format des données
        if not isinstance(donnees_optimisees, dict):
            st.error("❌ Format de données invalide reçu de l'IA")
            return
        
        # Étape 4: Génération du rapport et affichage des résultats
        rapport = generer_rapport_optimisation(cv_content, donnees_optimisees, mission_content)
        
        st.success("✅ Analyse terminée avec succès !")
        
        # Métriques principales
        col_metric1, col_metric2, col_metric3, col_metric4 = st.columns(4)
        
        with col_metric1:
            score_pct = rapport['score_adequation'] * 100
            st.metric(
                "Score d'adéquation",
                f"{score_pct:.0f}%",
                delta="Excellent" if score_pct > 80 else "Bon" if score_pct > 60 else "Moyen"
            )
        
        with col_metric2:
            st.metric("Domaine détecté", rapport['domaine_detecte'])
        
        with col_metric3:
            st.metric("Expériences", rapport['nb_experiences'])
        
        with col_metric4:
            st.metric("Formations", rapport['nb_formations'])
        
        # Détails du rapport
        with st.expander("📊 Rapport détaillé d'optimisation", expanded=True):
            col_rapport1, col_rapport2 = st.columns(2)
            
            with col_rapport1:
                st.markdown("**🎯 Catégories identifiées dans le profil:**")
                for cat in rapport['categories_identifiees'][:5]:
                    st.markdown(f"• {cat}")
            
            with col_rapport2:
                st.markdown("**🔧 Catégories optimisées pour la mission:**")
                for cat in rapport['categories_competences'][:5]:
                    st.markdown(f"• {cat}")
        
        with st.expander("📊 Données complètes optimisées", expanded=False):
            st.json(donnees_optimisees)
        
        # Étape 5: Génération du CV
        with st.spinner("📝 Génération du CV optimisé..."):
            try:
                # Utiliser le template uploadé par l'utilisateur
                doc = generer_cv_depuis_template_avec_entete_preserve(template_file, donnees_optimisees)
                
                if doc:
                    # Sauvegarder en mémoire
                    doc_buffer = io.BytesIO()
                    doc.save(doc_buffer)
                    doc_buffer.seek(0)
                    
                    st.success(f"🎉 CV optimisé généré avec succès !")
                    
                    # Bouton de téléchargement
                    st.download_button(
                        label="📥 Télécharger le CV Optimisé",
                        data=doc_buffer.getvalue(),
                        file_name=nom_fichier,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        type="primary"
                    )
                    
                    # Sauvegarder l'historique
                    sauvegarder_historique_generation(donnees_optimisees, nom_fichier, rapport)
                    
                    # Statistiques de l'optimisation
                    st.subheader("📈 Résumé de l'optimisation")
                    col_stats1, col_stats2, col_stats3 = st.columns(3)
                    
                    with col_stats1:
                        st.metric(
                            "Expériences adaptées", 
                            len(donnees_optimisees.get('experiences', []))
                        )
                    
                    with col_stats2:
                        st.metric(
                            "Points forts", 
                            len(donnees_optimisees.get('points_forts', []))
                        )
                    
                    with col_stats3:
                        st.metric(
                            "Catégories de compétences", 
                            len(donnees_optimisees.get('connaissances', {}))
                        )
                else:
                    st.error("❌ Erreur lors de la génération du document")
                
            except Exception as e:
                st.error(f"❌ Erreur lors de la génération : {str(e)}")
                import traceback
                st.error(f"Détails: {traceback.format_exc()}")

# Guide d'utilisation
def afficher_guide():
    st.markdown("---")
    st.subheader("📚 Guide d'utilisation")
    
    st.markdown("""
    **Processus d'optimisation automatique :**
    
    1. **📋 Chargez la description de mission** (PDF/TXT)
       - Document contenant les détails de la mission/poste
       - L'IA analysera les compétences requises
    
    2. **👤 Chargez le dossier de compétences** actuel (PDF/Word)  
       - CV ou dossier existant à adapter
       - Sera optimisé selon les besoins de la mission
    
    3. **📄 Chargez votre template Word** (.docx)
       - Template avec les placeholders Clinkast
       - Évite les problèmes de corruption de fichier
    
    4. **🤖 L'IA analyse et optimise automatiquement :**
       - Détection du domaine d'activité
       - Adaptation des compétences pertinentes
       - Reformulation intelligente des expériences
       - Priorisation des points forts
       - Exclusion des éléments non pertinents
    
    5. **📥 Téléchargez** le CV optimisé au format Word
    
    **Avantages :**
    - ✅ Adaptation automatique à chaque mission
    - ✅ Optimisation des compétences pertinentes  
    - ✅ Reformulation professionnelle
    - ✅ Flexibilité du template utilisateur
    - ✅ Score d'adéquation calculé
    - ✅ Support des profils multi-domaines
    - ✅ Maximisation des réalisations pertinentes
    
    **Domaines spécialisés supportés :**
    - Développement & Programmation
    - DevOps & Infrastructure  
    - Cybersécurité
    - Intelligence Artificielle & Data
    - Business Intelligence & Analytics
    - Architecture & Systèmes
    - Marketing Digital
    - Finance
    - Ressources Humaines
    - Logistique & Supply Chain
    - Consulting & Stratégie
    - Et plus...
    """)

if __name__ == "__main__":
    main()
    afficher_guide()