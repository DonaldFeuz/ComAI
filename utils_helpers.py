# utils.py - Fonctions utilitaires pour l'application CV Generator

import streamlit as st
import re
from typing import Dict, List, Any
import logging

# Configuration des logs
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def valider_fichier_upload(fichier, types_autorises: List[str], taille_max_mb: int = 10) -> bool:
    """
    Valide un fichier uploadé selon les critères spécifiés
    
    Args:
        fichier: Fichier uploadé via Streamlit
        types_autorises: Liste des extensions autorisées (ex: ['pdf', 'docx'])
        taille_max_mb: Taille maximale en MB
    
    Returns:
        bool: True si valide, False sinon
    """
    if not fichier:
        return False
    
    # Vérifier l'extension
    extension = fichier.name.split('.')[-1].lower()
    if extension not in types_autorises:
        st.error(f"Type de fichier non autorisé. Extensions acceptées: {', '.join(types_autorises)}")
        return False
    
    # Vérifier la taille
    taille_mb = len(fichier.getvalue()) / (1024 * 1024)
    if taille_mb > taille_max_mb:
        st.error(f"Fichier trop volumineux ({taille_mb:.1f} MB). Taille max: {taille_max_mb} MB")
        return False
    
    return True

def nettoyer_texte_mission(texte: str) -> str:
    """
    Nettoie et formate le texte de la mission pour l'analyse IA
    
    Args:
        texte: Texte brut extrait du document
    
    Returns:
        str: Texte nettoyé et formaté
    """
    if not texte:
        return ""
    
    # Supprimer les caractères de contrôle et espaces excessifs
    texte = re.sub(r'\s+', ' ', texte)
    texte = re.sub(r'[^\w\s\-\.,:;!?\(\)\/]', '', texte)
    
    # Nettoyer les lignes vides multiples
    texte = re.sub(r'\n\s*\n', '\n\n', texte)
    
    return texte.strip()

def detecter_domaine_mission(texte_mission: str) -> str:
    """
    Détecte le domaine d'activité principal de la mission
    
    Args:
        texte_mission: Texte de la description de mission
    
    Returns:
        str: Domaine détecté
    """
    
    domaines_keywords = {
        'IT/Tech': [
            'développeur', 'programmeur', 'software', 'application', 'système', 'réseau',
            'base de données', 'cloud', 'devops', 'api', 'framework', 'coding', 'javascript',
            'python', 'java', 'docker', 'kubernetes', 'git', 'agile', 'scrum'
        ],
        'Marketing': [
            'marketing', 'communication', 'campagne', 'publicité', 'brand', 'social media',
            'seo', 'sem', 'analytics', 'crm', 'lead', 'conversion', 'content', 'digital'
        ],
        'Finance': [
            'finance', 'comptabilité', 'budget', 'trésorerie', 'audit', 'contrôle gestion',
            'reporting financier', 'analyse financière', 'investissement', 'risque'
        ],
        'RH': [
            'ressources humaines', 'recrutement', 'formation', 'paie', 'talent',
            'compétences', 'évaluation', 'carrière', 'mobilité', 'sirh'
        ],
        'Logistique': [
            'logistique', 'supply chain', 'approvisionnement', 'stock', 'transport',
            'entreposage', 'distribution', 'procurement', 'planification'
        ],
        'Vente': [
            'commercial', 'vente', 'client', 'négociation', 'business development',
            'account manager', 'prospection', 'closing', 'pipeline'
        ],
        'Consulting': [
            'consultant', 'conseil', 'stratégie', 'transformation', 'audit',
            'accompagnement', 'optimisation', 'expertise'
        ],
        'Santé': [
            'médical', 'santé', 'patient', 'soins', 'clinique', 'hôpital',
            'pharmacie', 'thérapie', 'diagnostic'
        ],
        'Education': [
            'formation', 'enseignement', 'pédagogie', 'cours', 'étudiant',
            'apprentissage', 'curriculum', 'évaluation pédagogique'
        ],
        'Juridique': [
            'juridique', 'droit', 'contrat', 'compliance', 'réglementation',
            'contentieux', 'avocat', 'juriste'
        ]
    }
    
    texte_lower = texte_mission.lower()
    scores = {}
    
    for domaine, keywords in domaines_keywords.items():
        score = sum(1 for keyword in keywords if keyword in texte_lower)
        if score > 0:
            scores[domaine] = score
    
    if scores:
        return max(scores.items(), key=lambda x: x[1])[0]
    
    return 'Général'

def extraire_categories_connaissances_par_domaine(texte_cv: str, domaine: str) -> Dict[str, List[str]]:
    """
    Extrait et suggère des catégories de connaissances selon le domaine
    
    Args:
        texte_cv: Contenu du CV
        domaine: Domaine détecté de la mission
    
    Returns:
        Dict avec les catégories suggérées et mots-clés trouvés
    """
    
    categories_par_domaine = {
        'IT/Tech': {
            'Langages de programmation': ['python', 'java', 'javascript', 'c#', 'php', 'ruby', 'go', 'rust'],
            'Frameworks et Libraries': ['react', 'angular', 'vue', 'django', 'flask', 'spring', 'laravel'],
            'Bases de données': ['mysql', 'postgresql', 'mongodb', 'oracle', 'redis', 'elasticsearch'],
            'Cloud et DevOps': ['aws', 'azure', 'gcp', 'docker', 'kubernetes', 'terraform', 'jenkins'],
            'Outils et IDE': ['git', 'github', 'gitlab', 'vscode', 'intellij', 'eclipse', 'jira'],
            'Méthodologies': ['agile', 'scrum', 'kanban', 'devops', 'ci/cd', 'tdd', 'bdd']
        },
        'Marketing': {
            'Outils Marketing': ['hubspot', 'marketo', 'mailchimp', 'salesforce', 'pardot'],
            'Analytics et Mesure': ['google analytics', 'tag manager', 'data studio', 'tableau'],
            'Réseaux Sociaux': ['facebook ads', 'google ads', 'linkedin', 'instagram', 'twitter'],
            'CRM et Automation': ['salesforce', 'hubspot', 'pipedrive', 'zoho', 'crm'],
            'Design et Création': ['photoshop', 'illustrator', 'canva', 'figma', 'indesign'],
            'SEO et Content': ['seo', 'sem', 'content marketing', 'wordpress', 'cms']
        },
        'Finance': {
            'Logiciels Financiers': ['sap', 'oracle', 'sage', 'cegid', 'quickbooks'],
            'Outils d\'Analyse': ['excel', 'power bi', 'tableau', 'qlikview', 'r', 'python'],
            'Réglementation': ['ifrs', 'pcg', 'sox', 'bâle', 'solvabilité'],
            'Reporting': ['consolidation', 'reporting', 'business intelligence', 'kpi'],
            'Contrôle et Audit': ['contrôle interne', 'audit', 'risk management', 'compliance'],
            'Treasury': ['trésorerie', 'cash management', 'forex', 'dérivés', 'financement']
        },
        'RH': {
            'SIRH': ['sap hr', 'workday', 'adp', 'cornerstone', 'talentsoft', 'sirh'],
            'Recrutement': ['ats', 'linkedin recruiter', 'jobboard', 'sourcing', 'assessment'],
            'Formation': ['lms', 'e-learning', 'moodle', 'articulate', 'captivate'],
            'Paie': ['silae', 'sage paie', 'adp', 'meta4', 'système paie'],
            'Évaluation': ['entretiens', 'feedback 360', 'assessment center', 'potentiel'],
            'Droit Social': ['droit travail', 'convention collective', 'code travail', 'prud\'hommes']
        },
        'Logistique': {
            'Systèmes WMS': ['wms', 'warehouse management', 'manhattan', 'sap wm', 'reflex'],
            'Supply Chain': ['planification', 'mrp', 'prévision', 'demand planning'],
            'Transport': ['tms', 'transport management', 'optimisation tournées'],
            'Outils Logistiques': ['sap mm', 'oracle wms', 'pkms', 'generix'],
            'Standards': ['lean', 'six sigma', '5s', 'kanban', 'flux tiré'],
            'Réglementation': ['adr', 'douane', 'incoterms', 'réglementation transport']
        },
        'Vente': {
            'CRM': ['salesforce', 'hubspot', 'pipedrive', 'zoho', 'dynamics'],
            'Prospection': ['sales navigator', 'hunter', 'lemlist', 'outreach'],
            'Analytics Commercial': ['tableau', 'power bi', 'qlikview', 'looker'],
            'E-commerce': ['shopify', 'magento', 'woocommerce', 'prestashop'],
            'Communication': ['slack', 'teams', 'zoom', 'gotomeeting'],
            'Méthodologies': ['spin selling', 'challenger sale', 'inbound', 'account based']
        },
        'Santé': {
            'Systèmes Médicaux': ['dpi', 'ris', 'pacs', 'lims', 'pharma'],
            'Réglementation': ['gdp', 'gcp', 'ich', 'fda', 'ema', 'ansm'],
            'Qualité': ['iso 13485', 'iso 15189', 'iso 27001', 'validation'],
            'Outils Statistiques': ['sas', 'r', 'spss', 'stata', 'clinical trials'],
            'Standards': ['hl7', 'dicom', 'snomed', 'icd-10'],
            'Domaines': ['pharmacovigilance', 'affaires réglementaires', 'clinical research']
        }
    }
    
    if domaine not in categories_par_domaine:
        # Domaine généraliste ou non spécifique
        return {
            'Outils et Logiciels': [],
            'Compétences Techniques': [],
            'Méthodologies': [],
            'Certifications': [],
            'Langues': []
        }
    
    texte_lower = texte_cv.lower()
    categories_trouvees = {}
    
    for categorie, keywords in categories_par_domaine[domaine].items():
        mots_trouves = [keyword for keyword in keywords if keyword in texte_lower]
        if mots_trouves:
            categories_trouvees[categorie] = mots_trouves
    
    return categories_trouvees

def calculer_score_adequation(dossier_competences: str, mots_cles_mission: Dict[str, List[str]]) -> float:
    """
    Calcule un score d'adéquation entre le dossier et la mission
    
    Args:
        dossier_competences: Texte du dossier de compétences
        mots_cles_mission: Mots-clés extraits de la mission
    
    Returns:
        float: Score entre 0 et 1
    """
    dossier_lower = dossier_competences.lower()
    
    score_total = 0
    count_total = 0
    
    for categorie, mots in mots_cles_mission.items():
        for mot in mots:
            count_total += 1
            if mot in dossier_lower:
                score_total += 1
    
    return score_total / count_total if count_total > 0 else 0

def formater_donnees_pour_template(donnees: Dict[Any, Any]) -> Dict[str, Any]:
    """
    Formate les données optimisées pour être compatibles avec le template
    
    Args:
        donnees: Données brutes de l'IA
    
    Returns:
        Dict formaté pour le template
    """
    donnees_formatees = {}
    
    # Champs obligatoires avec valeurs par défaut
    champs_requis = {
        'nom_consultant': 'Consultant',
        'titre_du_poste': 'Poste à définir',
        'points_forts': [],
        'niveaux_intervention': [],
        'formations': [],
        'connaissances': {},
        'hobbies_divers': {'langues': 'Français', 'hobbies': 'À définir'},
        'experiences': []
    }
    
    # Remplir avec les données existantes ou les valeurs par défaut
    for champ, defaut in champs_requis.items():
        donnees_formatees[champ] = donnees.get(champ, defaut)
    
    # Validation et nettoyage spécifique
    
    # S'assurer que les listes ne sont pas vides
    if not donnees_formatees['points_forts']:
        donnees_formatees['points_forts'] = ['Compétences techniques', 'Adaptabilité']
    
    if not donnees_formatees['niveaux_intervention']:
        donnees_formatees['niveaux_intervention'] = ['Développeur', 'Consultant']
    
    # S'assurer que hobbies_divers a la bonne structure
    if not isinstance(donnees_formatees['hobbies_divers'], dict):
        donnees_formatees['hobbies_divers'] = {
            'langues': 'Français',
            'hobbies': str(donnees_formatees['hobbies_divers'])
        }
    
    return donnees_formatees

def generer_rapport_optimisation(donnees_originales: str, donnees_optimisees: Dict[str, Any], 
                                 mots_cles_mission: Dict[str, List[str]]) -> Dict[str, Any]:
    """
    Génère un rapport d'optimisation détaillé
    
    Args:
        donnees_originales: Texte original du dossier
        donnees_optimisees: Données optimisées par l'IA
        mots_cles_mission: Mots-clés de la mission
    
    Returns:
        Dict contenant le rapport d'analyse
    """
    
    rapport = {
        'score_adequation': calculer_score_adequation(donnees_originales, mots_cles_mission),
        'technologies_identifiees': mots_cles_mission.get('technologies', []),
        'competences_identifiees': mots_cles_mission.get('competences', []),
        'nb_experiences': len(donnees_optimisees.get('experiences', [])),
        'nb_formations': len(donnees_optimisees.get('formations', [])),
        'nb_points_forts': len(donnees_optimisees.get('points_forts', [])),
        'categories_competences': list(donnees_optimisees.get('connaissances', {}).keys())
    }
    
    return rapport

def afficher_metriques_optimisation(rapport: Dict[str, Any]):
    """
    Affiche les métriques d'optimisation dans Streamlit
    
    Args:
        rapport: Rapport généré par generer_rapport_optimisation
    """
    
    st.subheader("📊 Métriques d'optimisation")
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        score_pct = rapport['score_adequation'] * 100
        st.metric(
            "Score d'adéquation", 
            f"{score_pct:.0f}%",
            delta=f"{'Excellent' if score_pct > 80 else 'Bon' if score_pct > 60 else 'Moyen'}"
        )
    
    with col2:
        st.metric("Technologies identifiées", len(rapport['technologies_identifiees']))
    
    with col3:
        st.metric("Expériences adaptées", rapport['nb_experiences'])
    
    with col4:
        st.metric("Points forts", rapport['nb_points_forts'])
    
    # Détails supplémentaires
    if rapport['technologies_identifiees']:
        st.markdown("**🔧 Technologies clés identifiées :**")
        st.markdown(", ".join(rapport['technologies_identifiees'][:10]))  # Limiter l'affichage
    
    if rapport['competences_identifiees']:
        st.markdown("**🎯 Compétences clés identifiées :**")
        st.markdown(", ".join(rapport['competences_identifiees'][:8]))

def gerer_erreurs_api(func):
    """
    Décorateur pour gérer les erreurs d'API de manière uniforme
    """
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            logger.error(f"Erreur dans {func.__name__}: {str(e)}")
            st.error(f"Une erreur est survenue: {str(e)}")
            return None
    return wrapper

def sauvegarder_historique(donnees: Dict[str, Any], nom_fichier: str):
    """
    Sauvegarde l'historique des générations pour suivi
    
    Args:
        donnees: Données du CV généré
        nom_fichier: Nom du fichier généré
    """
    try:
        import json
        import datetime
        
        historique = {
            'timestamp': datetime.datetime.now().isoformat(),
            'nom_consultant': donnees.get('nom_consultant', 'Inconnu'),
            'titre_poste': donnees.get('titre_du_poste', 'Inconnu'),
            'nom_fichier': nom_fichier,
            'nb_experiences': len(donnees.get('experiences', [])),
            'nb_formations': len(donnees.get('formations', []))
        }
        
        # En production, sauvegarder dans une base de données
        # Ici, on log simplement
        logger.info(f"CV généré: {json.dumps(historique)}")
        
    except Exception as e:
        logger.error(f"Erreur sauvegarde historique: {str(e)}")

# Constantes de configuration
CONFIG = {
    'TAILLE_MAX_FICHIER_MB': 10,
    'TYPES_MISSION_AUTORISES': ['pdf', 'txt'],
    'TYPES_CV_AUTORISES': ['pdf', 'docx'],
    'TYPES_TEMPLATE_AUTORISES': ['docx'],
    'OPENAI_MODEL_DEFAULT': 'gpt-4',
    'OPENAI_MAX_TOKENS': 4000,
    'OPENAI_TEMPERATURE': 0.3
}