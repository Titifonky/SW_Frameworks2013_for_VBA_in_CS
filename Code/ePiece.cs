using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace Framework
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("57E51812-5820-11E2-82D5-34046188709B")]
    public interface IePiece
    {
        PartDoc SwPiece { get; }
        eModele Modele { get; }
        String Materiau { get; set; }
        Boolean Contient(TypeCorps_e T);
        eParametreTolerie ParametresDeTolerie { get; }
        void NumeroterDossier(Boolean Reinitialiser = false, Boolean Complet = false);
        ArrayList ListeDesCorps(String NomARechercher = "", TypeCorps_e TypeDeCorps = TypeCorps_e.cTous, Boolean PrendreEnCompteCache = false);
        ArrayList ListeDesDossiersDePiecesSoudees(TypeCorps_e TypeDeCorps = TypeCorps_e.cTous, Boolean PrendreEnCompteExclus = false);
        ArrayList ListeDesFonctionsDeArbre(String NomARechercher = "", String TypeDeLaFonction = "", Boolean AvecLesSousFonctions = false);
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("5E46FC3E-5820-11E2-86E5-38046188709B")]
    [ProgId("Frameworks.ePiece")]
    public class ePiece : IePiece
    {
        #region "Variables locales"

        private static readonly String cNOMCLASSE = typeof(ePiece).Name;

        private Boolean _EstInitialise = false;

        private eModele _Modele = null;
        private PartDoc _SwPiece = null;
        private eParametreTolerie _ParamTolerie = null;

        #endregion

        #region "Constructeur\Destructeur"

        public ePiece() { }

        #endregion

        #region "Propriétés"

        /// <summary>
        /// Renvoi l'objet PartDoc.
        /// </summary>
        public PartDoc SwPiece { get { Log.Methode(cNOMCLASSE); return _SwPiece; } }

        /// <summary>
        /// Renvoi l'objet ExtModele.
        /// </summary>
        public eModele Modele { get { Log.Methode(cNOMCLASSE); return _Modele; } }

        public String Materiau
        {
            get
            {
                String Db = "";
                String pMateriau = _SwPiece.GetMaterialPropertyName2(_Modele.GestDeConfigurations.ConfigurationActive.Nom, out Db);
                if (String.IsNullOrEmpty(pMateriau))
                    pMateriau = CONSTANTES.MATERIAUX_NON_SPECIFIE;
                return pMateriau;
            }
            set
            {
                String[] pBaseDeDonnees = _Modele.SW.SwSW.GetMaterialDatabases();

                // On test si pour chaque Base de donnée si le matériau à bien été appliqué.
                // Si oui, on sort de la boucle
                foreach (String Bdd in pBaseDeDonnees)
                {
                    _SwPiece.SetMaterialPropertyName2(_Modele.GestDeConfigurations.ConfigurationActive.Nom, Bdd, value);
                    if (Materiau == value)
                        break;
                }
            }
        }

        /// <summary>
        /// Retourne les parametres de tôlerie
        /// </summary>
        public eParametreTolerie ParametresDeTolerie
        {
            get
            {
                Log.Methode(cNOMCLASSE);

                if (_ParamTolerie == null)
                {
                    _ParamTolerie = new eParametreTolerie();
                    _ParamTolerie.Init(this);
                }

                if (_ParamTolerie.EstInitialise)
                    return _ParamTolerie;

                return null;
            }
        }

        /// <summary>
        /// Renvoi la valeur de l'initialisation.
        /// </summary>
        internal Boolean EstInitialise { get { Log.Methode(cNOMCLASSE); return _EstInitialise; } }

        #endregion

        #region "Méthodes"

        /// <summary>
        /// Méthode interne.
        /// Initialiser l'objet pièce.
        /// </summary>
        /// <param name="Modele"></param>
        /// <returns></returns>
        internal Boolean Init(eModele Modele)
        {
            Log.Methode(cNOMCLASSE);

            if ((Modele != null) && Modele.EstInitialise && (Modele.TypeDuModele == TypeFichier_e.cPiece))
            {
                Log.Message(Modele.FichierSw.Chemin);

                _Modele = Modele;
                _SwPiece = Modele.SwModele as PartDoc;
                _EstInitialise = true;
            }
            else
            {
                Log.Message("!!!!! Erreur d'initialisation");
            }

            return _EstInitialise;
        }

        /// <summary>
        /// Renvoi la fonction du dossier contenant les corps.
        /// </summary>
        /// <returns></returns>
        internal Feature DossierDesCorps()
        {
            Log.Methode(cNOMCLASSE);

            Feature pFonctionPiecesSoudees = _SwPiece.FirstFeature();

            while (pFonctionPiecesSoudees != null)
            {
                if (pFonctionPiecesSoudees.GetTypeName2() == "SolidBodyFolder")
                {
                    BodyFolder pDossier = pFonctionPiecesSoudees.GetSpecificFeature2();
                    pDossier.UpdateCutList();
                    return pFonctionPiecesSoudees;
                }

                pFonctionPiecesSoudees = pFonctionPiecesSoudees.GetNextFeature();
            }

            return null;
        }

        /// <summary>
        /// Renvoi Vrai si la piece contient des corps du type T
        /// A tester. Peut dans certain cas renvoyer un resultat erroné. ex :
        /// Si un corps de tolerie ou un profil a été créé puis supprimé,
        /// la fonction existe mais plus le corps. Pb également dans le cas de corps combiné
        /// </summary>
        /// <param name="T"></param>
        /// <returns></returns>
        public Boolean Contient(TypeCorps_e T)
        {
            Log.Methode(cNOMCLASSE);

            if (T.HasFlag(TypeCorps_e.cTole))
            {
                foreach (eFonction Fonction in _Modele.ListeDesFonctions())
                {
                    if ((Fonction.TypeDeLaFonction == "SMBaseFlange")
                        || (Fonction.TypeDeLaFonction == "SolidToSheetMetal")
                        || (Fonction.TypeDeLaFonction == "SheetMetal")
                        || (Fonction.TypeDeLaFonction == "FlatPattern"))
                        return true;
                }
            }

            if (T.HasFlag(TypeCorps_e.cBarre))
            {
                if (_Modele.ListeDesFonctions("", "WeldMemberFeat", false).Count > 0)
                    return true;
            }

            if (T.HasFlag(TypeCorps_e.cAutre))
            {
                if (ListeDesDossiersDePiecesSoudees(TypeCorps_e.cAutre, false).Count > 0)
                    return true;
            }

            return false;
        }

        /// <summary>
        /// Ajoute un No à chaque dossier
        /// </summary>
        public void NumeroterDossier(Boolean Reinitialiser = false, Boolean Complet = false)
        {
            int pNoDossierMax = 0;

            ArrayList ListeConfigs = new ArrayList();

            eConfiguration ConfCourante = _Modele.GestDeConfigurations.ConfigurationActive;

            if (Complet != false)
            {
                ListeConfigs = _Modele.GestDeConfigurations.ListerLesConfigs();
            }
            else
            {
                ListeConfigs.Add(_Modele.GestDeConfigurations.ConfigurationActive);
            }

            foreach (eConfiguration pConfig in ListeConfigs)
            {
                pConfig.Activer();

                foreach (eDossier pDossier in ListeDesDossiersDePiecesSoudees(TypeCorps_e.cTous, true))
                {
                    eGestDeProprietes pGestProps = pDossier.GestDeProprietes;

                    if (Reinitialiser != false)
                    {
                        pGestProps.SupprimerPropriete(CONSTANTES.NO_DOSSIER);
                    }
                    else if (pGestProps.ProprieteExiste(CONSTANTES.NO_DOSSIER))
                    {
                        ePropriete pProp = pGestProps.RecupererPropriete(CONSTANTES.NO_DOSSIER);
                        int pNoDossier = Convert.ToInt32(pProp.Valeur);
                        if (pNoDossier > pNoDossierMax)
                            pNoDossierMax = pNoDossier;
                    }
                }
            }

            foreach (eConfiguration pConfig in ListeConfigs)
            {
                pConfig.Activer();

                foreach (eDossier pDossier in ListeDesDossiersDePiecesSoudees(TypeCorps_e.cTous, true))
                {
                    eGestDeProprietes pGestProps = pDossier.GestDeProprietes;
                    if (!pGestProps.ProprieteExiste(CONSTANTES.NO_DOSSIER))
                    {
                        pNoDossierMax++;
                        pGestProps.AjouterPropriete(CONSTANTES.NO_DOSSIER, swCustomInfoType_e.swCustomInfoText, pNoDossierMax.ToString());
                    }
                    else if (Reinitialiser != false)
                    {
                        pNoDossierMax++;

                        ePropriete pProp = pGestProps.RecupererPropriete(CONSTANTES.NO_DOSSIER);
                        pProp.Expression = pNoDossierMax.ToString();
                    }
                        
                }
            }

            ConfCourante.Activer();
        }

        /// <summary>
        /// Méthode interne.
        /// Renvoi la liste des corps de la pièces filtrée par les arguments.
        /// </summary>
        /// <param name="TypeDeCorps"></param>
        /// <param name="PrendreEnCompteCache"></param>
        /// <returns></returns>
        public ArrayList ListeDesCorps(String NomARechercher = "", TypeCorps_e TypeDeCorps = TypeCorps_e.cTous, Boolean PrendreEnCompteCache = false)
        {
            Log.Methode(cNOMCLASSE);

            ArrayList Liste = new ArrayList();

            Object[] TableauDesCorps = _SwPiece.GetBodies2((int)swBodyType_e.swAllBodies, !PrendreEnCompteCache);

            if (TableauDesCorps.Length > 0)
            {
                foreach (Object ObjetCorps in TableauDesCorps)
                {
                    Body2 pSwCorps = (Body2)ObjetCorps;
                    eCorps pCorps = new eCorps();
                    if (pCorps.Init(pSwCorps, this) && TypeDeCorps.HasFlag(pCorps.TypeDeCorps) && Regex.IsMatch(pSwCorps.Name, NomARechercher))
                    {
                        Liste.Add(pCorps);
                    }
                }
            }

            return Liste;
        }

        /// <summary>
        /// Méthode interne
        /// Renvoi la liste des dossiers de pièces soudées de la pièce filtrée par les arguments.
        /// </summary>
        /// <param name="TypeDeCorps"></param>
        /// <param name="PrendreEnCompteExclus"></param>
        /// <returns></returns>
        public ArrayList ListeDesDossiersDePiecesSoudees(TypeCorps_e TypeDeCorps = TypeCorps_e.cTous, Boolean PrendreEnCompteExclus = false)
        {
            Log.Methode(cNOMCLASSE);

            ArrayList Liste = new ArrayList();

            Feature pFonction = DossierDesCorps();

            // S'il n'y a pas de liste, on arrete là
            if (pFonction == null)
                return null;

            pFonction = pFonction.GetFirstSubFeature();

            while (pFonction != null)
            {
                if (pFonction.GetTypeName2() == "CutListFolder")
                {
                    BodyFolder pSwDossier = pFonction.GetSpecificFeature2();
                    eDossier Dossier = new eDossier();

                    if (Dossier.Init(pSwDossier, this) && TypeDeCorps.HasFlag(Dossier.TypeDeCorps) && (!Dossier.EstExclu | PrendreEnCompteExclus))
                        Liste.Add(Dossier);

                    Dossier = null;
                }

                pFonction = pFonction.GetNextSubFeature();
            }

            return Liste;
        }

        /// <summary>
        /// Scanne les fonctions du FeatureManager
        /// </summary>
        /// <param name="Noeud"></param>
        /// <param name="ListeFonctions"></param>
        /// <param name="AvecLesSousFonctions"></param>
        private void ScannerFonctionsFeatureManager(TreeControlItem Noeud, ArrayList ListeFonctions, String NomARechercher, String TypeDeLaFonction, Boolean AvecLesSousFonctions)
        {
            Log.Methode(cNOMCLASSE);

            TreeControlItem pNoeud = Noeud.GetFirstChild();

            while (pNoeud != null)
            {
                eFonction pFonction = new eFonction();
                if (pNoeud.ObjectType == (int)swTreeControlItemType_e.swFeatureManagerItem_Feature)
                {
                    if (pFonction.Init(pNoeud.Object, Modele)
                        && Regex.IsMatch(pFonction.Nom, NomARechercher)
                        && Regex.IsMatch(pFonction.TypeDeLaFonction, TypeDeLaFonction))
                        ListeFonctions.Add(pFonction);
                }

                // On scanne dans tous les cas le dossier Tôlerie et le dossier Etat déplié
                // On rajoute un test d'initialisation de la fonction pour éviter les erreurs
                if (AvecLesSousFonctions ||
                    (pFonction.EstInitialise && (pFonction.TypeDeLaFonction == "TemplateSheetMetal"))
                    || (pFonction.EstInitialise && (pFonction.TypeDeLaFonction == "TemplateFlatPattern")))
                    ScannerFonctionsFeatureManager(pNoeud, ListeFonctions, NomARechercher, TypeDeLaFonction, AvecLesSousFonctions);

                pNoeud = pNoeud.GetNext();
            }
        }

        public ArrayList ListeDesFonctionsDeArbre(String NomARechercher = "", String TypeDeLaFonction = "", Boolean AvecLesSousFonctions = false)
        {
            ArrayList pListeFonction = new ArrayList();

            ScannerFonctionsFeatureManager(Modele.GestDeFonction_NoeudRacine(), pListeFonction, NomARechercher, TypeDeLaFonction, AvecLesSousFonctions);

            return pListeFonction;
        }

        internal eCorps CorpsDeplie()
        {
            ArrayList pListe = ListeDesCorps(CONSTANTES.NOM_CORPS_DEPLIEE);
            if (pListe.Count == 0)
                return null;

            eCorps pCorps = (eCorps)pListe[0];
            return pCorps;
        }

        #endregion

    }
}
