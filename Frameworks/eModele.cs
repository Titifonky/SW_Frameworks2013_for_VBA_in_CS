using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace Framework_SW2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("66AE684E-5820-11E2-BCFB-5D046188709B")]
    public interface IeModele
    {
        ModelDoc2 SwModele { get; }
        eSldWorks SW { get; }
        eComposant Composant { get; }
        eAssemblage Assemblage { get; }
        ePiece Piece { get; }
        eDessin Dessin { get; }
        eGestDeConfigurations GestDeConfigurations { get; }
        eGestDeProprietes GestDeProprietes { get; }
        eGestDeSelection GestDeSelection { get; }
        TypeFichier_e TypeDuModele { get; }
        eFichierSW FichierSw { get; }
        Boolean EstActif { get; set; }
        void Activer(swRebuildOnActivation_e Reconstruire = swRebuildOnActivation_e.swUserDecision);
        void Sauver();
        void Fermer();
        void Redessiner();
        void Reconstruire();
        void ForcerAToutReconstruire();
        void ZoomEtendu();
        void EffacerLesSelections();
        void ActiverInterfaceUtilisateur(Boolean Activer);
        ArrayList ListeDesFonctions(String NomARechercher = "", String TypeDeLaFonction = "", Boolean AvecLesSousFonctions = false);
        eFonction DerniereFonction();
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("6AFCE66E-5820-11E2-B651-77046188709B")]
    [ProgId("Frameworks.eModele")]
    public class eModele : IeModele, IComparable<eModele>, IComparer<eModele>, IEquatable<eModele>
    {
#region "Variables locales"
        
        private Boolean _EstInitialise = false;

        private ModelDoc2 _SwModele = null;
        private eSldWorks _SW = null;
        private eComposant _Composant = null;
        private eAssemblage _Assemblage = null;
        private ePiece _Piece = null;
        private eDessin _Dessin = null;
        private eGestDeSelection _GestDeSelection = null;
        private eFichierSW _FichierSw = null;
        private int Erreur = 0;
        private int Warning = 0;

#endregion

#region "Constructeur\Destructeur"

        public eModele() { }

#endregion

#region "Propriétés"

        /// <summary>
        /// Retourne le modele ModleDoc2 associé.
        /// </summary>
        public ModelDoc2 SwModele
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                return _SwModele;
            }
        }

        /// <summary>
        /// Retourne le parent ExtSldWorks.
        /// </summary>
        public eSldWorks SW { get { Debug.Info(MethodBase.GetCurrentMethod());  return _SW; } }

        /// <summary>
        /// Retourne le composant ExtComposant lié au modele.
        /// </summary>
        public eComposant Composant { get { Debug.Info(MethodBase.GetCurrentMethod());  return _Composant; } internal set { Debug.Info(MethodBase.GetCurrentMethod());  _Composant = value; } }

        /// <summary>
        /// Retourne l'assemblage ExtAssemblage si celui ci est valide.
        /// </summary>
        public eAssemblage Assemblage
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());

                if (_Assemblage == null)
                {
                    _Assemblage = new eAssemblage();
                    _Assemblage.Init(this);
                }

                if (_Assemblage.EstInitialise)
                    return _Assemblage;

                return null;
            }
        }

        /// <summary>
        /// Retourne la pièce ExtPiece si celui ci est valide.
        /// </summary>
        public ePiece Piece
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());

                if (_Piece == null)
                {
                    _Piece = new ePiece();
                    _Piece.Init(this);
                }

                if (_Piece.EstInitialise)
                    return _Piece;

                return null;
            }
        }

        /// <summary>
        /// Retourne le dessin ExtDessin si celui ci est valide.
        /// </summary>
        public eDessin Dessin
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());

                if (_Dessin == null)
                {
                    _Dessin = new eDessin();
                    _Dessin.Init(this);
                }

                if (_Dessin.EstInitialise)
                    return _Dessin;
                
                return null;
            }
        }

        /// <summary>
        /// Retourne le gestionnaire de configuration GestDeConfigurations.
        /// </summary>
        public eGestDeConfigurations GestDeConfigurations
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());

                eGestDeConfigurations pGestDeConfigurations = new eGestDeConfigurations();

                if (pGestDeConfigurations.Init(this))
                    return pGestDeConfigurations;

                return null;
            }
        }

        /// <summary>
        /// Retourne le gestionnaire de propriétés GestDeProprietes.
        /// </summary>
        public eGestDeProprietes GestDeProprietes
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());

                eGestDeProprietes pGestDeProprietes = new eGestDeProprietes();

                if (pGestDeProprietes.Init(SwModele.Extension.get_CustomPropertyManager(""), this))
                    return pGestDeProprietes;

                return null;
            }
        }

        /// <summary>
        /// Retourne le gestionnaire de propriétés GestDeProprietes.
        /// </summary>
        public eGestDeSelection GestDeSelection
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());

                if (_GestDeSelection == null)
                {
                    _GestDeSelection = new eGestDeSelection();
                    _GestDeSelection.Init(SwModele.SelectionManager, this);
                }

                if (_GestDeSelection.EstInitialise)
                    return _GestDeSelection;

                return null;
            }
        }

        /// <summary>
        /// Retourne le type du modele.
        /// </summary>
        public TypeFichier_e TypeDuModele
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                switch (_SwModele.GetType())
                {
                    case (int)swDocumentTypes_e.swDocASSEMBLY:
                        return TypeFichier_e.cAssemblage;

                    case (int)swDocumentTypes_e.swDocPART:
                        return TypeFichier_e.cPiece;

                    case (int)swDocumentTypes_e.swDocDRAWING:
                        return TypeFichier_e.cDessin;

                    default:
                        return TypeFichier_e.cAutre;
                }
            }
        }

        /// <summary>
        /// Retourne l'objet ExtFichierSw
        /// </summary>
        public eFichierSW FichierSw { get { Debug.Info(MethodBase.GetCurrentMethod()); return _FichierSw; } }

        public Boolean EstActif
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());

                if (SW.Modele().Equals(this))
                    return true;

                return false;
            }
            set
            {
                Debug.Info(MethodBase.GetCurrentMethod());

                _SW.SwSW.ActivateDoc3(SwModele.GetPathName(), true, (int)swRebuildOnActivation_e.swUserDecision, Erreur);
                _Composant.Configuration.Activer();
            }
        }

        /// <summary>
        /// Fonction interne.
        /// Test l'initialisation de l'objet ExtModele.
        /// </summary>
        internal Boolean EstInitialise { get { Debug.Info(MethodBase.GetCurrentMethod());  return _EstInitialise; } }

#endregion

#region "Méthodes"

        /// <summary>
        /// Méthode interne.
        /// Initialiser l'objet ExtModele.
        /// </summary>
        /// <param name="SwModele"></param>
        /// <param name="Sw"></param>
        /// <returns></returns>
        internal Boolean Init(ModelDoc2 SwModele, eSldWorks Sw)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            if ((SwModele != null) && (Sw != null) && Sw.EstInitialise)
            {
                _SwModele = SwModele;
                _SW = Sw;
                Debug.Info(_SwModele.GetPathName());

                // On valide l'initialisation
                _EstInitialise = true;

                // On créer l'objet ExtFichierSw associé
                _FichierSw = new eFichierSW();
                if (_FichierSw.Init(Sw))
                    _FichierSw.Chemin = _SwModele.GetPathName();
                else
                    _EstInitialise = false;

                // Si c'est un assemblage ou une pièce, on va chercher le composant associé
                if ((TypeDuModele == TypeFichier_e.cAssemblage) || (TypeDuModele == TypeFichier_e.cPiece))
                {
                    Debug.Info("Referencement de la configuration");
                    Configuration pConfigActive = _SwModele.ConfigurationManager.ActiveConfiguration;
                    _FichierSw.Configuration = pConfigActive.Name;

                    Debug.Info("Referencement du composant");
                    _Composant = new eComposant();
                    if (_Composant.Init(pConfigActive.GetRootComponent3(false), this) == false)
                        _EstInitialise = false;
                }
            }
            else
            {
                Debug.Info("!!!!! Erreur d'initialisation");
            }

            return _EstInitialise;
        }

        internal void ReinitialiserComposant()
        {
            // Si c'est un assemblage ou une pièce, on va chercher le composant associé
            if ((TypeDuModele == TypeFichier_e.cAssemblage) || (TypeDuModele == TypeFichier_e.cPiece))
            {
                Debug.Info("Referencement de la configuration");
                Configuration pConfigActive = _SwModele.ConfigurationManager.ActiveConfiguration;
                _FichierSw.Configuration = pConfigActive.Name;

                Debug.Info("Referencement du composant");
                _Composant = new eComposant();
                if (_Composant.Init(pConfigActive.GetRootComponent3(false), this) == false)
                    _EstInitialise = false;
            }
        }

        /// <summary>
        /// Active le modele et le met au premier plan.
        /// </summary>
        public void Activer(swRebuildOnActivation_e Reconstruire = swRebuildOnActivation_e.swUserDecision)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            _SW.SwSW.ActivateDoc3(SwModele.GetPathName(), true, (int)Reconstruire, Erreur);
            if (TypeDuModele != TypeFichier_e.cDessin)
                _Composant.Configuration.Activer();

            ZoomEtendu();
        }

        /// <summary>
        /// Sauve le modele.
        /// </summary>
        public void Sauver()
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            SwModele.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref Erreur, ref Warning);
            _FichierSw.Chemin = SwModele.GetPathName();
        }

        /// <summary>
        /// Ferme le modele.
        /// </summary>
        public void Fermer()
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            _SW.SwSW.CloseDoc(SwModele.GetPathName());
        }

        /// <summary>
        /// Redessine le modele.
        /// </summary>
        public void Redessiner()
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            SwModele.ActiveView.GraphicsRedraw();
        }

        /// <summary>
        /// Reconstruit le modele.
        /// </summary>
        public void Reconstruire()
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            SwModele.EditRebuild3();
        }

        /// <summary>
        /// Force à tout reconstruire.
        /// </summary>
        public void ForcerAToutReconstruire()
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            SwModele.ForceRebuild3(false);
        }

        /// <summary>
        /// Zoom étendu du modele.
        /// </summary>
        public void ZoomEtendu()
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            SwModele.ViewZoomtofit2();
        }

        /// <summary>
        /// Effacer les selections
        /// </summary>
        public void EffacerLesSelections()
        {
            _SwModele.ClearSelection2(true);
        }

        /// <summary>
        /// Activer ou désactiver les éléments de l'interface pour accélérer les macros
        /// </summary>
        /// <param name="Activer"></param>
        public void ActiverInterfaceUtilisateur(Boolean Activer)
        {
            _SwModele.FeatureManager.EnableFeatureTree = Activer;
            _SwModele.FeatureManager.EnableFeatureTreeWindow = Activer;
            _SwModele.ConfigurationManager.EnableConfigurationTree = Activer;

            if (Activer)
                _SwModele.FeatureManager.UpdateFeatureTree();
        }

        /// <summary>
        /// Méthode interne
        /// Renvoi la liste des fonctions filtrée par les arguments.
        /// </summary>
        /// <param name="NomARechercher"></param>
        /// <param name="AvecLesSousFonctions"></param>
        /// <returns></returns>
        internal List<eFonction> ListListeDesFonctions(String NomARechercher = "", String TypeDeLaFonction = "", Boolean AvecLesSousFonctions = false)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            _Composant.Configuration.Activer();

            List<eFonction> pListeFonctions = new List<eFonction>();

            Feature pSwFonction = _SwModele.FirstFeature();

            while (pSwFonction != null)
            {
                eFonction pFonction = new eFonction();

                if ((Regex.IsMatch(pSwFonction.Name, NomARechercher))
                    && (Regex.IsMatch(pSwFonction.GetTypeName2(), TypeDeLaFonction))
                    && pFonction.Init(pSwFonction, this)
                    && !(pListeFonctions.Contains(pFonction))
                    )
                    pListeFonctions.Add(pFonction);

                if (AvecLesSousFonctions)
                {
                    Feature pSwSousFonction = pSwFonction.GetFirstSubFeature();

                    while (pSwSousFonction != null)
                    {
                        eFonction pSousFonction = new eFonction();

                        if ((Regex.IsMatch(pSwSousFonction.Name, NomARechercher))
                            && (Regex.IsMatch(pSwSousFonction.GetTypeName2(), TypeDeLaFonction))
                            && pSousFonction.Init(pSwSousFonction, this)
                            && !(pListeFonctions.Contains(pSousFonction))
                            )
                            pListeFonctions.Add(pSousFonction);

                        pSwSousFonction = pSwSousFonction.GetNextSubFeature();
                    }
                }

                pSwFonction = pSwFonction.GetNextFeature();
            }

#if SW2013
            if (EstActif)
            {
                List<TreeControlItem> pListeNoeuds = new List<TreeControlItem>();
                TreeControlItem pNoeudRacine = _SwModele.FeatureManager.GetFeatureTreeRootItem2((int)swFeatMgrPane_e.swFeatMgrPaneTopHidden);
                ScannerFonctionsFeatureManager(pNoeudRacine, pListeNoeuds, AvecLesSousFonctions);
                foreach(TreeControlItem pNoeud in pListeNoeuds)
                {
                    Feature pSwFonctionNoeud = pNoeud.Object;
                    eFonction pFonction = new eFonction();

                    if ((Regex.IsMatch(pSwFonctionNoeud.Name, NomARechercher))
                            && (Regex.IsMatch(pSwFonctionNoeud.GetTypeName2(), TypeDeLaFonction))
                            && pFonction.Init(pSwFonctionNoeud, this)
                            && !(pListeFonctions.Contains(pFonction))
                            )
                        pListeFonctions.Add(pFonction);
                    
                }
            }
#endif

            return pListeFonctions;
        }

        /// <summary>
        /// Scanne les fonctions du FeatureManager
        /// </summary>
        /// <param name="Noeud"></param>
        /// <param name="ListeNoeuds"></param>
        /// <param name="AvecLesSousFonctions"></param>
        private void ScannerFonctionsFeatureManager(TreeControlItem Noeud, List<TreeControlItem> ListeNoeuds, Boolean AvecLesSousFonctions)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            if (Noeud.ObjectType == (int)swTreeControlItemType_e.swFeatureManagerItem_Feature)
                ListeNoeuds.Add(Noeud);

            TreeControlItem pNoeud = Noeud.GetFirstChild();

            while (pNoeud != null)
            {
                if (pNoeud.ObjectType == (int)swTreeControlItemType_e.swFeatureManagerItem_Feature)
                    ListeNoeuds.Add(pNoeud);

                // On scanne dans tous les cas le dossier Tôlerie et le dossier Etat déplié
                if (AvecLesSousFonctions || (pNoeud.Text == "Tôlerie") || (pNoeud.Text == "Etat déplié"))
                    ScannerFonctionsFeatureManager(pNoeud, ListeNoeuds, AvecLesSousFonctions);

                pNoeud = pNoeud.GetNext();
            }
        }

        /// <summary>
        /// Renvoi la liste des fonctions filtrée par les arguments.
        /// </summary>
        /// <param name="NomARechercher"></param>
        /// <param name="AvecLesSousFonctions"></param>
        /// <returns></returns>
        public ArrayList ListeDesFonctions(String NomARechercher = "", String TypeDeLaFonction = "", Boolean AvecLesSousFonctions = false)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            List<eFonction> pListeFonctions = ListListeDesFonctions(NomARechercher, TypeDeLaFonction, AvecLesSousFonctions);
            ArrayList pArrayFonctions = new ArrayList();

            if (pListeFonctions.Count > 0)
                pArrayFonctions = new ArrayList(pListeFonctions);

            return pArrayFonctions;
        }

        /// <summary>
        /// Renvoi la dernière fonction crée
        /// </summary>
        /// <returns></returns>
        public eFonction DerniereFonction()
        {
            eFonction pFonc = new eFonction();
            if (pFonc.Init(_SwModele.Extension.GetLastFeatureAdded(), this))
                return pFonc;

            return null;
        }

#endregion

#region "Interfaces génériques"

        public int CompareTo(eModele Modele)
        {
            return _SwModele.GetPathName().CompareTo(Modele.SwModele.GetPathName());
        }

        public int Compare(eModele Modele1, eModele Modele2)
        {
            return Modele1.SwModele.GetPathName().CompareTo(Modele2.SwModele.GetPathName());
        }

        public Boolean Equals(eModele Modele)
        {
            return _SwModele.GetPathName().Equals(Modele.SwModele.GetPathName());
        }

#endregion
    }
}
