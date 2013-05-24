using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
        void Activer();
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
        public ModelDoc2 SwModele { get { Debug.Info(MethodBase.GetCurrentMethod());  return _SwModele; } }

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

                eAssemblage Assemblage = new eAssemblage();

                if (Assemblage.Init(this))
                    return Assemblage;

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

                ePiece Piece = new ePiece();

                if (Piece.Init(this))
                    return Piece;

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

                eDessin Dessin = new eDessin();

                if (Dessin.Init(this))
                    return Dessin;
                
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

                eGestDeConfigurations pGestConfigs = new eGestDeConfigurations();
                if (pGestConfigs.Init(this))
                    return pGestConfigs;

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

                eGestDeProprietes pGestProps = new eGestDeProprietes();
                if (pGestProps.Init(SwModele.Extension.get_CustomPropertyManager(""), this))
                    return pGestProps;

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

                eGestDeSelection pGestSelection = new eGestDeSelection();
                if (pGestSelection.Init(SwModele.SelectionManager, this))
                    return pGestSelection;

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
                switch (SwModele.GetType())
                {
                    case (int)swDocumentTypes_e.swDocASSEMBLY:
                        return TypeFichier_e.cAssemblage;

                    case (int)swDocumentTypes_e.swDocPART:
                        return TypeFichier_e.cPiece;

                    case (int)swDocumentTypes_e.swDocDRAWING:
                        return TypeFichier_e.cDessin;

                    default:
                        return TypeFichier_e.cAucun;
                }
            }
        }

        /// <summary>
        /// Retourne l'objet ExtFichierSw
        /// </summary>
        public eFichierSW FichierSw
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());

                if (_FichierSw.EstInitialise)
                {
                    _FichierSw.Chemin = SwModele.GetPathName();
                    _FichierSw.Configuration = _Composant.Configuration.Nom;
                    _FichierSw.Nb = 1;
                    return _FichierSw;
                }
                return null;
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
                if (_FichierSw.Init(_SW))
                {
                    _FichierSw.Chemin = _SwModele.GetPathName();
                    _FichierSw.Configuration = _SwModele.ConfigurationManager.ActiveConfiguration.Name;
                    _FichierSw.Nb = 1;
                }

                // Si c'est un assemblage ou une pièce, on va chercher le composant associé
                if ((TypeDuModele == TypeFichier_e.cAssemblage) || (TypeDuModele == TypeFichier_e.cPiece))
                {
                    Debug.Info("Referencement du composant");
                    _Composant = new eComposant();
                    if (_Composant.Init(_SwModele.ConfigurationManager.ActiveConfiguration.GetRootComponent3(true), this) == false)
                        _EstInitialise = false;
                }
            }
            else
            {
                Debug.Info("!!!!! Erreur d'initialisation");
            }

            return _EstInitialise;
        }

        /// <summary>
        /// Active le modele et le met au premier plan.
        /// </summary>
        public void Activer()
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            _SW.SwSW.ActivateDoc3(SwModele.GetPathName(), true, 0, Erreur);
            ZoomEtendu();
            Redessiner();
        }

        /// <summary>
        /// Sauve le modele.
        /// </summary>
        public void Sauver()
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            SwModele.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref Erreur, ref Warning);
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

            List<eFonction> pListeFonctions = new List<eFonction>();

            Feature pSwFonction = _SwModele.FirstFeature();

            while (pSwFonction != null)
            {
                eFonction pFonction = new eFonction();

                if ((Regex.IsMatch(pSwFonction.Name, NomARechercher))
                    && (Regex.IsMatch(pSwFonction.GetTypeName2(), TypeDeLaFonction))
                    && pFonction.Init(pSwFonction, this)
                    && !(pListeFonctions.Contains(pFonction)))
                    pListeFonctions.Add(pFonction);

                if (AvecLesSousFonctions)
                {
                    Feature pSwSousFonction = pSwFonction.GetFirstSubFeature();

                    while (pSwSousFonction != null)
                    {
                        eFonction pSousFonction = new eFonction();

                        if ((Regex.IsMatch(pSwSousFonction.Name, NomARechercher))
                            && pSousFonction.Init(pSwSousFonction, this)
                            && !(pListeFonctions.Contains(pSousFonction)))
                            pListeFonctions.Add(pSousFonction);

                        pSwSousFonction = pSwSousFonction.GetNextSubFeature();
                    }
                }

                pSwFonction = pSwFonction.GetNextFeature();
            }

            return pListeFonctions.Distinct().ToList();

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
            eFonction pDerniereFonction = null;
            foreach (eFonction pFonction in ListListeDesFonctions())
            {
                if (pFonction.TypeDeLaFonction == "FlatPattern")
                    break;

                pDerniereFonction = pFonction;
            }

            return pDerniereFonction;
        }

        #endregion

        #region "Interfaces génériques"

        int IComparable<eModele>.CompareTo(eModele Modele)
        {
            return _SwModele.GetPathName().CompareTo(Modele.SwModele.GetPathName());
        }

        int IComparer<eModele>.Compare(eModele Modele1, eModele Modele2)
        {
            return Modele1.SwModele.GetPathName().CompareTo(Modele2.SwModele.GetPathName());
        }

        bool IEquatable<eModele>.Equals(eModele Modele)
        {
            return Modele.SwModele.GetPathName().Equals(_SwModele.GetPathName());
        }

        #endregion
    }
}
