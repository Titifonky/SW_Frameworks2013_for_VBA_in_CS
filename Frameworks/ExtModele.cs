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
    public interface IExtModele
    {
        ModelDoc2 SwModele { get; }
        ExtSldWorks SW { get; }
        ExtComposant Composant { get; }
        ExtAssemblage Assemblage { get; }
        ExtPiece Piece { get; }
        ExtDessin Dessin { get; }
        GestDeConfigurations GestDeConfigurations { get; }
        GestDeProprietes GestDeProprietes { get; }
        GestDeSelection GestDeSelection { get; }
        TypeFichier_e TypeDuModele { get; }
        ExtFichierSW FichierSw { get; }
        //String Chemin { get; }
        //String NomDuFichier { get; }
        //String NomDuFichierSansExt { get; }
        //String NomDuDossier { get; }
        void Activer();
        void Sauver();
        void Fermer();
        void Redessiner();
        void Reconstruire();
        void ForcerAToutReconstruire();
        void ZoomEtendu();
        ArrayList ListeDesFonctions(String NomARechercher = "", Boolean AvecLesSousFonctions = false);
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("6AFCE66E-5820-11E2-B651-77046188709B")]
    [ProgId("Frameworks.ExtModele")]
    public class ExtModele : IExtModele, IComparable<ExtModele>, IComparer<ExtModele>, IEquatable<ExtModele>
    {
        #region "Variables locales"
        
        private Boolean _EstInitialise = false;

        private ModelDoc2 _SwModele;
        private ExtSldWorks _SW;
        private ExtComposant _Composant;
        private ExtFichierSW _FichierSw;
        private int Erreur = 0;
        private int Warning = 0;

        #endregion

        #region "Constructeur\Destructeur"

        public ExtModele() { }

        #endregion

        #region "Propriétés"

        /// <summary>
        /// Retourne le modele ModleDoc2 associé.
        /// </summary>
        public ModelDoc2 SwModele { get { Debug.Info(MethodBase.GetCurrentMethod());  return _SwModele; } }

        /// <summary>
        /// Retourne le parent ExtSldWorks.
        /// </summary>
        public ExtSldWorks SW { get { Debug.Info(MethodBase.GetCurrentMethod());  return _SW; } }

        /// <summary>
        /// Retourne le composant ExtComposant lié au modele.
        /// </summary>
        public ExtComposant Composant { get { Debug.Info(MethodBase.GetCurrentMethod());  return _Composant; } internal set { Debug.Info(MethodBase.GetCurrentMethod());  _Composant = value; } }

        /// <summary>
        /// Retourne l'assemblage ExtAssemblage si celui ci est valide.
        /// </summary>
        public ExtAssemblage Assemblage
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());

                ExtAssemblage Assemblage = new ExtAssemblage();

                if (Assemblage.Init(this))
                    return Assemblage;

                return null;
            }
        }

        /// <summary>
        /// Retourne la pièce ExtPiece si celui ci est valide.
        /// </summary>
        public ExtPiece Piece
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());

                ExtPiece Piece = new ExtPiece();

                if (Piece.Init(this))
                    return Piece;

                return null;
            }
        }

        /// <summary>
        /// Retourne le dessin ExtDessin si celui ci est valide.
        /// </summary>
        public ExtDessin Dessin
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());

                ExtDessin Dessin = new ExtDessin();

                if (Dessin.Init(this))
                    return Dessin;
                
                return null;
            }
        }

        /// <summary>
        /// Retourne le gestionnaire de configuration GestDeConfigurations.
        /// </summary>
        public GestDeConfigurations GestDeConfigurations
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());

                GestDeConfigurations pGestConfigs = new GestDeConfigurations();
                if (pGestConfigs.Init(this))
                    return pGestConfigs;

                return null;
            }
        }

        /// <summary>
        /// Retourne le gestionnaire de propriétés GestDeProprietes.
        /// </summary>
        public GestDeProprietes GestDeProprietes
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());

                GestDeProprietes pGestProps = new GestDeProprietes();
                if (pGestProps.Init(SwModele.Extension.get_CustomPropertyManager(""), this))
                    return pGestProps;

                return null;
            }
        }

        /// <summary>
        /// Retourne le gestionnaire de propriétés GestDeProprietes.
        /// </summary>
        public GestDeSelection GestDeSelection
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());

                GestDeSelection pGestSelection = new GestDeSelection();
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
        public ExtFichierSW FichierSw
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
        internal Boolean Init(ModelDoc2 SwModele, ExtSldWorks Sw)
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
                _FichierSw = new ExtFichierSW();
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
                    _Composant = new ExtComposant();
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
            //Redessiner();
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
        /// Méthode interne
        /// Renvoi la liste des fonctions filtrée par les arguments.
        /// </summary>
        /// <param name="NomARechercher"></param>
        /// <param name="AvecLesSousFonctions"></param>
        /// <returns></returns>
        internal List<ExtFonction> ListListeDesFonctions(String NomARechercher = "", Boolean AvecLesSousFonctions = false)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            List<ExtFonction> pListeFonctions = new List<ExtFonction>();

            Feature pSwFonction = _SwModele.FirstFeature();

            while (pSwFonction != null)
            {
                ExtFonction pFonction = new ExtFonction();

                if ((Regex.IsMatch(pSwFonction.Name, NomARechercher))
                    && pFonction.Init(pSwFonction, this)
                    && !(pListeFonctions.Contains(pFonction)))
                    pListeFonctions.Add(pFonction);

                if (AvecLesSousFonctions)
                {
                    Feature pSwSousFonction = pSwFonction.GetFirstSubFeature();

                    while (pSwSousFonction != null)
                    {
                        ExtFonction pSousFonction = new ExtFonction();

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
        public ArrayList ListeDesFonctions(String NomARechercher = "", Boolean AvecLesSousFonctions = false)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            List<ExtFonction> pListeFonctions = ListListeDesFonctions(NomARechercher, AvecLesSousFonctions);
            ArrayList pArrayFonctions = new ArrayList();

            if (pListeFonctions.Count > 0)
                pArrayFonctions = new ArrayList(pListeFonctions);

            return pArrayFonctions;
        }

        #endregion

        #region "Interfaces génériques"

        int IComparable<ExtModele>.CompareTo(ExtModele Modele)
        {
            return _SwModele.GetPathName().CompareTo(Modele.SwModele.GetPathName());
        }

        int IComparer<ExtModele>.Compare(ExtModele Modele1, ExtModele Modele2)
        {
            return Modele1.SwModele.GetPathName().CompareTo(Modele2.SwModele.GetPathName());
        }

        bool IEquatable<ExtModele>.Equals(ExtModele Modele)
        {
            return Modele.SwModele.GetPathName().Equals(_SwModele.GetPathName());
        }

        #endregion
    }
}
