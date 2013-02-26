using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

/////////////////////////// Implementation terminée ///////////////////////////

namespace Framework_SW2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("66AE684E-5820-11E2-BCFB-5D046188709B")]
    public interface IExtModele
    {
        ModelDoc2 SwModele { get; }
        ExtSldWorks SW { get; }
        ExtComposant Composant { get; set; }
        ExtAssemblage Assemblage { get; }
        ExtPiece Piece { get; }
        ExtDessin Dessin { get; }
        GestDeConfigurations GestDeConfigurations { get; }
        GestDeProprietes GestDeProprietes { get; }
        TypeFichier_e TypeDuModele { get; }
        String Chemin { get; }
        String NomDuFichier { get; }
        String NomDuFichierSansExt { get; }
        String NomDuDossier { get; }
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
        private Debug _Debug = Debug.Instance;
        private Boolean _EstInitialise = false;

        private ModelDoc2 _SwModele;
        private ExtSldWorks _SW;
        private ExtComposant _Composant;
        private int Erreur = 0;
        private int Warning = 0;

        #endregion

        #region "Constructeur\Destructeur"

        public ExtModele() { }

        #endregion

        #region "Propriétés"

        public ModelDoc2 SwModele { get { return _SwModele; } }

        public ExtSldWorks SW { get { return _SW; } }

        public ExtComposant Composant { get { return _Composant; } set { _Composant = value; } }

        public ExtAssemblage Assemblage
        {
            get
            {
                _Debug.DebugAjouterLigne(this.GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name);

                ExtAssemblage Assemblage = new ExtAssemblage();

                if (Assemblage.Init(this))
                    return Assemblage;

                return null;
            }
        }

        public ExtPiece Piece
        {
            get
            {
                _Debug.DebugAjouterLigne(this.GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name);

                ExtPiece Piece = new ExtPiece();

                if (Piece.Init(this))
                    return Piece;

                return null;
            }
        }

        public ExtDessin Dessin
        {
            get
            {
                _Debug.DebugAjouterLigne(this.GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name);

                ExtDessin Dessin = new ExtDessin();

                if (Dessin.Init(this))
                    return Dessin;
                
                return null;
            }
        }

        public GestDeConfigurations GestDeConfigurations
        {
            get
            {
                _Debug.DebugAjouterLigne(this.GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name);

                GestDeConfigurations pGestConfigs = new GestDeConfigurations();
                if (pGestConfigs.Init(this))
                    return pGestConfigs;

                return null;
            }
        }

        public GestDeProprietes GestDeProprietes
        {
            get
            {
                _Debug.DebugAjouterLigne(this.GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name);

                GestDeProprietes pGestProps = new GestDeProprietes();
                if (pGestProps.Init(SwModele.Extension.get_CustomPropertyManager(""), this))
                    return pGestProps;

                return null;
            }
        }

        public TypeFichier_e TypeDuModele
        {
            get
            {
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

        public String Chemin { get { return SwModele.GetPathName(); } }
        public String NomDuFichier { get { return Path.GetFileName(SwModele.GetPathName()); } }
        public String NomDuFichierSansExt { get { return Path.GetFileNameWithoutExtension(SwModele.GetPathName()); } }
        public String NomDuDossier { get { return Path.GetDirectoryName(SwModele.GetPathName()); } }

        internal Boolean EstInitialise { get { return _EstInitialise; } }

        #endregion

        #region "Méthodes"

        internal Boolean Init(ModelDoc2 SwModele, ExtSldWorks Sw)
        {
            _Debug.DebugAjouterLigne(this.GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name);

            if ((SwModele != null) && (Sw != null))
            {
                _SwModele = SwModele;
                _SW = Sw;
                _Debug.DebugAjouterLigne("\t -> " + this.Chemin);

                // On valide l'initialisation
                _EstInitialise = true;

                // Si c'est un assemblage ou une pièce, on va chercher le composant associé
                if ((TypeDuModele == TypeFichier_e.cAssemblage) || (TypeDuModele == TypeFichier_e.cPiece))
                {
                    _Debug.DebugAjouterLigne("\t -> Referencement du composant");
                    _Composant = new ExtComposant();
                    if (_Composant.Init(_SwModele.ConfigurationManager.ActiveConfiguration.GetRootComponent3(true), this) == false)
                        _EstInitialise = false;
                }
            }
            else
            {
                _Debug.DebugAjouterLigne("\t !!!!! Erreur d'initialisation");
            }

            return _EstInitialise;
        }

        public void Activer()
        {
            _Debug.DebugAjouterLigne(this.GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name);

            _SW.SwSW.ActivateDoc3(SwModele.GetPathName(), true, 0, Erreur);
            ZoomEtendu();
            Redessiner();
        }

        public void Sauver()
        {
            _Debug.DebugAjouterLigne(this.GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name);

            SwModele.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref Erreur, ref Warning);
        }

        public void Fermer()
        {
            _Debug.DebugAjouterLigne(this.GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name);

            _SW.SwSW.CloseDoc(SwModele.GetPathName());
        }

        public void Redessiner()
        {
            _Debug.DebugAjouterLigne(this.GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name);

            SwModele.ActiveView.GraphicsRedraw();
        }

        public void Reconstruire()
        {
            _Debug.DebugAjouterLigne(this.GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name);

            SwModele.EditRebuild3();
        }

        public void ForcerAToutReconstruire()
        {
            _Debug.DebugAjouterLigne(this.GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name);

            SwModele.ForceRebuild3(false);
        }

        public void ZoomEtendu()
        {
            _Debug.DebugAjouterLigne(this.GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name);

            SwModele.ViewZoomtofit2();
        }

        internal List<ExtFonction> ListListeDesFonctions(String NomARechercher = "", Boolean AvecLesSousFonctions = false)
        {
            _Debug.DebugAjouterLigne(this.GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name);

            List<ExtFonction> pListeFonctions = new List<ExtFonction>();

            Feature pSwFonction = _SwModele.FirstFeature();

            while (pSwFonction != null)
            {
                ExtFonction pFonction = new ExtFonction();

                if ((Regex.IsMatch(pSwFonction.Name, NomARechercher)) && pFonction.Init(pSwFonction, this))
                    pListeFonctions.Add(pFonction);

                if (AvecLesSousFonctions)
                {
                    Feature pSwSousFonction = pSwFonction.GetFirstSubFeature();

                    while (pSwSousFonction != null)
                    {
                        ExtFonction pSousFonction = new ExtFonction();

                        if ((Regex.IsMatch(pSwFonction.Name, NomARechercher)) && pSousFonction.Init(pSwSousFonction, this))
                            pListeFonctions.Add(pSousFonction);

                        pSwSousFonction = pSwSousFonction.GetNextSubFeature();
                    }
                }

                pSwFonction = pSwFonction.GetNextFeature();
            }


            return pListeFonctions;

        }

        public ArrayList ListeDesFonctions(String NomARechercher = "", Boolean AvecLesSousFonctions = false)
        {
            _Debug.DebugAjouterLigne(this.GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name);

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
