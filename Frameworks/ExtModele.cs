using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.IO;

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
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("6AFCE66E-5820-11E2-B651-77046188709B")]
    [ProgId("Frameworks.ExtModele")]
    public class ExtModele : IExtModele, IComparable<ExtModele>, IComparer<ExtModele>, IEquatable<ExtModele>
    {
        #region "Variables locales"
        private Debug _Debug = Debug.Instance;
        private Boolean _EstInitialise = false;

        private ModelDoc2 _swModele;
        private ExtSldWorks _SW;
        private ExtComposant _Composant;
        private int Erreur = 0;
        private int Warning = 0;

        #endregion

        #region "Constructeur\Destructeur"

        public ExtModele(){}

        #endregion

        #region "Propriétés"

        public ModelDoc2 SwModele { get { return _swModele; } }

        public ExtSldWorks SW { get { return _SW; } }

        public ExtComposant Composant { get { return _Composant; } set { _Composant = value; } }

        public ExtAssemblage Assemblage
        {
            get
            {
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
                if (TypeDuModele == TypeFichier_e.cDessin)
                    return new ExtDessin();
                else
                    return null;
            }
        }

        public GestDeConfigurations GestDeConfigurations
        {
            get
            {
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
                GestDeProprietes pGestProps = new GestDeProprietes();
                if (pGestProps.Init(_swModele.Extension.get_CustomPropertyManager("") ,this))
                    return pGestProps;

                return null;
            }
        }

        public TypeFichier_e TypeDuModele
        {
            get
            {
                switch (_swModele.GetType())
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

        public String Chemin { get { return _swModele.GetPathName(); } }
        public String NomDuFichier { get { return Path.GetFileName(_swModele.GetPathName()); } }
        public String NomDuFichierSansExt { get { return Path.GetFileNameWithoutExtension(_swModele.GetPathName()); } }
        public String NomDuDossier { get { return Path.GetDirectoryName(_swModele.GetPathName()); } }

        internal Boolean EstInitialise { get { return _EstInitialise; } }

        #endregion

        #region "Méthodes"

        internal Boolean Init(ModelDoc2 SwModele, ExtSldWorks Sw)
        {
            _MethodBase Methode = System.Reflection.MethodBase.GetCurrentMethod();

            if ((SwModele != null) && (Sw != null))
            {
                _swModele = SwModele;
                _SW = Sw;
                _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name);

                // On valide l'initialisation
                _EstInitialise = true;

                // Si c'est un assemblage ou une pièce, on va chercher le composant associé
                if ((TypeDuModele == TypeFichier_e.cAssemblage) || (TypeDuModele == TypeFichier_e.cPiece))
                {
                    _Debug.DebugAjouterLigne("\t" + this.GetType().Name + " -> " + "Referencement du composant");
                    _Composant = new ExtComposant();
                    if (_Composant.Init(_swModele.ConfigurationManager.ActiveConfiguration.GetRootComponent3(true), this) == false)
                        _EstInitialise = false;
                }
            }
            else
            {
                _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name + " : Erreur d'initialisation");
            }

            return _EstInitialise;
        }

        public void Activer()
        {
            _SW.SwSW.ActivateDoc3(_swModele.GetPathName(), true, 0, Erreur);
            ZoomEtendu();
            Redessiner();
        }

        public void Sauver()
        {
            _swModele.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref Erreur, ref Warning);
        }

        public void Fermer()
        {
            _SW.SwSW.CloseDoc(_swModele.GetPathName());
        }

        public void Redessiner()
        {
            _swModele.ActiveView.GraphicsRedraw();
        }

        public void Reconstruire()
        {
            _swModele.EditRebuild3();
        }

        public void ForcerAToutReconstruire()
        {
            _swModele.ForceRebuild3(false);
        }

        public void ZoomEtendu()
        {
            _swModele.ViewZoomtofit2();
        }

        #endregion

        #region "Interfaces génériques"

        int IComparable<ExtModele>.CompareTo(ExtModele Modele)
        {
            return _swModele.GetPathName().CompareTo(Modele.SwModele.GetPathName());
        }

        int IComparer<ExtModele>.Compare(ExtModele Modele1, ExtModele Modele2)
        {
            return Modele1.SwModele.GetPathName().CompareTo(Modele2.SwModele.GetPathName());
        }

        bool IEquatable<ExtModele>.Equals(ExtModele Modele)
        {
            return Modele.SwModele.GetPathName().Equals(_swModele.GetPathName());
        }

        #endregion
    }
}
