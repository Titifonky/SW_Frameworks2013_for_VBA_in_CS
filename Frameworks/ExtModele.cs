using System;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Collections.Generic;

namespace Framework2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("66AE684E-5820-11E2-BCFB-5D046188709B")]
    public interface IExtModele
    {
        ModelDoc2 swModele { get; }
        ExtSldWorks SW { get; }
        ExtComposant Composant { get; set; }
        TypeFichier_e TypeDuModele { get; }
        String Chemin { get; }
        Boolean Init(ModelDoc2 ModeleDoc, ExtSldWorks Sw);
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
    public class ExtModele : IExtModele
    {
        #region "Variables locales"
        private Debug _Debug = Debug.Instance;

        private ModelDoc2 _swModele;
        private ExtSldWorks _SW;
        private ExtComposant _Composant;
        private int Erreur = 0;
        private int Warning = 0;

        #endregion

        #region "Constructeur\Destructeur"

        public ExtModele()
        {
        }

        #endregion

        #region "Propriétés"

        public ModelDoc2 swModele
        {
            get { return _swModele; }
        }

        public ExtSldWorks SW
        {
            get { return _SW; }
        }

        public ExtComposant Composant
        {
            get { return _Composant; }
            set { _Composant = value; }
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

        public String Chemin
        {
            get { return _swModele.GetPathName(); }
        }

        #endregion

        #region "Méthodes"

        public Boolean Init(ModelDoc2 ModeleDoc, ExtSldWorks Sw)
        {
            _MethodBase Methode = System.Reflection.MethodBase.GetCurrentMethod();

            if (!((ModeleDoc == null) && (Sw == null)))
            {
                _swModele = ModeleDoc;
                _SW = Sw;
                _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name);

                if ((TypeDuModele == TypeFichier_e.cAssemblage) || (TypeDuModele == TypeFichier_e.cPiece))
                {
                    _Debug.DebugAjouterLigne("\t" + this.GetType().Name + " -> " + "Referencement du composant");
                    _Composant = new ExtComposant();
                    _Composant.Init(_swModele.ConfigurationManager.ActiveConfiguration.GetRootComponent3(true), this);
                }
                return true;
            }

            _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name + " : Erreur d'initialisation");
            return false;
        }

        public void Activer()
        {
            _SW.swSW.ActivateDoc3(_swModele.GetPathName(), true, 0, Erreur);
            ZoomEtendu();
            Redessiner();
        }

        public void Sauver()
        {
            _swModele.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent, ref Erreur, ref Warning);
        }

        public void Fermer()
        {
            _SW.swSW.CloseDoc(_swModele.GetPathName());
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

        int IComparable<ExtModele>.CompareTo(ExtModele Modele)
        {
            return _swModele.GetPathName().CompareTo(Modele.swModele.GetPathName());
        }

        int IComparer<ExtModele>.Compare(ExtModele x, ExtModele y)
        {
            return x.swModele.GetPathName().CompareTo(y.swModele.GetPathName());
        }
    }
}
