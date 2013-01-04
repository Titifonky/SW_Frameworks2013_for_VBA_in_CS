using System;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace Frameworks2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("0F8FAD5B-D9CB-469F-A165-70867728950E")]
    public interface IExtModele
    {
        ModelDoc2 swModele { get; }
        ExtSldWorks SW { get; }
        ExtComposant Composant { get; set; }
        TypeFichier_e TypeDuModele { get; }
        String Chemin { get; }
        ExtRecherche NouvelleRecherche { get; }
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
    [Guid("7C9E6679-7425-40DE-944B-E07FC1F90AE7")]
    [ProgId("Frameworks.ExtModele")]
    public class ExtModele : IExtModele
    {
        #region "Variables locales"

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
                TypeFichier_e Type;

                switch (_swModele.GetType())
                {
                    case (int)swDocumentTypes_e.swDocASSEMBLY:
                        Type = TypeFichier_e.cAssemblage;
                        break;

                    case (int)swDocumentTypes_e.swDocPART:
                        Type = TypeFichier_e.cPiece;
                        break;

                    case (int)swDocumentTypes_e.swDocDRAWING:
                        Type = TypeFichier_e.cDessin;
                        break;

                    default:
                        Type = TypeFichier_e.cAucun;
                        break;
                }

                return Type;
            }
        }

        public String Chemin
        {
            get { return _swModele.GetPathName(); }
        }

        /// <summary>
        /// Renvoi un nouvel objet Recherche
        /// </summary>
        public ExtRecherche NouvelleRecherche
        {
            get
            {
                ExtRecherche pNouvelleRecherche = new ExtRecherche();
                return pNouvelleRecherche;
            }
        }

        #endregion

        #region "Méthodes"

        public Boolean Init(ModelDoc2 ModeleDoc, ExtSldWorks Sw)
        {

            if (!((ModeleDoc.Equals(null)) && (Sw.Equals(null))))
            {
                _swModele = ModeleDoc;
                _SW = Sw;
                if ((TypeDuModele == TypeFichier_e.cAssemblage) || (TypeDuModele == TypeFichier_e.cPiece))
                {
                    _Composant = new ExtComposant();
                    _Composant.Init(_swModele.ConfigurationManager.ActiveConfiguration.GetRootComponent3(true),this);
                }
                return true;
            }
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
            _swModele.Save3((int)swSaveAsOptions_e.swSaveAsOptions_Silent,ref Erreur, ref Warning);
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
    }
}
