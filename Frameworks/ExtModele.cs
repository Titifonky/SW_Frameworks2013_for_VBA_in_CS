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
        ExtSldWorks SW { get; set; }
        ModelDoc2 swModele { get; set; }
        ExtRecherche NouvelleRecherche();
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

        private ExtSldWorks _SW;
        private ModelDoc2 _swModele;
        private int Erreur = 0;
        private int Warning = 0;

        #endregion

        #region "Constructeur\Destructeur"

        public ExtModele()
        {
        }

        #endregion

        #region "Propriétés"

        public ExtSldWorks SW
        {
            get { return _SW; }
            set { _SW = value; }
        }

        public ModelDoc2 swModele
        {
            get { return _swModele; }
            set { _swModele = value; }
        }

        #endregion

        #region "Méthodes"

        public ExtRecherche NouvelleRecherche()
        {
            ExtRecherche pNouvelleRecherche = new ExtRecherche();
            return pNouvelleRecherche;
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
