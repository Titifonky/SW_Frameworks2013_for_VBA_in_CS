using System;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;

namespace Framework_SW2013
{

    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("7EECBE39-3E38-49D7-A617-2E3AFEF915ED")]
    public interface IExtVue
    {
        View SwVue { get; }
        ExtFeuille Feuille { get; }
        String Nom { get; set; }
        ExtModele ModeleDeReference { get; }
        ExtConfiguration ConfigurationDeReference { get; }
        ExtDimensionVue Dimensions { get; }
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("0A46F613-F061-4FC2-8DAD-F4EA5BBBBD8E")]
    [ProgId("Frameworks.ExtVue")]
    public class ExtVue : IExtVue
    {
        #region "Variables locales"
        private Debug _Debug = Debug.Instance;
        private Boolean _EstInitialise = false;

        private ExtFeuille _Feuille;
        private View _SwVue;
        #endregion

        #region "Constructeur\Destructeur"

        public ExtVue() { }

        #endregion

        #region "Propriétés"

        public View SwVue { get { return _SwVue; } }

        public ExtFeuille Feuille { get { return _Feuille; } }

        public String Nom { get { return _SwVue.GetName2(); } set { _SwVue.SetName2(value); } }

        public ExtModele ModeleDeReference
        {
            get
            {
                ExtModele pModele = new ExtModele();
                if (pModele.Init(_SwVue.ReferencedDocument, _Feuille.Dessin.Modele.SW))
                    return pModele;

                return null;
            }
        }

        public ExtConfiguration ConfigurationDeReference
        {
            get
            {
                ExtConfiguration pConfig = new ExtConfiguration();
                if (pConfig.Init(_SwVue.ReferencedDocument.GetConfigurationByName(_SwVue.ReferencedConfiguration), ModeleDeReference))
                    return pConfig;

                return null;
            }
        }

        public ExtDimensionVue Dimensions
        {
            get
            {
                ExtDimensionVue pDimensions = new ExtDimensionVue();

                if (pDimensions.Init(this))
                    return pDimensions;

                return null;
            }
        }

        internal Boolean EstInitialise { get { return _EstInitialise; } }

        #endregion

        #region "Méthodes"

        internal Boolean Init(View SwVue, ExtFeuille Feuille)
        {
            _MethodBase Methode = System.Reflection.MethodBase.GetCurrentMethod();

            if ((SwVue != null) && (Feuille != null) && Feuille.EstInitialise)
            {
                _Feuille = Feuille;
                _SwVue = SwVue;

                _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name + " : " + this.Nom);
                _EstInitialise = true;
            }
            else
            {
                _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name + " : Erreur d'initialisation");
            }
            return _EstInitialise;
        }

        #endregion

    }
}
