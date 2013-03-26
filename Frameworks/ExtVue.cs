using System;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using System.Reflection;

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
        
        private Boolean _EstInitialise = false;

        private ExtFeuille _Feuille;
        private View _SwVue;
        #endregion

        #region "Constructeur\Destructeur"

        public ExtVue() { }

        #endregion

        #region "Propriétés"

        /// <summary>
        /// Retourne l'objet View associé.
        /// </summary>
        public View SwVue { get { Debug.Info(MethodBase.GetCurrentMethod());  return _SwVue; } }

        /// <summary>
        /// Retourne le parent ExtFeuille.
        /// </summary>
        public ExtFeuille Feuille { get { Debug.Info(MethodBase.GetCurrentMethod());  return _Feuille; } }

        /// <summary>
        /// Retourne ou défini le nom de la feuille.
        /// </summary>
        public String Nom { get { Debug.Info(MethodBase.GetCurrentMethod());  return _SwVue.GetName2(); } set { Debug.Info(MethodBase.GetCurrentMethod());  _SwVue.SetName2(value); } }

        /// <summary>
        /// Retourne le modele ExtModele référencé par la vue.
        /// </summary>
        public ExtModele ModeleDeReference
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                ExtModele pModele = new ExtModele();
                if (pModele.Init(_SwVue.ReferencedDocument, _Feuille.Dessin.Modele.SW))
                    return pModele;

                return null;
            }
        }

        /// <summary>
        /// Retourne la configuration ExtConfiguration référencée par la vue.
        /// </summary>
        public ExtConfiguration ConfigurationDeReference
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                ExtConfiguration pConfig = new ExtConfiguration();
                if (pConfig.Init(_SwVue.ReferencedDocument.GetConfigurationByName(_SwVue.ReferencedConfiguration), ModeleDeReference))
                    return pConfig;

                return null;
            }
        }

        /// <summary>
        /// Retourne les dimensions de la vue.
        /// </summary>
        public ExtDimensionVue Dimensions
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                ExtDimensionVue pDimensions = new ExtDimensionVue();

                if (pDimensions.Init(this))
                    return pDimensions;

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
        /// Initialiser l'objet ExtVue.
        /// </summary>
        /// <param name="SwVue"></param>
        /// <param name="Feuille"></param>
        /// <returns></returns>
        internal Boolean Init(View SwVue, ExtFeuille Feuille)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            if ((SwVue != null) && (Feuille != null) && Feuille.EstInitialise)
            {
                _Feuille = Feuille;
                _SwVue = SwVue;

                Debug.Info(this.Nom);
                _EstInitialise = true;
            }
            else
            {
                Debug.Info("!!!!! Erreur d'initialisation");
            }
            return _EstInitialise;
        }

        #endregion

    }
}
