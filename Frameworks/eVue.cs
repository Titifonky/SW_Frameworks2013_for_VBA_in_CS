using System;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using System.Reflection;

namespace Framework_SW2013
{

    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("7EECBE39-3E38-49D7-A617-2E3AFEF915ED")]
    public interface IeVue
    {
        View SwVue { get; }
        eFeuille Feuille { get; }
        String Nom { get; set; }
        eModele ModeleDeReference { get; }
        eConfiguration ConfigurationDeReference { get; }
        eDimensionVue Dimensions { get; }
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("0A46F613-F061-4FC2-8DAD-F4EA5BBBBD8E")]
    [ProgId("Frameworks.eVue")]
    public class eVue : IeVue
    {
        #region "Variables locales"
        
        private Boolean _EstInitialise = false;

        private eFeuille _Feuille;
        private View _SwVue;
        #endregion

        #region "Constructeur\Destructeur"

        public eVue() { }

        #endregion

        #region "Propriétés"

        /// <summary>
        /// Retourne l'objet View associé.
        /// </summary>
        public View SwVue { get { Debug.Info(MethodBase.GetCurrentMethod());  return _SwVue; } }

        /// <summary>
        /// Retourne le parent ExtFeuille.
        /// </summary>
        public eFeuille Feuille { get { Debug.Info(MethodBase.GetCurrentMethod());  return _Feuille; } }

        /// <summary>
        /// Retourne ou défini le nom de la feuille.
        /// </summary>
        public String Nom { get { Debug.Info(MethodBase.GetCurrentMethod());  return _SwVue.GetName2(); } set { Debug.Info(MethodBase.GetCurrentMethod());  _SwVue.SetName2(value); } }

        /// <summary>
        /// Retourne le modele ExtModele référencé par la vue.
        /// </summary>
        public eModele ModeleDeReference
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                eModele pModele = new eModele();
                if (pModele.Init(_SwVue.ReferencedDocument, _Feuille.Dessin.Modele.SW))
                    return pModele;

                return null;
            }
        }

        /// <summary>
        /// Retourne la configuration ExtConfiguration référencée par la vue.
        /// </summary>
        public eConfiguration ConfigurationDeReference
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                eConfiguration pConfig = new eConfiguration();
                if (pConfig.Init(_SwVue.ReferencedDocument.GetConfigurationByName(_SwVue.ReferencedConfiguration), ModeleDeReference))
                    return pConfig;

                return null;
            }
        }

        /// <summary>
        /// Retourne les dimensions de la vue.
        /// </summary>
        public eDimensionVue Dimensions
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                eDimensionVue pDimensions = new eDimensionVue();

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
        internal Boolean Init(View SwVue, eFeuille Feuille)
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

        /// <summary>
        /// Méthode interne.
        /// Initialiser l'objet ExtVue.
        /// </summary>
        /// <param name="SwVue"></param>
        /// <param name="Modele"></param>
        /// <returns></returns>
        internal Boolean Init(View SwVue, eModele Modele)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            if ((SwVue != null) && (Modele != null) && Modele.EstInitialise && (Modele.TypeDuModele == TypeFichier_e.cDessin))
            {
                eFeuille Feuille = new eFeuille();
                if (Feuille.Init(SwVue.Sheet, Modele.Dessin))
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
            }
            return _EstInitialise;
        }

        #endregion

    }
}
