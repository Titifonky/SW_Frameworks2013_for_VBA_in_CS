using System;
using System.Reflection;
using System.Runtime.InteropServices;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.sldworks;

namespace Framework_SW2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("FC29A423-E1E0-40D4-BC58-3CB893ECF9AF")]
    public interface IeObjet
    {
        eModele Modele { get; }
        dynamic Objet { get; }
        swSelectType_e TypeDeObjet { get; }
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("54AC1164-1C5C-4CB1-B839-5256537D7355")]
    [ProgId("Frameworks.eObjet")]
    public class eObjet : IeObjet
    {
        #region "Variables locales"

        private Boolean _EstInitialise = false;

        private eModele _Modele;
        private dynamic _SwObjet;
        private swSelectType_e _TypeObjet;

        #endregion

        #region "Constructeur\Destructeur"

        public eObjet() { }

        #endregion

        #region "Propriétés"

        /// <summary>
        /// Retourne le modele associé.
        /// </summary>
        public eModele Modele
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                return _Modele;
            }
        }

        /// <summary>
        /// Retourne l'objet associé.
        /// </summary>
        public dynamic Objet { get { Debug.Info(MethodBase.GetCurrentMethod()); return _SwObjet; } }

        /// <summary>
        /// Retourne le type de l'objet.
        /// </summary>
        public swSelectType_e TypeDeObjet
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                if (_EstInitialise)
                    return _TypeObjet;
                else
                    return swSelectType_e.swSelNOTHING;
            }
        }

        /// <summary>
        /// Fonction interne.
        /// Test l'initialisation de l'objet ExtCorps.
        /// </summary>
        internal Boolean EstInitialise { get { Debug.Info(MethodBase.GetCurrentMethod()); return _EstInitialise; } }

        #endregion

        #region "Méthodes"

        /// <summary>
        /// Méthode interne.
        /// Initialiser l'objet eObjet.
        /// </summary>
        /// <param name="Modele"></param>
        /// <param name="SwObjet"></param>
        /// <param name="TypeDeObjet"></param>
        /// <returns></returns>
        internal Boolean Init(eModele Modele, dynamic SwObjet, swSelectType_e TypeDeObjet)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            if ((SwObjet != null) && (Modele != null) && Modele.EstInitialise)
            {
                _Modele = Modele;
                _TypeObjet = TypeDeObjet;

                switch (TypeDeObjet)
                {
                    case swSelectType_e.swSelCOMPONENTS:
                        Component2 pSwComposant = SwObjet;
                        eComposant pComposant = new eComposant();
                        if (pComposant.Init(pSwComposant, Modele))
                        {
                            Modele.Composant = pComposant;
                            _SwObjet = pComposant;
                            _EstInitialise = true;
                        }
                        break;
                    case swSelectType_e.swSelCONFIGURATIONS:
                        Feature pSwFonction = SwObjet;
                        Configuration pSwConfiguration = pSwFonction.GetSpecificFeature2();
                        eConfiguration pConfiguration = new eConfiguration();
                        if (pConfiguration.Init(pSwConfiguration, Modele))
                        {
                            _SwObjet = pConfiguration;
                            _EstInitialise = true;
                        }
                        break;
                    case swSelectType_e.swSelDRAWINGVIEWS:
                        View pSwVue = SwObjet;
                        eVue pVue = new eVue();
                        if (pVue.Init(pSwVue, Modele))
                        {
                            _SwObjet = pVue;
                            _EstInitialise = true;
                        }
                        break;
                    case swSelectType_e.swSelSHEETS:
                        Sheet pSwFeuille = SwObjet;
                        eFeuille pFeuille = new eFeuille();
                        if (pFeuille.Init(pSwFeuille, Modele))
                        {
                            _SwObjet = pFeuille;
                            _EstInitialise = true;
                        }
                        break;
                    case swSelectType_e.swSelSOLIDBODIES:
                        Body2 pSwCorps = SwObjet;
                        eCorps pCorps = new eCorps();
                        if (pCorps.Init(pSwCorps, Modele))
                        {
                            _SwObjet = pCorps;
                            _EstInitialise = true;
                        }
                        break;
                    case swSelectType_e.swSelDATUMPLANES:
                    case swSelectType_e.swSelDATUMAXES:
                    case swSelectType_e.swSelDATUMPOINTS:
                    case swSelectType_e.swSelATTRIBUTES:
                    case swSelectType_e.swSelSKETCHES:
                    case swSelectType_e.swSelSECTIONLINES:
                    case swSelectType_e.swSelDETAILCIRCLES:
                    case swSelectType_e.swSelMATES:
                    case swSelectType_e.swSelBODYFEATURES:
                    case swSelectType_e.swSelREFCURVES:
                    case swSelectType_e.swSelREFERENCECURVES:
                    case swSelectType_e.swSelREFSILHOUETTE:
                    case swSelectType_e.swSelCAMERAS:
                    case swSelectType_e.swSelSWIFTANNOTATIONS:
                    case swSelectType_e.swSelSWIFTFEATURES:
                        eFonction pFonction = new eFonction();
                        if (pFonction.Init(SwObjet, Modele))
                        {
                            _SwObjet = pFonction;
                            _EstInitialise = true;
                        }
                        break;

                }
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
