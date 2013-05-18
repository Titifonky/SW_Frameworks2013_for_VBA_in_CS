using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Reflection;

namespace Framework_SW2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("040632B0-59D0-11E2-9576-4F376288709B")]
    public interface IGestDeConfigurations
    {
        eModele Modele { get; }
        eConfiguration ConfigurationActive { get; }
        ArrayList ListerLesConfigs(TypeConfig_e TypeConfig = TypeConfig_e.cToutesLesTypesDeConfig, String NomConfigDeBase = "");
        eConfiguration ConfigurationAvecLeNom(String NomConfiguration);
        eConfiguration AjouterUneConfigurationDeBase(String NomConfiguration);
        void SupprimerConfiguration(String NomConfiguration);
        void SupprimerLesConfigurationsDepliee(String NomConfigurationPliee = "");
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("0C78B576-59D0-11E2-BA7E-53376288709B")]
    [ProgId("Frameworks.GestDeConfigurations")]
    public class GestDeConfigurations : IGestDeConfigurations
    {
        #region "Variables locales"
        
        private Boolean _EstInitialise = false;

        private eModele _Modele;
        #endregion

        #region "Constructeur\Destructeur"

        public GestDeConfigurations() { }

        #endregion

        #region "Propriétés"

        /// <summary>
        /// Retourne le parent ExtModele.
        /// </summary>
        public eModele Modele { get { Debug.Info(MethodBase.GetCurrentMethod());  return _Modele; } }

        /// <summary>
        /// Retourne la configuration active.
        /// </summary>
        public eConfiguration ConfigurationActive
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());

                eConfiguration pConfig = new eConfiguration();
                if (pConfig.Init(_Modele.SwModele.ConfigurationManager.ActiveConfiguration, _Modele))
                    return pConfig;
                
                return null;
            }
        }

        /// <summary>
        /// Fonction interne.
        /// Test l'initialisation de l'objet GestDeConfigurations.
        /// </summary>
        internal Boolean EstInitialise { get { Debug.Info(MethodBase.GetCurrentMethod());  return _EstInitialise; } }

        #endregion

        #region "Méthodes"

        /// <summary>
        /// Méthode interne.
        /// Initialiser l'objet GestDeConfigurations.
        /// </summary>
        /// <param name="Modele"></param>
        /// <returns></returns>
        internal Boolean Init(eModele Modele)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            if ((Modele != null) && Modele.EstInitialise)
            {
                _Modele = Modele;
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
        /// Renvoi la liste des configurations filtrée par les arguments.
        /// </summary>
        /// <param name="TypeConfig"></param>
        /// <param name="NomConfigDeBase"></param>
        /// <returns></returns>
        internal List<eConfiguration> ListListerLesConfigs(TypeConfig_e TypeConfig = TypeConfig_e.cToutesLesTypesDeConfig, String NomConfigDeBase = "")
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            List<eConfiguration> pListConfig = new List<eConfiguration>();

            foreach (String pNomConfig in _Modele.SwModele.GetConfigurationNames())
            {
                eConfiguration pConfig = ConfigurationAvecLeNom(pNomConfig);
                eConfiguration pConfigParent = pConfig.ConfigurationParent;
                string NomConfigParent = "";

                if (pConfigParent != null)
                    NomConfigParent = pConfigParent.Nom;

                if (pConfig.Est(TypeConfig) && Regex.IsMatch(NomConfigParent, NomConfigDeBase))
                    pListConfig.Add(pConfig);
            }
            return pListConfig;
        }

        /// <summary>
        /// Renvoi la liste des configurations filtrée par les arguments.
        /// </summary>
        /// <param name="TypeConfig"></param>
        /// <param name="NomConfigDeBase"></param>
        /// <returns></returns>
        public ArrayList ListerLesConfigs(TypeConfig_e TypeConfig = TypeConfig_e.cToutesLesTypesDeConfig, String NomConfigDeBase = "")
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            List<eConfiguration> pListeConfigs = ListListerLesConfigs(TypeConfig, NomConfigDeBase);
            ArrayList pArrayConfigs = new ArrayList();

            if (pListeConfigs.Count > 0)
                pArrayConfigs = new ArrayList(pListeConfigs);

            return pArrayConfigs;
        }

        /// <summary>
        /// Renvoi la configuration à partir du nom.
        /// </summary>
        /// <param name="NomConfiguration"></param>
        /// <returns></returns>
        public eConfiguration ConfigurationAvecLeNom(String NomConfiguration)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            eConfiguration pConfig = new eConfiguration();
            if (pConfig.Init(_Modele.SwModele.GetConfigurationByName(NomConfiguration), _Modele))
                return pConfig;

            return null;
        }

        /// <summary>
        /// Ajouter une configuration de base.
        /// </summary>
        /// <param name="NomConfiguration"></param>
        /// <returns></returns>
        public eConfiguration AjouterUneConfigurationDeBase(String NomConfiguration)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            eConfiguration pConfig = new eConfiguration();
            if (pConfig.Init(_Modele.SwModele.ConfigurationManager.AddConfiguration(NomConfiguration, NomConfiguration, "", 0, "", ""), _Modele))
                return pConfig;

            return null;
        }

        /// <summary>
        /// Supprimer une configuration à partir du nom.
        /// </summary>
        /// <param name="NomConfiguration"></param>
        /// <returns></returns>
        public void SupprimerConfiguration(String NomConfiguration)
        {
            Debug.Info(MethodBase.GetCurrentMethod());
            _Modele.SwModele.DeleteConfiguration2(NomConfiguration);
        }

        /// <summary>
        /// Supprimer les configurations dépliées
        /// </summary>
        /// <param name="NomConfigurationPliee"></param>
        public void SupprimerLesConfigurationsDepliee(String NomConfigurationPliee = "")
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            foreach(eConfiguration Config in ListListerLesConfigs(TypeConfig_e.cDepliee,NomConfigurationPliee))
            {
                Config.Supprimer();
            }
        }

        #endregion
    }
}
