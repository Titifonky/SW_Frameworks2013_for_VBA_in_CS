﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using System.Text.RegularExpressions;

namespace Framework2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("040632B0-59D0-11E2-9576-4F376288709B")]
    public interface IGestDeConfigurations
    {
        ExtModele Modele { get; }
        ExtConfiguration ConfigurationActive { get; }
        Boolean Init(ExtModele Modele);
        ArrayList ListerLesConfigs(TypeConfig_e TypeConfig = TypeConfig_e.cToutesLesTypesDeConfig, String NomConfigDeBase = "");
        ExtConfiguration Configuration(String NomConfiguration);
        ExtConfiguration AjouterUneConfigurationDeBase(String NomConfiguration);
        void SupprimerLesConfigurationsDepliee(String NomConfigurationPliee = "");
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("0C78B576-59D0-11E2-BA7E-53376288709B")]
    [ProgId("Frameworks.GestDeConfigurations")]
    public class GestDeConfigurations : IGestDeConfigurations
    {
        #region "Variables locales"
        private Debug _Debug = Debug.Instance;

        ExtModele _Modele;

        #endregion

        #region "Constructeur\Destructeur"
        #endregion

        #region "Propriétés"

        public ExtModele Modele { get { return _Modele; } }

        public ExtConfiguration ConfigurationActive
        {
            get
            {
                Configuration pSwConfig = _Modele.SwModele.ConfigurationManager.ActiveConfiguration;
                if (pSwConfig == null)
                    return null;
                else
                {
                    ExtConfiguration pConfig = new ExtConfiguration();
                    pConfig.Init(pSwConfig, _Modele);
                    return pConfig;
                }
            }
        }

        #endregion

        #region "Méthodes"

        public Boolean Init(ExtModele Modele)
        {
            _MethodBase Methode = System.Reflection.MethodBase.GetCurrentMethod();

            if (Modele != null)
            {
                _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name);

                _Modele = Modele;
                return true;
            }

            _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name + " : Erreur d'initialisation");
            return false;
        }

        internal List<ExtConfiguration> ListListerLesConfigs(TypeConfig_e TypeConfig = TypeConfig_e.cToutesLesTypesDeConfig, String NomConfigDeBase = "")
        {
            List<ExtConfiguration> pListConfig = new List<ExtConfiguration>();

            foreach (String pNomConfig in _Modele.SwModele.GetConfigurationNames())
            {
                ExtConfiguration pConfig = Configuration(pNomConfig);
                ExtConfiguration pConfigParent = pConfig.ConfigurationParent;
                string NomConfigParent = "";

                if (pConfigParent != null)
                    NomConfigParent = pConfigParent.Nom;

                if (pConfig.Est(TypeConfig) && Regex.IsMatch(NomConfigParent, NomConfigDeBase))
                    pListConfig.Add(pConfig);
            }
            return pListConfig;
        }

        public ArrayList ListerLesConfigs(TypeConfig_e TypeConfig = TypeConfig_e.cToutesLesTypesDeConfig, String NomConfigDeBase = "")
        {
            _MethodBase Methode = System.Reflection.MethodBase.GetCurrentMethod();
            _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name);

            List<ExtConfiguration> pListeConfigs = ListListerLesConfigs(TypeConfig, NomConfigDeBase);
            ArrayList pArrayConfigs = new ArrayList();

            if (pListeConfigs.Count > 0)
                pArrayConfigs = new ArrayList(pListeConfigs);

            return pArrayConfigs;
        }

        public ExtConfiguration Configuration(String NomConfiguration)
        {
            Configuration pSwConfig = _Modele.SwModele.GetConfigurationByName(NomConfiguration);

            if (pSwConfig != null)
            {
                ExtConfiguration pConfig = new ExtConfiguration();
                pConfig.Init(pSwConfig, _Modele);
                return pConfig;
            }

            return null;
        }

        public ExtConfiguration AjouterUneConfigurationDeBase(String NomConfig)
        {
            ExtConfiguration pConfig = new ExtConfiguration();
            pConfig.Init(_Modele.SwModele.ConfigurationManager.AddConfiguration(NomConfig,NomConfig,"",0,"",""), _Modele);
            return pConfig;
        }

        public void SupprimerLesConfigurationsDepliee(String NomConfigurationPliee = "")
        {
            foreach(ExtConfiguration Config in ListListerLesConfigs(TypeConfig_e.cDepliee,NomConfigurationPliee))
            {
                
            }
        }

        #endregion
    }
}
