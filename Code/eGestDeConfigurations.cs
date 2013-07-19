using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using SolidWorks.Interop.sldworks;

namespace Framework
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("040632B0-59D0-11E2-9576-4F376288709B")]
    public interface IeGestDeConfigurations
    {
        eModele Modele { get; }
        eConfiguration ConfigurationActive { get; }
        ArrayList ListerLesConfigs(TypeConfig_e TypeConfig = TypeConfig_e.cTous, String NomConfigDeBase = "");
        eConfiguration ConfigurationAvecLeNom(String NomConfiguration);
        eConfiguration AjouterUneConfigurationRacine(String NomConfiguration, Boolean Ecraser = false);
        Boolean ConfigurationExiste(String Nom);
        void SupprimerConfiguration(String NomConfiguration);
        void SupprimerLesConfigurationsDepliee(String NomConfigurationPliee = "");
        Boolean SupprimerTousLesEtatsAffichages();
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("0C78B576-59D0-11E2-BA7E-53376288709B")]
    [ProgId("Frameworks.eGestDeConfigurations")]
    public class eGestDeConfigurations : IeGestDeConfigurations
    {
#region "Variables locales"

        private Boolean _EstInitialise = false;

        private eModele _Modele = null;

#endregion

#region "Constructeur\Destructeur"

        public eGestDeConfigurations() { }

#endregion

#region "Propriétés"

        /// <summary>
        /// Retourne le parent ExtModele.
        /// </summary>
        public eModele Modele { get { Debug.Print(MethodBase.GetCurrentMethod()); return _Modele; } }

        /// <summary>
        /// Retourne la configuration active.
        /// </summary>
        public eConfiguration ConfigurationActive
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());

                eConfiguration pConfig = new eConfiguration();

                Configuration pSwConfiguration = _Modele.SwModele.ConfigurationManager.ActiveConfiguration;

                if (pConfig.Init(pSwConfiguration, _Modele))
                    return pConfig;

                return null;
            }
        }

        /// <summary>
        /// Fonction interne.
        /// Test l'initialisation de l'objet GestDeConfigurations.
        /// </summary>
        internal Boolean EstInitialise { get { Debug.Print(MethodBase.GetCurrentMethod()); return _EstInitialise; } }

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
            Debug.Print(MethodBase.GetCurrentMethod());

            if ((Modele != null) && Modele.EstInitialise)
            {
                _Modele = Modele;
                _EstInitialise = true;
            }
            else
            {
                Debug.Print("!!!!! Erreur d'initialisation");
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
        internal List<eConfiguration> ListListerLesConfigs(TypeConfig_e TypeConfig = TypeConfig_e.cTous, String NomConfigDeBase = "")
        {
            Debug.Print(MethodBase.GetCurrentMethod());

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
        public ArrayList ListerLesConfigs(TypeConfig_e TypeConfig = TypeConfig_e.cTous, String NomConfigDeBase = "")
        {
            Debug.Print(MethodBase.GetCurrentMethod());

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
            Debug.Print(MethodBase.GetCurrentMethod());

            Configuration pSwConfig = _Modele.SwModele.GetConfigurationByName(NomConfiguration);
            eConfiguration pConfig = new eConfiguration();

            if (pConfig.Init(pSwConfig, _Modele))
                return pConfig;

            return null;
        }

        /// <summary>
        /// Ajouter une configuration de base.
        /// </summary>
        /// <param name="NomConfiguration"></param>
        /// <returns></returns>
        public eConfiguration AjouterUneConfigurationRacine(String NomConfiguration, Boolean Ecraser = false)
        {
            Debug.Print(MethodBase.GetCurrentMethod());

            eConfiguration pConfig = ConfigurationAvecLeNom(NomConfiguration);

            if (pConfig != null)
            {
                if (Ecraser == true)
                    pConfig.Supprimer();
            }
            else
            {
                pConfig = new eConfiguration();
            }



            if (!pConfig.EstInitialise)
                pConfig.Init(_Modele.SwModele.ConfigurationManager.AddConfiguration(NomConfiguration, NomConfiguration, "", 0, "", ""), _Modele);

            return pConfig;
        }

        public Boolean ConfigurationExiste(String Nom)
        {
            if (ConfigurationAvecLeNom(Nom) != null)
                return true;

            return false;
        }

        /// <summary>
        /// Supprimer une configuration à partir du nom.
        /// </summary>
        /// <param name="NomConfiguration"></param>
        /// <returns></returns>
        public void SupprimerConfiguration(String NomConfiguration)
        {
            Debug.Print(MethodBase.GetCurrentMethod());

            Configuration pSwConfig = _Modele.SwModele.GetConfigurationByName(NomConfiguration);
            eConfiguration pConfig = new eConfiguration();
            if (pConfig.Init(pSwConfig, Modele))
                pConfig.Supprimer();
        }

        /// <summary>
        /// Supprimer les configurations dépliées
        /// </summary>
        /// <param name="NomConfigurationPliee"></param>
        public void SupprimerLesConfigurationsDepliee(String NomConfigurationPliee = "")
        {
            Debug.Print(MethodBase.GetCurrentMethod());

            ConfigurationActive.ConfigurationRacine.Activer();

            foreach (eConfiguration Config in ListListerLesConfigs(TypeConfig_e.cDepliee, NomConfigurationPliee))
            {
                Config.Supprimer();
            }
        }

        public Boolean SupprimerTousLesEtatsAffichages()
        {
            if (Modele.TypeDuModele == TypeFichier_e.cPiece)
                return Modele.Piece.SwPiece.RemoveAllDisplayStates();

            return false;
        }

#endregion
    }
}
