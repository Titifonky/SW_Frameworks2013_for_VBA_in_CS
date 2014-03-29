using SolidWorks.Interop.sldworks;
using System;
using System.Collections;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace Framework
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("040632B0-59D0-11E2-9576-4F376288709B")]
    public interface IeGestDeConfigurations
    {
        eModele Modele { get; }
        eConfiguration ConfigurationActive { get; }
        Boolean LierLesAffichagesAuxConfigurations { get; set; }
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

        private static readonly String cNOMCLASSE = typeof(eGestDeConfigurations).Name;

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
        public eModele Modele { get { Log.Methode(cNOMCLASSE); return _Modele; } }

        /// <summary>
        /// Retourne la configuration active.
        /// </summary>
        public eConfiguration ConfigurationActive
        {
            get
            {
                Log.Methode(cNOMCLASSE);

                eConfiguration pConfig = new eConfiguration();

                Configuration pSwConfiguration = _Modele.SwModele.ConfigurationManager.ActiveConfiguration;

                if (pConfig.Init(pSwConfiguration, _Modele))
                    return pConfig;

                return null;
            }
        }

        public Boolean LierLesAffichagesAuxConfigurations
        {
            get
            {
                if (_Modele.SwModele.ConfigurationManager.LinkDisplayStatesToConfigurations == false)
                    return false;
                else
                    return true;
            }
            set
            {
                ConfigurationManager ConfigMgr = _Modele.SwModele.ConfigurationManager;
                if (value)
                    ConfigMgr.LinkDisplayStatesToConfigurations = true;
                else
                    ConfigMgr.LinkDisplayStatesToConfigurations = false;
            }
        }

        /// <summary>
        /// Fonction interne.
        /// Test l'initialisation de l'objet GestDeConfigurations.
        /// </summary>
        internal Boolean EstInitialise { get { Log.Methode(cNOMCLASSE); return _EstInitialise; } }

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
            Log.Methode(cNOMCLASSE);

            if ((Modele != null) && Modele.EstInitialise)
            {
                _Modele = Modele;
                _EstInitialise = true;
            }
            else
            {
                Log.Message("!!!!! Erreur d'initialisation");
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
        public ArrayList ListerLesConfigs(TypeConfig_e TypeConfig = TypeConfig_e.cTous, String NomConfigDeBase = "")
        {
            Log.Methode(cNOMCLASSE);

            ArrayList pListConfig = new ArrayList();

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
        /// Renvoi la configuration à partir du nom.
        /// </summary>
        /// <param name="NomConfiguration"></param>
        /// <returns></returns>
        public eConfiguration ConfigurationAvecLeNom(String NomConfiguration)
        {
            Log.Methode(cNOMCLASSE);

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
            Log.Methode(cNOMCLASSE);

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
            Log.Methode(cNOMCLASSE);

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
            Log.Methode(cNOMCLASSE);

            ConfigurationActive.ConfigurationRacine.Activer();

            foreach (eConfiguration Config in ListerLesConfigs(TypeConfig_e.cDepliee, NomConfigurationPliee))
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
