using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using SolidWorks.Interop.sldworks;
using System.Reflection;

namespace Framework_SW2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("93F8AAEE-5820-11E2-A4D1-91046188709B")]
    public interface IExtConfiguration
    {
        Configuration SwConfiguration { get; }
        ExtModele Modele { get; }
        String Nom { get; set; }
        TypeConfig_e TypeConfig { get; }
        ExtConfiguration ConfigurationParent { get; }
        ExtConfiguration ConfigurationRacine { get; }
        GestDeProprietes GestDeProprietes { get; }
        Boolean Est(TypeConfig_e T);
        Boolean Activer();
        Boolean Supprimer();
        ExtConfiguration AjouterUneConfigurationDerivee(String Nom);
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("9818CA82-5820-11E2-852A-92046188709B")]
    [ProgId("Frameworks.ExtConfiguration")]
    public class ExtConfiguration : IExtConfiguration, IComparable<ExtConfiguration>, IComparer<ExtConfiguration>, IEquatable<ExtConfiguration>
    {
        #region "Variables locales"
        
        private Boolean _EstInitialise = false;

        private Configuration _SwConfiguration;
        private ExtModele _Modele;

        #endregion

        #region "Constructeur\Destructeur"

        public ExtConfiguration()
        {
        }

        #endregion

        #region "Propriétés"

        /// <summary>
        /// Retourne l'objet Configuration associé.
        /// </summary>
        public Configuration SwConfiguration { get { Debug.Info(MethodBase.GetCurrentMethod());  return _SwConfiguration; } }

        /// <summary>
        /// Retourne le parent ExtModele.
        /// </summary>
        public ExtModele Modele { get { Debug.Info(MethodBase.GetCurrentMethod());  return _Modele; } }

        /// <summary>
        /// Retourne ou défini le nom de la configuration.
        /// </summary>
        public String Nom { get { Debug.Info(MethodBase.GetCurrentMethod());  return SwConfiguration.Name; } set { Debug.Info(MethodBase.GetCurrentMethod());  SwConfiguration.Name = value; } }

        /// <summary>
        /// Retourne le type de la configuration.
        /// </summary>
        public TypeConfig_e TypeConfig
        {
            get
            {
                TypeConfig_e T = 0;
                if (Regex.IsMatch(SwConfiguration.Name, Constantes.CONFIG_DEPLIEE))
                    T = TypeConfig_e.cDepliee;
                else if (Regex.IsMatch(SwConfiguration.Name, Constantes.CONFIG_PLIEE))
                    T = TypeConfig_e.cPliee;

                if (SwConfiguration.IsDerived() != false)
                    T |= TypeConfig_e.cDerivee;
                else
                    T |= TypeConfig_e.cDeBase;

                return T;
            }
        }

        /// <summary>
        /// Retourne la configuration parent.
        /// </summary>
        public ExtConfiguration ConfigurationParent
        {
            get
            {
                ExtConfiguration pConfigParent = new ExtConfiguration();
                if ((TypeConfig == TypeConfig_e.cDerivee) && pConfigParent.Init(SwConfiguration.GetParent(), _Modele))
                    return pConfigParent;

                return null;
            }
        }

        /// <summary>
        /// Retourne la configuration racine, c'est la configuration parent la plus haute.
        /// </summary>
        public ExtConfiguration ConfigurationRacine
        {
            get
            {
                // Si elle est derivée, on lance la recherche
                if (Est(TypeConfig_e.cDerivee))
                {
                    ExtConfiguration pConfig = ConfigurationParent;

                    // Tant que pConfig est derivee, on boucle.
                    while (pConfig.Est(TypeConfig_e.cDerivee))
                    {
                        pConfig = pConfig.ConfigurationParent;
                    }

                    // Et on renvoi la dernière
                    return pConfig;
                }

                // Si elle n'est pas derivée, on la renvoi
                return this;
            }
        }

        /// <summary>
        /// Retourne le getionnaire de propriétés.
        /// </summary>
        public GestDeProprietes GestDeProprietes
        {
            get
            {
                GestDeProprietes pGestProps = new GestDeProprietes();
                if (pGestProps.Init(SwConfiguration.CustomPropertyManager, _Modele))
                    return pGestProps;

                return null;
            }
        }

        /// <summary>
        /// Fonction interne.
        /// Test l'initialisation de l'objet ExtConfiguration.
        /// </summary>
        internal Boolean EstInitialise { get { Debug.Info(MethodBase.GetCurrentMethod());  return _EstInitialise; } }

        #endregion

        #region "Méthodes"

        /// <summary>
        /// Méthode interne.
        /// Initialiser l'objet ExtConfiguration.
        /// </summary>
        /// <param name="Config"></param>
        /// <param name="Modele"></param>
        /// <returns></returns>
        internal Boolean Init(Configuration Config, ExtModele Modele)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            if ((Config != null) && (Modele != null) && Modele.EstInitialise)
            {
                _SwConfiguration = Config;
                _Modele = Modele;

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
        /// Test si la configuration est de type T.
        /// </summary>
        /// <param name="T"></param>
        /// <returns></returns>
        public Boolean Est(TypeConfig_e T)
        {
            Debug.Info(MethodBase.GetCurrentMethod());
            return Convert.ToBoolean(TypeConfig & T);
        }

        /// <summary>
        /// Active la configurationS.
        /// </summary>
        /// <returns></returns>
        public Boolean Activer()
        {
            Debug.Info(MethodBase.GetCurrentMethod());
            return Convert.ToBoolean(_Modele.SwModele.ShowConfiguration2(Nom));
        }

        /// <summary>
        /// Supprime la configuration.
        /// </summary>
        /// <returns></returns>
        public Boolean Supprimer()
        {
            Debug.Info(MethodBase.GetCurrentMethod());
            return _Modele.SwModele.DeleteConfiguration2(Nom);
        }

        /// <summary>
        /// Ajoute une configuration dérivée.
        /// </summary>
        /// <param name="NomConfigDerivee"></param>
        /// <returns></returns>
        public ExtConfiguration AjouterUneConfigurationDerivee(String NomConfigDerivee)
        {
            Debug.Info(MethodBase.GetCurrentMethod());
            ExtConfiguration pConfig = new ExtConfiguration();
            Configuration pSwConfig = _Modele.SwModele.ConfigurationManager.AddConfiguration(NomConfigDerivee, NomConfigDerivee, "", 0, NomConfigDerivee, "");

            if (pConfig.Init(pSwConfig ,_Modele))
                return pConfig;

            return null;
        }

        #endregion

        #region "Interfaces génériques"

        int IComparable<ExtConfiguration>.CompareTo(ExtConfiguration Conf)
        {
            String Nom1 = _Modele.SwModele.GetPathName() + _SwConfiguration.Name;
            String Nom2 = Conf._Modele.SwModele.GetPathName() + Conf._SwConfiguration.Name;
            return Nom1.CompareTo(Nom2);
        }

        int IComparer<ExtConfiguration>.Compare(ExtConfiguration Conf1, ExtConfiguration Conf2)
        {
            String Nom1 = Conf1._Modele.SwModele.GetPathName() + Conf1._SwConfiguration.Name;
            String Nom2 = Conf2._Modele.SwModele.GetPathName() + Conf2._SwConfiguration.Name;
            return Nom1.CompareTo(Nom2);
        }

        bool IEquatable<ExtConfiguration>.Equals(ExtConfiguration Conf)
        {
            String Nom1 = _Modele.SwModele.GetPathName() + _SwConfiguration.Name;
            String Nom2 = Conf._Modele.SwModele.GetPathName() + Conf._SwConfiguration.Name;
            return Nom1.Equals(Nom2);
        }

        #endregion
    }
}
