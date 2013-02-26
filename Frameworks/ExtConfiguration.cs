using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using SolidWorks.Interop.sldworks;

/////////////////////////// Implementation terminée ///////////////////////////

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
        private Debug _Debug = Debug.Instance;
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

        public Configuration SwConfiguration { get { return _SwConfiguration; } }

        public ExtModele Modele { get { return _Modele; } }

        public String Nom { get { return SwConfiguration.Name; } set { SwConfiguration.Name = value; } }

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

        internal Boolean EstInitialise { get { return _EstInitialise; } }

        #endregion

        #region "Méthodes"

        internal Boolean Init(Configuration Config, ExtModele Modele)
        {
            _Debug.DebugAjouterLigne(this.GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name);

            if ((Config != null) && (Modele != null) && Modele.EstInitialise)
            {
                _SwConfiguration = Config;
                _Modele = Modele;

                _Debug.DebugAjouterLigne("\t -> " + this.Nom);
                _EstInitialise = true;
            }
            else
            {
                _Debug.DebugAjouterLigne("\t !!!!! Erreur d'initialisation");
            }

            return _EstInitialise;
        }

        public Boolean Est(TypeConfig_e T)
        {
            return Convert.ToBoolean(TypeConfig & T);
        }

        public Boolean Activer()
        {
            return Convert.ToBoolean(_Modele.SwModele.ShowConfiguration2(Nom));
        }

        public Boolean Supprimer()
        {
            return _Modele.SwModele.DeleteConfiguration2(Nom);
        }

        public ExtConfiguration AjouterUneConfigurationDerivee(String NomConfigDerivee)
        {
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
