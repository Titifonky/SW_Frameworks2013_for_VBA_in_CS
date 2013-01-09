using System;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using System.Text.RegularExpressions;

namespace Framework2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("93F8AAEE-5820-11E2-A4D1-91046188709B")]
    public interface IExtConfiguration
    {
        Configuration swConfiguration { get; }
        ExtModele Modele { get; }
        String Nom { get; set; }
        TypeConfig_e TypeConfig { get; }
        ExtConfiguration ConfigurationParent { get; }
        //Boolean Init(Configuration Config, ExtModele Modele);
        Boolean Est(TypeConfig_e T);
        Boolean Supprimer();
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("9818CA82-5820-11E2-852A-92046188709B")]
    [ProgId("Frameworks.ExtConfiguration")]
    public class ExtConfiguration : IExtConfiguration
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

        public Configuration swConfiguration { get { return _SwConfiguration; } }

        public ExtModele Modele { get { return _Modele.Init(); } }

        public String Nom { get { return _SwConfiguration.Name; } set { _SwConfiguration.Name = value; } }

        public TypeConfig_e TypeConfig
        {
            get
            {
                TypeConfig_e T = 0;
                if (Regex.IsMatch(_SwConfiguration.Name, Constantes.CONFIG_DEPLIEE))
                    T = TypeConfig_e.cDepliee;
                else if (Regex.IsMatch(_SwConfiguration.Name, Constantes.CONFIG_PLIEE))
                    T = TypeConfig_e.cPliee;

                if (_SwConfiguration.IsDerived() != false)
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
                return pConfigParent.Init(swConfiguration.GetParent(), _Modele);
            }
        }

        #endregion

        #region "Méthodes"

        internal ExtConfiguration Init(Configuration Config, ExtModele Modele)
        {
            _MethodBase Methode = System.Reflection.MethodBase.GetCurrentMethod();

            if ((Config != null) && (Modele != null) && (Modele.Init() != null))
            {
                _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name);

                _SwConfiguration = Config;
                _Modele = Modele;
                _EstInitialise = true;

                return this;
            }

            _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name + " : Erreur d'initialisation");
            _EstInitialise = false;
            return null;
        }

        internal ExtConfiguration Init()
        {
            if (_EstInitialise)
                return this;
            else
                return null;
        }

        public Boolean Est(TypeConfig_e T)
        {
            return TypeConfig.HasFlag(T);
        }

        public Boolean Supprimer()
        {
            return _Modele.SwModele.DeleteConfiguration2(Nom);
        }

        #endregion
    }
}
