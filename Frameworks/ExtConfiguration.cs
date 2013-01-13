using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using SolidWorks.Interop.sldworks;

namespace Framework_SW2013
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
        Boolean Est(TypeConfig_e T);
        Boolean Supprimer();
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

        public Configuration swConfiguration { get { return _SwConfiguration; } }

        public ExtModele Modele { get { return _Modele; } }

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
                if (pConfigParent.Init(swConfiguration.GetParent(), _Modele))
                    return pConfigParent;

                return null;
            }
        }

        internal Boolean EstInitialise { get { return _EstInitialise; } }

        #endregion

        #region "Méthodes"

        internal Boolean Init(Configuration Config, ExtModele Modele)
        {
            _MethodBase Methode = System.Reflection.MethodBase.GetCurrentMethod();

            if ((Config != null) && (Modele != null) && Modele.EstInitialise)
            {
                _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name);

                _SwConfiguration = Config;
                _Modele = Modele;
                _EstInitialise = true;
            }
            else
            {
                _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name + " : Erreur d'initialisation");
            }

            return _EstInitialise;
        }

        public Boolean Est(TypeConfig_e T)
        {
            return Convert.ToBoolean(TypeConfig & T);
        }

        public Boolean Supprimer()
        {
            return _Modele.SwModele.DeleteConfiguration2(Nom);
        }

        #endregion

        #region "Interfaces génériques"

        int IComparable<ExtConfiguration>.CompareTo(ExtConfiguration Conf)
        {
            String Nom1 = Modele.Chemin + Nom;
            String Nom2 = Conf._Modele.Chemin + Conf.Nom;
            return Nom1.CompareTo(Nom2);
        }

        int IComparer<ExtConfiguration>.Compare(ExtConfiguration Conf1, ExtConfiguration Conf2)
        {
            String Nom1 = Conf1._Modele.Chemin + Conf1.Nom;
            String Nom2 = Conf2._Modele.Chemin + Conf2.Nom;
            return Nom1.CompareTo(Nom2);
        }

        bool IEquatable<ExtConfiguration>.Equals(ExtConfiguration Conf)
        {
            String Nom1 = Modele.Chemin + Nom;
            String Nom2 = Conf._Modele.Chemin + Conf.Nom;
            return Nom1.Equals(Nom2);
        }

        #endregion
    }
}
