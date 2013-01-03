using System;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace Frameworks2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("AB42D550-558F-11E2-801F-870D6288709B")]
    public interface IExtConfiguration
    {
        Configuration swConfiguration { get; }
        ExtModele Modele { get; }
        String Nom { get; set; }
        Boolean Init(Configuration Config, ExtModele Modele);
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("B1FC377E-558F-11E2-BF68-8E0D6288709B")]
    [ProgId("Frameworks.ExtConfiguration")]
    public class ExtConfiguration : IExtConfiguration
    {
        #region "Variables locales"

        private Configuration _swConfiguration;
        private ExtModele _Modele;

        #endregion

        #region "Constructeur\Destructeur"

        public ExtConfiguration()
        {
        }

        #endregion

        #region "Propriétés"

        public Configuration swConfiguration
        {
            get { return _swConfiguration; }
        }

        public ExtModele Modele
        {
            get { return _Modele; }
        }

        public String Nom
        {
            get { return _swConfiguration.Name; }
            set { _swConfiguration.Name = value; }
        }

        #endregion

        #region "Méthodes"

        public Boolean Init(Configuration Config, ExtModele Modele)
        {
            if (!((Config.Equals(null)) && (Modele.Equals(null))))
            {
                _swConfiguration = Config;
                _Modele = Modele;

                return true;
            }
            return false;
        }

        #endregion
    }
}
