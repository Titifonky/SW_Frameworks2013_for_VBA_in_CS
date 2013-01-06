using System;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace Framework2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("93F8AAEE-5820-11E2-A4D1-91046188709B")]
    public interface IExtConfiguration
    {
        Configuration swConfiguration { get; }
        ExtModele Modele { get; }
        String Nom { get; set; }
        Boolean Init(Configuration Config, ExtModele Modele);
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("9818CA82-5820-11E2-852A-92046188709B")]
    [ProgId("Frameworks.ExtConfiguration")]
    public class ExtConfiguration : IExtConfiguration
    {
        #region "Variables locales"
        private Debug _Debug = Debug.Instance;
        
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
            _MethodBase Methode = System.Reflection.MethodBase.GetCurrentMethod();

            if (!((Config == null) && (Modele == null)))
            {
                _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name);

                _swConfiguration = Config;
                _Modele = Modele;

                return true;
            }
            
            _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name + " : Erreur d'initialisation");
            return false;
        }

        #endregion
    }
}
