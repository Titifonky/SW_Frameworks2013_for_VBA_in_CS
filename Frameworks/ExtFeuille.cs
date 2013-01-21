using System;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;

namespace Framework_SW2013
{

    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("17F1BCFD-2428-4DF1-8338-8FFA142E2A97")]
    public interface IExtFeuille
    {
        Sheet SwFeuille { get; }
        ExtDessin Dessin { get; }
        String Nom { get; set; }
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("AB11E456-34CF-4540-A7E3-E01D7C63E324")]
    [ProgId("Frameworks.ExtFeuille")]
    public class ExtFeuille : IExtFeuille
    {
        #region "Variables locales"
        private Debug _Debug = Debug.Instance;
        private Boolean _EstInitialise = false;

        private ExtDessin _Dessin;
        private Sheet _SwFeuille;
        #endregion

        #region "Constructeur\Destructeur"

        public ExtFeuille() { }

        #endregion

        #region "Propriétés"

        public Sheet SwFeuille { get { return _SwFeuille; } }

        public ExtDessin Dessin { get { return _Dessin; } }

        public String Nom { get { return _SwFeuille.GetName(); } set { _SwFeuille.SetName(value); } }

        internal Boolean EstInitialise { get { return _EstInitialise; } }

        #endregion

        #region "Méthodes"

        internal Boolean Init(Sheet SwFeuille, ExtDessin Dessin)
        {
            _MethodBase Methode = System.Reflection.MethodBase.GetCurrentMethod();

            if ((SwFeuille != null) && (Dessin != null) && Dessin.EstInitialise)
            {
                _Dessin = Dessin;
                _SwFeuille = SwFeuille;

                _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name + " : " + this.Nom);
                _EstInitialise = true;
            }
            else
            {
                _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name + " : Erreur d'initialisation");
            }
            return _EstInitialise;
        }

        #endregion

    }
}
