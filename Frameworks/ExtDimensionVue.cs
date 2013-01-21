using System;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;

namespace Framework_SW2013
{

    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("685D0EE5-1B10-41C6-9DCE-B8C9D4B85B84")]
    public interface IExtDimensionVue
    {
        ExtVue Vue { get; }
        Point Centre { get; set; }
        Dimensions Dimensions { get; }
        Coordonnees Coordonnees { get; }
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("A6A0EE53-8B0A-4BAC-8AE6-9D9BD029391F")]
    [ProgId("Frameworks.ExtDimensionVue")]
    public class ExtDimensionVue : IExtDimensionVue
    {
        #region "Variables locales"
        private Debug _Debug = Debug.Instance;
        private Boolean _EstInitialise = false;

        private ExtVue _Vue;
        #endregion

        #region "Constructeur\Destructeur"

        public ExtDimensionVue() { }

        #endregion

        #region "Propriétés"

        public ExtVue Vue { get { return _Vue; } }

        public Point Centre
        {
            get
            {
                Point pCentre;
                Object[] pArrayResult;
                pArrayResult = _Vue.SwVue.Position;

                pCentre.X = (Double)pArrayResult[0];
                pCentre.Y = (Double)pArrayResult[1];
                pCentre.Z = 0;

                return pCentre;
            }
            set
            {
                Object[] pCentre = { value.X, value.Y };
                _Vue.SwVue.Position = pCentre;
            }
        }

        public Dimensions Dimensions
        {
            get
            {
                Dimensions pDim;
                pDim.Lg = Coordonnees.MaxX - Coordonnees.MinX;
                pDim.Ht = Coordonnees.MaxY - Coordonnees.MinY;

                return pDim;
            }
        }

        public Coordonnees Coordonnees
        {
            get
            {
                Coordonnees pCoord;
                Object[] pArr = _Vue.SwVue.GetOutline();

                pCoord.MinX = (Double)pArr[0];
                pCoord.MinY = (Double)pArr[1];
                pCoord.MaxX = (Double)pArr[2];
                pCoord.MaxY = (Double)pArr[3];

                return pCoord;
            }
            }

        internal Boolean EstInitialise { get { return _EstInitialise; } }

        #endregion

        #region "Méthodes"

        internal Boolean Init(ExtVue Vue)
        {
            _MethodBase Methode = System.Reflection.MethodBase.GetCurrentMethod();

            if ((Vue != null) && Vue.EstInitialise)
            {
                _Vue = Vue;

                _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name);
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
