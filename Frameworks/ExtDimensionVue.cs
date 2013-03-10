using System;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using System.Reflection;

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
        
        private Boolean _EstInitialise = false;

        private ExtVue _Vue;
        #endregion

        #region "Constructeur\Destructeur"

        public ExtDimensionVue() { }

        #endregion

        #region "Propriétés"

        public ExtVue Vue { get { Debug.Info(MethodBase.GetCurrentMethod());  return _Vue; } }

        public Point Centre
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                
                Point pCentre;
                Double[] pArrayResult;
                pArrayResult = _Vue.SwVue.Position;

                pCentre.X = pArrayResult[0];
                pCentre.Y = pArrayResult[1];
                pCentre.Z = 0;

                return pCentre;
            }
            set
            {
                Debug.Info(MethodBase.GetCurrentMethod());

                Double[] pCentre = { value.X, value.Y };
                _Vue.SwVue.Position = pCentre;
            }
        }

        public Dimensions Dimensions
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
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
                Debug.Info(MethodBase.GetCurrentMethod());
                Coordonnees pCoord;
                Object[] pArr = _Vue.SwVue.GetOutline();

                pCoord.MinX = (Double)pArr[0];
                pCoord.MinY = (Double)pArr[1];
                pCoord.MaxX = (Double)pArr[2];
                pCoord.MaxY = (Double)pArr[3];

                return pCoord;
            }
            }

        internal Boolean EstInitialise { get { Debug.Info(MethodBase.GetCurrentMethod());  return _EstInitialise; } }

        #endregion

        #region "Méthodes"

        internal Boolean Init(ExtVue Vue)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            if ((Vue != null) && Vue.EstInitialise)
            {
                _Vue = Vue;

                _EstInitialise = true;
            }
            else
            {
                Debug.Info("!!!!! Erreur d'initialisation");
            }
            return _EstInitialise;
        }

        #endregion

    }
}
