using System;
using System.Reflection;
using System.Runtime.InteropServices;

namespace Framework_SW2013
{

    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("685D0EE5-1B10-41C6-9DCE-B8C9D4B85B84")]
    public interface IeDimensionVue
    {
        eVue Vue { get; }
        ePoint Centre { get; set; }
        eRectangle Dimensions { get; }
        eZone Zone { get; }
        Double Angle { get; set; }
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("A6A0EE53-8B0A-4BAC-8AE6-9D9BD029391F")]
    [ProgId("Frameworks.eDimensionVue")]
    public class eDimensionVue : IeDimensionVue
    {
#region "Variables locales"
        
        private Boolean _EstInitialise = false;

        private eVue _Vue = null;
#endregion

#region "Constructeur\Destructeur"

        public eDimensionVue() { }

#endregion

#region "Propriétés"

        /// <summary>
        /// Retourne le parent ExtVue.
        /// </summary>
        public eVue Vue { get { Debug.Info(MethodBase.GetCurrentMethod());  return _Vue; } }

        /// <summary>
        /// Retourne ou défini le centre de la vue.
        /// </summary>
        public ePoint Centre
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                
                ePoint pCentre = new ePoint();
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

        /// <summary>
        /// Retourne les dimensions de la vue, hauteur et largeur.
        /// </summary>
        public eRectangle Dimensions
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                eRectangle pDim = new eRectangle() ;
                pDim.Lg = Zone.PointMax.X - Zone.PointMin.X;
                pDim.Ht = Zone.PointMax.Y - Zone.PointMin.Y;

                return pDim;
            }
        }

        /// <summary>
        /// Retourne les coordonnées des coins "Bas-Gouche" et "Haut-Droit" de la vue.
        /// </summary>
        public eZone Zone
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                eZone pCoord = new eZone();
                Double[] pArr = _Vue.SwVue.GetOutline();

                pCoord.PointMin.X = pArr[0];
                pCoord.PointMin.Y = pArr[1];
                pCoord.PointMax.X = pArr[2];
                pCoord.PointMax.Y = pArr[3];

                return pCoord;
            }
            }

        /// <summary>
        /// Retourne ou défini l'angle de la vue.
        /// </summary>
        public Double Angle
        {
            get {
                Debug.Info(MethodBase.GetCurrentMethod());
                return _Vue.SwVue.Angle;
            }
            set {
                Debug.Info(MethodBase.GetCurrentMethod());
                _Vue.SwVue.Angle = value;
            }
        }

        /// <summary>
        /// Fonction interne.
        /// Test l'initialisation de l'objet ExtDimensionVue.
        /// </summary>
        internal Boolean EstInitialise { get { Debug.Info(MethodBase.GetCurrentMethod());  return _EstInitialise; } }

#endregion

#region "Méthodes"

        /// <summary>
        /// Méthode interne.
        /// Initialiser l'objet ExtDimensionVue.
        /// </summary>
        /// <param name="Vue"></param>
        /// <returns></returns>
        internal Boolean Init(eVue Vue)
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
