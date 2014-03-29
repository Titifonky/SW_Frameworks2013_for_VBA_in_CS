using System;
using System.Runtime.InteropServices;

namespace Framework
{

    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("685D0EE5-1B10-41C6-9DCE-B8C9D4B85B84")]
    public interface IeDimensionVue
    {
        eVue Vue { get; }
        ePoint Centre { get; set; }
        eRectangle Rectangle { get; }
        eZone Zone { get; }
        Double Angle { get; set; }
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("A6A0EE53-8B0A-4BAC-8AE6-9D9BD029391F")]
    [ProgId("Frameworks.eDimensionVue")]
    public class eDimensionVue : IeDimensionVue
    {
        #region "Variables locales"

        private static readonly String cNOMCLASSE = typeof(eDimensionVue).Name;

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
        public eVue Vue { get { Log.Methode(cNOMCLASSE); return _Vue; } }

        /// <summary>
        /// Retourne ou défini le centre de la vue.
        /// </summary>
        public ePoint Centre
        {
            get
            {
                Log.Methode(cNOMCLASSE);

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
                Log.Methode(cNOMCLASSE);

                Double[] pCentre = { value.X, value.Y };
                _Vue.SwVue.Position = pCentre;
            }
        }

        /// <summary>
        /// Retourne les dimensions de la vue, hauteur et largeur.
        /// </summary>
        public eRectangle Rectangle
        {
            get
            {
                Log.Methode(cNOMCLASSE);
                eRectangle pDim = new eRectangle();
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
                Log.Methode(cNOMCLASSE);
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
            get
            {
                Log.Methode(cNOMCLASSE);
                return _Vue.SwVue.Angle * 180.0 / Math.PI;
            }
            set
            {
                Log.Methode(cNOMCLASSE);
                _Vue.SwVue.Angle = value * Math.PI / 180.0;
            }
        }

        /// <summary>
        /// Fonction interne.
        /// Test l'initialisation de l'objet ExtDimensionVue.
        /// </summary>
        internal Boolean EstInitialise { get { Log.Methode(cNOMCLASSE); return _EstInitialise; } }

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
            Log.Methode(cNOMCLASSE);

            if ((Vue != null) && Vue.EstInitialise)
            {
                _Vue = Vue;

                _EstInitialise = true;
            }
            else
            {
                Log.Message("!!!!! Erreur d'initialisation");
            }
            return _EstInitialise;
        }

        #endregion

    }
}
