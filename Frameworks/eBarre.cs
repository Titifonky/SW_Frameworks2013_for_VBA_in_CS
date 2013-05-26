using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using System.Collections;
using System.Reflection;

namespace Framework_SW2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("75D0F041-D524-454D-A6E8-7337B9588633")]
    public interface IeBarre
    {
        eCorps Corps { get; }
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("C974BBDC-3C5A-495F-A532-2A61F7509025")]
    [ProgId("Frameworks.eBarre")]
    public class eBarre : IeBarre
    {
        #region "Variables locales"

        private Boolean _EstInitialise = false;

        private eCorps _Corps = null;

        #endregion

        #region "Constructeur\Destructeur"

        public eBarre() { }

        #endregion

        #region "Propriétés"

        /// <summary>
        /// Retourne le parent ExtPiece.
        /// </summary>
        public eCorps Corps { get { Debug.Info(MethodBase.GetCurrentMethod()); return _Corps; } }

        /// <summary>
        /// Fonction interne.
        /// Test l'initialisation de l'objet ExtCorps.
        /// </summary>
        internal Boolean EstInitialise { get { Debug.Info(MethodBase.GetCurrentMethod()); return _EstInitialise; } }

        #endregion

        #region "Méthodes"

        /// <summary>
        /// Méthode interne.
        /// Initialiser l'objet eBarre.
        /// </summary>
        /// <param name="SwCorps"></param>
        /// <param name="Corps"></param>
        /// <returns></returns>
        internal Boolean Init(eCorps Corps)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            if ((Corps != null) && Corps.EstInitialise && (Corps.TypeDeCorps == TypeCorps_e.cBarre))
            {
                _Corps = Corps;
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
