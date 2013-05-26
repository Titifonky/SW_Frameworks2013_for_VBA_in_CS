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
    [Guid("44D349FF-E4E1-4C9A-86EB-197EC4AE0DBE")]
    public interface IeTole
    {
        eCorps Corps { get; }
        eFonction FonctionTolerie();
        eFonction FonctionDeplie();
        eFonction FonctionCubeDeVisualisation();
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("DCE24670-A986-4606-88FE-577980D34685")]
    [ProgId("Frameworks.eTole")]
    public class eTole : IeTole
    {
        #region "Variables locales"
        
        private Boolean _EstInitialise = false;

        private eCorps _Corps = null;

        #endregion

        #region "Constructeur\Destructeur"

        public eTole() { }

        #endregion

        #region "Propriétés"

        /// <summary>
        /// Retourne le parent ExtPiece.
        /// </summary>
        public eCorps Corps { get { Debug.Info(MethodBase.GetCurrentMethod());  return _Corps; } }

        /// <summary>
        /// Fonction interne.
        /// Test l'initialisation de l'objet eTole.
        /// </summary>
        internal Boolean EstInitialise { get { Debug.Info(MethodBase.GetCurrentMethod());  return _EstInitialise; } }

        #endregion

        #region "Méthodes"

        /// <summary>
        /// Méthode interne.
        /// Initialiser l'objet ExtCorps.
        /// </summary>
        /// <param name="SwCorps"></param>
        /// <param name="Corps"></param>
        /// <returns></returns>
        internal Boolean Init(eCorps Corps)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            if ((Corps != null) && Corps.EstInitialise && (Corps.TypeDeCorps == TypeCorps_e.cTole))
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

        /// <summary>
        /// Renvoi la fonction Tolerie du corps
        /// </summary>
        /// <returns></returns>
        public eFonction FonctionTolerie()
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            swFeatureType_e TypeFonc = new swFeatureType_e();

            foreach (eFonction pFonc in _Corps.ListListeDesFonctions())
            {
                if (pFonc.TypeDeLaFonction == TypeFonc.swTnSheetMetal)
                    return pFonc;
            }

            return null;
        }

        /// <summary>
        /// Renvoi la fonction EtatDeplie du corps
        /// </summary>
        /// <returns></returns>
        public eFonction FonctionDeplie()
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            swFeatureType_e TypeFonc = new swFeatureType_e();

            foreach (eFonction pFonc in _Corps.ListListeDesFonctions())
            {
                if (pFonc.TypeDeLaFonction == TypeFonc.swTnFlatPattern)
                    return pFonc;
            }

            return null;
        }

        /// <summary>
        /// Renvoi la fonction CubeDeVisualisation du corps
        /// </summary>
        /// <returns></returns>
        public eFonction FonctionCubeDeVisualisation()
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            return this.FonctionDeplie().ListListeDesSousFonctions(CONSTANTES.CUBE_DE_VISUALISATION)[0];
        }

        #endregion
    }
}
