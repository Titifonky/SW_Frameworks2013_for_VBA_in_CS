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
        Double Epaisseur { get; set; }
        Double Rayon { get; set; }
        Double FacteurK { get; set; }
        eFonction FonctionTolerie { get; }
        eFonction FonctionToleDeBase { get; }
        eFonction FonctionDeplie { get; }
        eFonction FonctionCubeDeVisualisation { get; }
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

        private ModelDoc2 SwModele
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                if(_Corps.Piece.Modele.Equals(_Corps.Piece.Modele.SW.Modele()))
                    return _Corps.Piece.Modele.SwModele;

                return null;
            }
        }

        private Component2 SwComposant
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                if (SwModele == null)
                    return _Corps.Piece.Modele.Composant.SwComposant;

                return null;
            }
        }

        private SheetMetalFeatureData SwParametreTolerie
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                return FonctionTolerie.SwFonction.GetDefinition();
            }
        }

        /// <summary>
        /// Retourne ou défini l'épaisseur de la tole
        /// </summary>
        public Double Epaisseur
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                return SwParametreTolerie.Thickness * 1000;
            }
            set
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                ModelDoc2 pSwModele = SwModele;
                Component2 pSwComposant = SwComposant;

#if SW2013
                SheetMetalFeatureData pParam = SwParametreTolerie;
                Feature pSwFonctionTolerie = FonctionTolerie.SwFonction;
#else
                Feature pSwFonction = FonctionToleDeBase.SwFonction;
                BaseFlangeFeatureData pParam = pSwFonction.GetDefinition();
#endif
                pParam.AccessSelections(pSwModele, pSwComposant);
                pParam.Thickness = value * 0.001;
                pSwFonction.ModifyDefinition(pParam, pSwModele, pSwComposant);
                pParam.ReleaseSelectionAccess();
            }
        }

        /// <summary>
        /// Retourne ou défini le rayon intérieur de pliage
        /// </summary>
        public Double Rayon
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                return SwParametreTolerie.BendRadius * 1000;
            }
            set
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                ModelDoc2 pSwModele = SwModele;
                Component2 pSwComposant = SwComposant;
                SheetMetalFeatureData pParam = SwParametreTolerie;
                Feature pSwFonctionTolerie = FonctionTolerie.SwFonction;
                pParam.AccessSelections(pSwModele, pSwComposant);
                pParam.BendRadius = value * 0.001;
                pSwFonctionTolerie.ModifyDefinition(pParam, pSwModele, pSwComposant);
                pParam.ReleaseSelectionAccess();
            }
        }

        /// <summary>
        /// Retourne ou défini le facteur K pour le developpé
        /// </summary>
        public Double FacteurK
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                return SwParametreTolerie.GetCustomBendAllowance().KFactor;
            }
            set
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                ModelDoc2 pSwModele = SwModele;
                Component2 pSwComposant = SwComposant;
                SheetMetalFeatureData pParam = SwParametreTolerie;
                Feature pSwFonctionTolerie = FonctionTolerie.SwFonction;
                CustomBendAllowance pParamPli = pParam.GetCustomBendAllowance();
                pParam.AccessSelections(pSwModele, pSwComposant);
                pParamPli.KFactor = value;
                pParam.SetCustomBendAllowance(pParamPli);
                pSwFonctionTolerie.ModifyDefinition(pParam, pSwModele, pSwComposant);
                pParam.ReleaseSelectionAccess();
            }
        }

        /// <summary>
        /// Renvoi la fonction Tolerie du corps
        /// </summary>
        /// <returns></returns>
        public eFonction FonctionTolerie
        {
            get
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
        }

        /// <summary>
        /// Renvoi la fonction Tolerie du corps
        /// </summary>
        /// <returns></returns>
        public eFonction FonctionToleDeBase
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());

                swFeatureType_e TypeFonc = new swFeatureType_e();

                foreach (eFonction pFonc in _Corps.ListListeDesFonctions())
                {
                    if ((pFonc.TypeDeLaFonction == TypeFonc.swTnBaseFlange) || (pFonc.TypeDeLaFonction == TypeFonc.swTnSolidToSheetMetal))
                        return pFonc;

                }

                return null;
            }
        }

        /// <summary>
        /// Renvoi la fonction EtatDeplie du corps
        /// </summary>
        /// <returns></returns>
        public eFonction FonctionDeplie
        {
            get
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
        }

        /// <summary>
        /// Renvoi la fonction CubeDeVisualisation du corps
        /// </summary>
        /// <returns></returns>
        public eFonction FonctionCubeDeVisualisation
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());

                return this.FonctionDeplie.ListListeDesSousFonctions(CONSTANTES.CUBE_DE_VISUALISATION)[0];
            }
        }

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

        #endregion
    }
}
