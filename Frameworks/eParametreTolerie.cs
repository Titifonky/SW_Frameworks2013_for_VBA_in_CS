using System;
using System.Reflection;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace Framework_SW2013
{
#if SW2013
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("3E082AF3-8F22-4687-A8AB-5DBE544FA5D8")]
    public interface IeParametreTolerie
    {
        ePiece Piece { get; }
        eCorps Corps { get; }
        Double Epaisseur { get; set; }
        Double Rayon { get; set; }
        Double FacteurK { get; set; }
        Boolean EcraserEpaisseur { get; set; }
        Boolean EcraserRayon { get; set; }
        Boolean EcraserFacteurK { get; set; }
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("F699797C-AD86-4E2C-957E-2BE8EE033941")]
    [ProgId("Frameworks.eParametreTolerie")]
    public class eParametreTolerie : IeParametreTolerie
    {
        #region "Variables locales"

        private Boolean _EstInitialise = false;

        private ePiece _Piece = null;
        private eCorps _Corps = null;

        #endregion

        #region "Constructeur\Destructeur"

        public eParametreTolerie() { }

        #endregion

        #region "Propriétés"

        /// <summary>
        /// Retourne le parent ExtPiece.
        /// </summary>
        public ePiece Piece { get { Debug.Print(MethodBase.GetCurrentMethod()); return _Piece; } }

        /// <summary>
        /// Retourne le parent ExtPiece.
        /// </summary>
        public eCorps Corps { get { Debug.Print(MethodBase.GetCurrentMethod()); return _Corps; } }

        /// <summary>
        /// Retourne ou défini l'épaisseur de la tole
        /// </summary>
        public Double Epaisseur
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                Dimension pSwDim = FonctionTolerie.SwFonction.Parameter("D7");
                return Math.Round(pSwDim.GetSystemValue3((int)swInConfigurationOpts_e.swThisConfiguration, null)[0] * 1000, 5);
            }
            set
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                Dimension pSwDim = FonctionTolerie.SwFonction.Parameter("D7");
                if (pSwDim == null) return;
                pSwDim.SetSystemValue3(value * 0.001, (int)swInConfigurationOpts_e.swThisConfiguration, null);
            }
        }

        /// <summary>
        /// Retourne ou défini le rayon intérieur de pliage
        /// </summary>
        public Double Rayon
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                Dimension pSwDim = FonctionTolerie.SwFonction.Parameter("D1");
                return Math.Round(pSwDim.GetSystemValue3((int)swInConfigurationOpts_e.swThisConfiguration, null)[0] * 1000, 5);
            }
            set
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                Dimension pSwDim = FonctionTolerie.SwFonction.Parameter("D1");
                if (pSwDim == null) return;
                Debug.Print("Modification du rayon : " + pSwDim.SetSystemValue3(value * 0.001, (int)swInConfigurationOpts_e.swThisConfiguration, null).ToString());
            }
        }

        /// <summary>
        /// Retourne ou défini le facteur K pour le developpé
        /// </summary>
        public Double FacteurK
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                Dimension pSwDim = FonctionTolerie.SwFonction.Parameter("D2");
                return Math.Round(pSwDim.GetSystemValue3((int)swInConfigurationOpts_e.swThisConfiguration, null)[0], 5);
            }
            set
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                Dimension pSwDim = FonctionTolerie.SwFonction.Parameter("D2");
                if (pSwDim == null) return;
                pSwDim.SetSystemValue3(value, (int)swInConfigurationOpts_e.swThisConfiguration, null);
            }
        }

        public Boolean EcraserEpaisseur
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                if (_Corps == null)
                    return false;
                BaseFlangeFeatureData pParam = FonctionToleDeBase.SwFonction.GetDefinition();
                return pParam.OverrideThickness;
            }
            set
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                if (_Corps == null)
                    return;
                ModelDoc2 pSwModele = SwModele;
                Component2 pSwComposant = SwComposant;
                Feature pSwFonction = FonctionToleDeBase.SwFonction;
                BaseFlangeFeatureData pParam = pSwFonction.GetDefinition();
                pParam.AccessSelections(pSwModele, pSwComposant);
                pParam.OverrideThickness = value;
                pSwFonction.ModifyDefinition(pParam, pSwModele, pSwComposant);
                pParam.ReleaseSelectionAccess();
            }
        }

        public Boolean EcraserRayon
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                if (_Corps == null)
                    return false;
                BaseFlangeFeatureData pParam = FonctionToleDeBase.SwFonction.GetDefinition();
                return pParam.OverrideRadius;
            }
            set
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                if (_Corps == null)
                    return;
                ModelDoc2 pSwModele = SwModele;
                Component2 pSwComposant = SwComposant;
                Feature pSwFonction = FonctionToleDeBase.SwFonction;
                BaseFlangeFeatureData pParam = pSwFonction.GetDefinition();
                pParam.AccessSelections(pSwModele, pSwComposant);
                pParam.OverrideRadius = value;
                pSwFonction.ModifyDefinition(pParam, pSwModele, pSwComposant);
                pParam.ReleaseSelectionAccess();
            }
        }

        public Boolean EcraserFacteurK
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                if (_Corps == null)
                    return false;
                BaseFlangeFeatureData pParam = FonctionToleDeBase.SwFonction.GetDefinition();
                return pParam.OverrideKFactor;
            }
            set
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                if (_Corps == null)
                    return;
                ModelDoc2 pSwModele = SwModele;
                Component2 pSwComposant = SwComposant;
                Feature pSwFonction = FonctionToleDeBase.SwFonction;
                BaseFlangeFeatureData pParam = pSwFonction.GetDefinition();
                pParam.AccessSelections(pSwModele, pSwComposant);
                pParam.OverrideKFactor = value;
                pSwFonction.ModifyDefinition(pParam, pSwModele, pSwComposant);
                pParam.ReleaseSelectionAccess();
            }
        }

        /// <summary>
        /// Renvoi la fonction Tolerie
        /// </summary>
        /// <returns></returns>
        private eFonction FonctionTolerie
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());

                if (_Piece != null)
                    return _Piece.Modele.ListListeDesFonctions("^Tôlerie$", "", false, true)[0];
                else
                    return _Corps.Tole.FonctionTolerie;
            }
        }

        /// <summary>
        /// Renvoi la fonction Tolerie
        /// </summary>
        /// <returns></returns>
        private eFonction FonctionToleDeBase
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());

                if (_Corps != null)
                    return _Corps.Tole.FonctionToleDeBase;

                return null;
            }
        }

        private ModelDoc2 SwModele
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                if (_Corps.Piece.Modele.Equals(_Corps.Piece.Modele.SW.Modele()))
                    return _Corps.Piece.Modele.SwModele;

                return null;
            }
        }

        private Component2 SwComposant
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                if (SwModele == null)
                    return _Corps.Piece.Modele.Composant.SwComposant;

                return null;
            }
        }

        /// <summary>
        /// Fonction interne.
        /// Test l'initialisation de l'objet eTole.
        /// </summary>
        internal Boolean EstInitialise { get { Debug.Print(MethodBase.GetCurrentMethod()); return _EstInitialise; } }

        #endregion

        #region "Méthodes"

        /// <summary>
        /// Méthode interne.
        /// Initialiser
        /// </summary>
        /// <param name="SwCorps"></param>
        /// <param name="Corps"></param>
        /// <returns></returns>
        internal Boolean Init(eCorps Corps)
        {
            Debug.Print(MethodBase.GetCurrentMethod());

            if ((Corps != null) && Corps.EstInitialise && (Corps.TypeDeCorps == TypeCorps_e.cTole))
            {
                _Corps = Corps;
                _EstInitialise = true;
            }
            else
            {
                Debug.Print("!!!!! Erreur d'initialisation");
            }
            return _EstInitialise;
        }

        /// <summary>
        /// Méthode interne.
        /// Initialiser
        /// </summary>
        /// <param name="SwCorps"></param>
        /// <param name="Piece"></param>
        /// <returns></returns>
        internal Boolean Init(ePiece Piece)
        {
            Debug.Print(MethodBase.GetCurrentMethod());

            if ((Piece != null) && Piece.EstInitialise && (Piece.Contient(TypeCorps_e.cTole)))
            {
                _Piece = Piece;
                _EstInitialise = true;
            }
            else
            {
                Debug.Print("!!!!! Erreur d'initialisation");
            }
            return _EstInitialise;
        }

        #endregion
    }

#endif
}
