using SolidWorks.Interop.swconst;
using System;
using System.Runtime.InteropServices;

namespace Framework
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("0BAA9214-6253-45EB-898C-253813F915D3")]
    public interface IeGestOptions
    {
        swDxfFormat_e DxfDwg_Version { get; set; }
        Boolean DxfDwg_PolicesAutoCAD { get; set; }
        Boolean DxfDwg_StylesAutoCAD { get; set; }
        Boolean DxfDwg_SortieEchelle1 { get; set; }
        Boolean DxfDwg_JoindreExtremites { get; set; }
        Double DxfDwg_JoindreExtremitesTolerance { get; set; }
        Boolean DxfDwg_JoindreExtremitesHauteQualite { get; set; }
        Boolean DxfDwg_ExporterSplineEnPolyligne { get; set; }
        swDxfMultisheet_e DxfDwg_ExporterToutesLesFeuilles { get; set; }
        Boolean DxfDwg_ExporterFeuilleDansEspacePapier { get; set; }

        Boolean Pdf_ExporterEnCouleur { get; set; }
        Boolean Pdf_IncorporerLesPolices { get; set; }
        Boolean Pdf_ExporterEnHauteQualite { get; set; }
        Boolean Pdf_ImprimerEnTeteEtPiedDePage { get; set; }
        Boolean Pdf_UtiliserLesEpaisseursDeLigneDeImprimante { get; set; }

        void ActiverEffetsVisuels();
        void DesactiverEffetsVisuels();

    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("8EF59A57-49BF-4475-9675-BFB4E48D586E")]
    [ProgId("Frameworks.eGestOptions")]
    public class eGestOptions : IeGestOptions
    {
        #region "Variables locales"

        private static readonly String cNOMCLASSE = typeof(eGestOptions).Name;

        private Boolean _EstInitialise = false;

        private eSldWorks _SW = null;

        private Boolean _TransitionDesactive = false;
        private Double _TransitionVue = 0;
        private Double _TransitionAfficherCacher = 0;
        private Double _TransitionIsoler = 0;
        private Double _TransitionSelecteur = 0;
        private Double _AnimationContrainte = 0;

        #endregion

        #region "Constructeur\Destructeur"

        public eGestOptions() { }

        #endregion

        #region "Propriétés"

        public swDxfFormat_e DxfDwg_Version
        {
            get
            {
                Log.Methode(cNOMCLASSE);
                return (swDxfFormat_e)_SW.SwSW.GetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swDxfVersion);
            }
            set
            {
                Log.Methode(cNOMCLASSE);
                _SW.SwSW.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swDxfVersion, (int)value);
            }
        }

        public Boolean DxfDwg_PolicesAutoCAD
        {
            get
            {
                Log.Methode(cNOMCLASSE);
                return !Convert.ToBoolean(_SW.SwSW.GetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swDxfOutputFonts));
            }
            set
            {
                Log.Methode(cNOMCLASSE);
                _SW.SwSW.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swDxfOutputFonts, Convert.ToInt32(!value));
            }
        }

        public Boolean DxfDwg_StylesAutoCAD
        {
            get
            {
                Log.Methode(cNOMCLASSE);
                return !Convert.ToBoolean(_SW.SwSW.GetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swDxfOutputLineStyles));
            }
            set
            {
                Log.Methode(cNOMCLASSE);
                _SW.SwSW.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swDxfOutputLineStyles, Convert.ToInt32(!value));
            }
        }

        public Boolean DxfDwg_SortieEchelle1
        {
            get
            {
                Log.Methode(cNOMCLASSE);
                return Convert.ToBoolean(_SW.SwSW.GetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swDxfOutputNoScale));
            }
            set
            {
                Log.Methode(cNOMCLASSE);
                _SW.SwSW.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swDxfOutputNoScale, Convert.ToInt32(value));
            }
        }

        public Boolean DxfDwg_JoindreExtremites
        {
            get
            {
                Log.Methode(cNOMCLASSE);
                return Convert.ToBoolean(_SW.SwSW.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swDxfEndPointMerge));
            }
            set
            {
                Log.Methode(cNOMCLASSE);
                _SW.SwSW.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swDxfEndPointMerge, value);
            }
        }

        public Double DxfDwg_JoindreExtremitesTolerance
        {
            get
            {
                Log.Methode(cNOMCLASSE);
                return _SW.SwSW.GetUserPreferenceDoubleValue((int)swUserPreferenceDoubleValue_e.swDxfMergingDistance);
            }
            set
            {
                Log.Methode(cNOMCLASSE);
                _SW.SwSW.SetUserPreferenceDoubleValue((int)swUserPreferenceDoubleValue_e.swDxfMergingDistance, value);
            }
        }

        public Boolean DxfDwg_JoindreExtremitesHauteQualite
        {
            get
            {
                Log.Methode(cNOMCLASSE);
                return Convert.ToBoolean(_SW.SwSW.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swDXFHighQualityExport));
            }
            set
            {
                Log.Methode(cNOMCLASSE);
                _SW.SwSW.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swDXFHighQualityExport, value);
            }
        }

        public Boolean DxfDwg_ExporterSplineEnPolyligne
        {
            get
            {
                Log.Methode(cNOMCLASSE);
                return !Convert.ToBoolean(_SW.SwSW.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swDxfExportSplinesAsSplines));
            }
            set
            {
                Log.Methode(cNOMCLASSE);
                Log.Message("--------------------> " + !value);
                _SW.SwSW.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swDxfExportSplinesAsSplines, !value);
            }
        }

        public swDxfMultisheet_e DxfDwg_ExporterToutesLesFeuilles
        {
            get
            {
                Log.Methode(cNOMCLASSE);
                return (swDxfMultisheet_e)_SW.SwSW.GetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swDxfMultiSheetOption);
            }
            set
            {
                Log.Methode(cNOMCLASSE);
                _SW.SwSW.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swDxfMultiSheetOption, (int)value);
            }
        }

        public Boolean DxfDwg_ExporterFeuilleDansEspacePapier
        {
            get
            {
                Log.Methode(cNOMCLASSE);
                return Convert.ToBoolean(_SW.SwSW.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swDxfExportAllSheetsToPaperSpace));
            }
            set
            {
                Log.Methode(cNOMCLASSE);
                _SW.SwSW.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swDxfExportAllSheetsToPaperSpace, value);
            }
        }

        public Boolean Pdf_ExporterEnCouleur
        {
            get
            {
                Log.Methode(cNOMCLASSE);
                return Convert.ToBoolean(_SW.SwSW.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swPDFExportInColor));
            }
            set
            {
                Log.Methode(cNOMCLASSE);
                _SW.SwSW.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swPDFExportInColor, value);
            }
        }

        public Boolean Pdf_IncorporerLesPolices
        {
            get
            {
                Log.Methode(cNOMCLASSE);
                return Convert.ToBoolean(_SW.SwSW.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swPDFExportEmbedFonts));
            }
            set
            {
                Log.Methode(cNOMCLASSE);
                _SW.SwSW.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swPDFExportEmbedFonts, value);
            }
        }

        public Boolean Pdf_ExporterEnHauteQualite
        {
            get
            {
                Log.Methode(cNOMCLASSE);
                return Convert.ToBoolean(_SW.SwSW.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swPDFExportHighQuality));
            }
            set
            {
                Log.Methode(cNOMCLASSE);
                _SW.SwSW.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swPDFExportHighQuality, value);
            }
        }

        public Boolean Pdf_ImprimerEnTeteEtPiedDePage
        {
            get
            {
                Log.Methode(cNOMCLASSE);
                return Convert.ToBoolean(_SW.SwSW.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swPDFExportPrintHeaderFooter));
            }
            set
            {
                Log.Methode(cNOMCLASSE);
                _SW.SwSW.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swPDFExportPrintHeaderFooter, value);
            }
        }

        public Boolean Pdf_UtiliserLesEpaisseursDeLigneDeImprimante
        {
            get
            {
                Log.Methode(cNOMCLASSE);
                return Convert.ToBoolean(_SW.SwSW.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swPDFExportUseCurrentPrintLineWeights));
            }
            set
            {
                Log.Methode(cNOMCLASSE);
                _SW.SwSW.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swPDFExportUseCurrentPrintLineWeights, value);
            }
        }

        /// <summary>
        /// Fonction interne
        /// Test l'initialisation de l'objet GestOptions
        /// </summary>
        internal Boolean EstInitialise { get { Log.Methode(cNOMCLASSE); return _EstInitialise; } }

        #endregion

        #region "Méthodes"

        /// <summary>
        /// Méthode interne
        /// Initialise l'objet GestOptions
        /// </summary>
        /// <param name="SwGestionnaire"></param>
        /// <param name="Modele"></param>
        /// <returns></returns>
        internal Boolean Init(eSldWorks SW)
        {
            Log.Methode(cNOMCLASSE);

            if ((SW != null) && SW.EstInitialise)
            {
                _SW = SW;
                _EstInitialise = true;
            }
            else
            {
                Log.Message("!!!!! Erreur d'initialisation");
            }

            return _EstInitialise;
        }

        public void ActiverEffetsVisuels()
        {
            if (_TransitionDesactive)
            {
                _SW.SwSW.SetUserPreferenceDoubleValue((int)swUserPreferenceDoubleValue_e.swViewAnimationSpeed, _TransitionVue);
                _SW.SwSW.SetUserPreferenceDoubleValue((int)swUserPreferenceDoubleValue_e.swViewTransitionHideShowComponent, _TransitionAfficherCacher);
                _SW.SwSW.SetUserPreferenceDoubleValue((int)swUserPreferenceDoubleValue_e.swViewTransitionIsolate, _TransitionIsoler);
                _SW.SwSW.SetUserPreferenceDoubleValue((int)swUserPreferenceDoubleValue_e.swViewSelectorSpeed, _TransitionSelecteur);
                _SW.SwSW.SetUserPreferenceDoubleValue((int)swUserPreferenceDoubleValue_e.swMateAnimationSpeed, _AnimationContrainte);
            }

            _TransitionDesactive = false;
        }

        public void DesactiverEffetsVisuels()
        {
            _TransitionVue = _SW.SwSW.GetUserPreferenceDoubleValue((int)swUserPreferenceDoubleValue_e.swViewAnimationSpeed);
            _TransitionAfficherCacher = _SW.SwSW.GetUserPreferenceDoubleValue((int)swUserPreferenceDoubleValue_e.swViewTransitionHideShowComponent);
            _TransitionIsoler = _SW.SwSW.GetUserPreferenceDoubleValue((int)swUserPreferenceDoubleValue_e.swViewTransitionIsolate);
            _TransitionSelecteur = _SW.SwSW.GetUserPreferenceDoubleValue((int)swUserPreferenceDoubleValue_e.swViewSelectorSpeed);
            _AnimationContrainte = _SW.SwSW.GetUserPreferenceDoubleValue((int)swUserPreferenceDoubleValue_e.swMateAnimationSpeed);

            _SW.SwSW.SetUserPreferenceDoubleValue((int)swUserPreferenceDoubleValue_e.swViewAnimationSpeed, 0);
            _SW.SwSW.SetUserPreferenceDoubleValue((int)swUserPreferenceDoubleValue_e.swViewTransitionHideShowComponent, 0);
            _SW.SwSW.SetUserPreferenceDoubleValue((int)swUserPreferenceDoubleValue_e.swViewTransitionIsolate, 0);
            _SW.SwSW.SetUserPreferenceDoubleValue((int)swUserPreferenceDoubleValue_e.swViewSelectorSpeed, 0);
            _SW.SwSW.SetUserPreferenceDoubleValue((int)swUserPreferenceDoubleValue_e.swMateAnimationSpeed, 0);

            _TransitionDesactive = true;
        }

        #endregion
    }
}
