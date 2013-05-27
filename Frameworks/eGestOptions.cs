using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Reflection;

namespace Framework_SW2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("0BAA9214-6253-45EB-898C-253813F915D3")]
    public interface IeGestOptions
    {
        swDxfFormat_e DxfDwg_Format { get; set; }
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
        Boolean Pdf_utiliserLesEpaisseursDeLigneDeImprimante { get; set; }
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("8EF59A57-49BF-4475-9675-BFB4E48D586E")]
    [ProgId("Frameworks.eGestOptions")]
    public class eGestOptions : IeGestOptions
    {
        #region "Variables locales"

        private Boolean _EstInitialise = false;

        private eSldWorks _SW = null;

        #endregion

        #region "Constructeur\Destructeur"

        public eGestOptions() { }

        #endregion

        #region "Propriétés"

        public swDxfFormat_e DxfDwg_Format
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                return (swDxfFormat_e)_SW.SwSW.GetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swDxfVersion);
            }
            set
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                _SW.SwSW.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swDxfVersion, (int)value);
            }
        }

        public Boolean DxfDwg_PolicesAutoCAD
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                return !Convert.ToBoolean(_SW.SwSW.GetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swDxfOutputFonts));
            }
            set
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                _SW.SwSW.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swDxfOutputFonts, Convert.ToInt32(!value));
            }
        }

        public Boolean DxfDwg_StylesAutoCAD
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                return !Convert.ToBoolean(_SW.SwSW.GetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swDxfOutputLineStyles));
            }
            set
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                _SW.SwSW.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swDxfOutputLineStyles, Convert.ToInt32(!value));
            }
        }

        public Boolean DxfDwg_SortieEchelle1
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                return Convert.ToBoolean(_SW.SwSW.GetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swDxfOutputNoScale));
            }
            set
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                _SW.SwSW.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swDxfOutputNoScale, Convert.ToInt32(value));
            }
        }

        public Boolean DxfDwg_JoindreExtremites
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                return Convert.ToBoolean(_SW.SwSW.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swDxfEndPointMerge));
            }
            set
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                _SW.SwSW.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swDxfEndPointMerge, value);
            }
        }

        public Double DxfDwg_JoindreExtremitesTolerance
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                return _SW.SwSW.GetUserPreferenceDoubleValue((int)swUserPreferenceDoubleValue_e.swDxfMergingDistance);
            }
            set
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                _SW.SwSW.SetUserPreferenceDoubleValue((int)swUserPreferenceDoubleValue_e.swDxfMergingDistance, value);
            }
        }

        public Boolean DxfDwg_JoindreExtremitesHauteQualite
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                return Convert.ToBoolean(_SW.SwSW.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swDXFHighQualityExport));
            }
            set
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                _SW.SwSW.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swDXFHighQualityExport, value);
            }
        }

        public Boolean DxfDwg_ExporterSplineEnPolyligne
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                return !Convert.ToBoolean(_SW.SwSW.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swDxfExportSplinesAsSplines));
            }
            set
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                _SW.SwSW.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swDxfExportSplinesAsSplines, !value);
            }
        }

        public swDxfMultisheet_e DxfDwg_ExporterToutesLesFeuilles
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                return (swDxfMultisheet_e)_SW.SwSW.GetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swDxfMultiSheetOption);
            }
            set
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                _SW.SwSW.SetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swDxfMultiSheetOption, (int)value);
            }
        }

        public Boolean DxfDwg_ExporterFeuilleDansEspacePapier
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                return Convert.ToBoolean(_SW.SwSW.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swDxfExportAllSheetsToPaperSpace));
            }
            set
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                _SW.SwSW.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swDxfExportAllSheetsToPaperSpace, value);
            }
        }

        public Boolean Pdf_ExporterEnCouleur
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                return Convert.ToBoolean(_SW.SwSW.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swPDFExportInColor));
            }
            set
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                _SW.SwSW.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swPDFExportInColor, value);
            }
        }

        public Boolean Pdf_IncorporerLesPolices
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                return Convert.ToBoolean(_SW.SwSW.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swPDFExportEmbedFonts));
            }
            set
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                _SW.SwSW.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swPDFExportEmbedFonts, value);
            }
        }
        
        public Boolean Pdf_ExporterEnHauteQualite
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                return Convert.ToBoolean(_SW.SwSW.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swPDFExportHighQuality));
            }
            set
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                _SW.SwSW.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swPDFExportHighQuality, value);
            }
        }
        
        public Boolean Pdf_ImprimerEnTeteEtPiedDePage
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                return Convert.ToBoolean(_SW.SwSW.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swPDFExportPrintHeaderFooter));
            }
            set
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                _SW.SwSW.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swPDFExportPrintHeaderFooter, value);
            }
        }
        
        public Boolean Pdf_utiliserLesEpaisseursDeLigneDeImprimante
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                return Convert.ToBoolean(_SW.SwSW.GetUserPreferenceToggle((int)swUserPreferenceToggle_e.swPDFExportUseCurrentPrintLineWeights));
            }
            set
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                _SW.SwSW.SetUserPreferenceToggle((int)swUserPreferenceToggle_e.swPDFExportUseCurrentPrintLineWeights, value);
            }
        }

        /// <summary>
        /// Fonction interne
        /// Test l'initialisation de l'objet GestOptions
        /// </summary>
        internal Boolean EstInitialise { get { Debug.Info(MethodBase.GetCurrentMethod()); return _EstInitialise; } }

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
            Debug.Info(MethodBase.GetCurrentMethod());

            if ((SW != null) && SW.EstInitialise)
            {
                _SW = SW;
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
