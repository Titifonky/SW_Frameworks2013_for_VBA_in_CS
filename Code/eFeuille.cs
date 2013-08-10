﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace Framework
{

    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("17F1BCFD-2428-4DF1-8338-8FFA142E2A97")]
    public interface IeFeuille
    {
        Sheet SwFeuille { get; }
        eDessin Dessin { get; }
        String Nom { get; set; }
        eVue PremiereVue { get; }
        eVue DerniereVue { get; }
        int NbDeVues { get; }
        eZone EnveloppeDesVues { get; }
        Format_e Format { get; set; }
        Orientation_e Orientation { get; set; }
        String GabaritDeFeuille { get; set; }
        void Activer();
        void Supprimer();
        void ZoomEtendu();
        void AjusterAutourDesVues();
        void Redimensionner(Double Largeur, Double Hauteur);
        eVue CreerVueToleDvp(ePiece Piece, eConfiguration Configuration, Boolean AfficherLesLignesDePliage);
        void ExporterEn(Extension_e TypeExport, String CheminDossier, String NomDuFichierAlternatif = "");
        ArrayList ListeDesVues(String NomARechercher = "");
        void MettreEnPagePourImpression(swPageSetupDrawingColor_e Couleur = swPageSetupDrawingColor_e.swPageSetup_AutomaticDrawingColor, Boolean HauteQualite = false);
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("AB11E456-34CF-4540-A7E3-E01D7C63E324")]
    [ProgId("Frameworks.eFeuille")]
    public class eFeuille : IeFeuille
    {
        #region "Variables locales"

        private Boolean _EstInitialise = false;

        private eDessin _Dessin = null;
        private Sheet _SwFeuille = null;
        #endregion

        #region "Constructeur\Destructeur"

        public eFeuille() { }

        #endregion

        #region "Propriétés"

        /// <summary>
        /// Retourne l'objet Sheet associé.
        /// </summary>
        public Sheet SwFeuille { get { Debug.Print(MethodBase.GetCurrentMethod()); return _SwFeuille; } }

        /// <summary>
        /// Retourne le parent ExtDessin.
        /// </summary>
        public eDessin Dessin { get { Debug.Print(MethodBase.GetCurrentMethod()); return _Dessin; } }

        /// <summary>
        /// Retourne ou défini le nom de la feuille.
        /// </summary>
        public String Nom { get { Debug.Print(MethodBase.GetCurrentMethod()); return _SwFeuille.GetName(); } set { Debug.Print(MethodBase.GetCurrentMethod()); _SwFeuille.SetName(value); } }

        /// <summary>
        /// Retourne la première vue du dessin
        /// </summary>
        public eVue PremiereVue
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                eVue pVue = new eVue();

                object[] pObjVues = _SwFeuille.GetViews();

                if ((pObjVues != null) && pVue.Init((View)pObjVues[0], this))
                    return pVue;

                return null;
            }
        }

        /// <summary>
        /// Retourne la dernière vue crée
        /// </summary>
        public eVue DerniereVue
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                eVue pVue = new eVue();

                object[] pObjVues = _SwFeuille.GetViews();

                if ((pObjVues != null) && pVue.Init((View)pObjVues[pObjVues.GetUpperBound(0)], this))
                    return pVue;

                return null;
            }
        }

        /// <summary>
        /// Retourne le nb de vues de la feuille
        /// </summary>
        public int NbDeVues
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());

                object[] pObjVues = _SwFeuille.GetViews();

                if (pObjVues == null)
                    return 0;

                return pObjVues.Length;
            }
        }

        /// <summary>
        /// Retourne la zone enveloppant les vues
        /// </summary>
        public eZone EnveloppeDesVues
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());

                eZone pEnveloppe = new eZone();

                pEnveloppe.PointMax.X = 0;
                pEnveloppe.PointMax.Y = 0;
                pEnveloppe.PointMin.X = 10000;
                pEnveloppe.PointMin.Y = 10000;

                List<eVue> pListeVues = ListListeDesVues();

                Debug.Print("============================= " + pListeVues.Count.ToString());

                if (pListeVues.Count == 0)
                {
                    pEnveloppe.PointMin.X = 0;
                    pEnveloppe.PointMin.Y = 0;
                    return pEnveloppe;
                }

                foreach (eVue Vue in pListeVues)
                {
                    pEnveloppe.PointMax.X = Math.Max(pEnveloppe.PointMax.X, Vue.Dimensions.Zone.PointMax.X);
                    pEnveloppe.PointMax.Y = Math.Max(pEnveloppe.PointMax.Y, Vue.Dimensions.Zone.PointMax.Y);
                    pEnveloppe.PointMin.X = Math.Min(pEnveloppe.PointMin.X, Math.Max(0, Vue.Dimensions.Zone.PointMin.X));
                    pEnveloppe.PointMin.Y = Math.Min(pEnveloppe.PointMin.Y, Math.Max(0, Vue.Dimensions.Zone.PointMin.Y));
                }

                return pEnveloppe;
            }
        }

        /// <summary>
        /// Retourne le format de la feuille
        /// </summary>
        public Format_e Format
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());

                Double pLargeur = 0;
                Double pHauteur = 0;
                swDwgPaperSizes_e pTaille = (swDwgPaperSizes_e)_SwFeuille.GetSize(ref pLargeur, ref pHauteur);

                if (Orientation == Orientation_e.cPortrait)
                {
                    Double pTmp = pLargeur;
                    pLargeur = pHauteur;
                    pHauteur = pTmp;
                }

                if ((pLargeur == 1.189) && (pHauteur == 0.841))
                    return Format_e.cA0;

                if ((pLargeur == 0.841) && (pHauteur == 0.594))
                    return Format_e.cA1;

                if ((pLargeur == 0.594) && (pHauteur == 0.420))
                    return Format_e.cA2;

                if ((pLargeur == 0.420) && (pHauteur == 0.297))
                    return Format_e.cA3;

                if ((pLargeur == 0.297) && (pHauteur == 0.210))
                    return Format_e.cA4;

                if ((pLargeur == 0.210) && (pHauteur == 0.148))
                    return Format_e.cA5;

                return Format_e.cUtilisateur;
            }
            set
            {
                Debug.Print(MethodBase.GetCurrentMethod());

                Double pLargeur = 0;
                Double pHauteur = 0;

                switch (value)
                {
                    case Format_e.cA0:
                        pLargeur = 1.189; pHauteur = 0.841;
                        break;
                    case Format_e.cA1:
                        pLargeur = 0.841; pHauteur = 0.594;
                        break;
                    case Format_e.cA2:
                        pLargeur = 0.594; pHauteur = 0.420;
                        break;
                    case Format_e.cA3:
                        pLargeur = 0.420; pHauteur = 0.297;
                        break;
                    case Format_e.cA4:
                        pLargeur = 0.297; pHauteur = 0.210;
                        break;
                    case Format_e.cA5:
                        pLargeur = 0.210; pHauteur = 0.148;
                        break;
                }

                if (Orientation == Orientation_e.cPortrait)
                {
                    Double pTmp = pLargeur;
                    pLargeur = pHauteur;
                    pHauteur = pTmp;
                }

                _SwFeuille.SetSize(12, pLargeur, pHauteur);
            }
        }

        /// <summary>
        /// Retourne l'orientation de la feuille : Portrait ou Paysage
        /// </summary>
        public Orientation_e Orientation
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());

                Double pLargeur = 0;
                Double pHauteur = 0;
                swDwgPaperSizes_e pTaille = (swDwgPaperSizes_e)_SwFeuille.GetSize(ref pLargeur, ref pHauteur);
                if (pLargeur > pHauteur)
                    return Orientation_e.cPaysage;

                return Orientation_e.cPortrait;
            }
            set
            {
                Debug.Print(MethodBase.GetCurrentMethod());

                Double pLargeur = 0;
                Double pHauteur = 0;
                swDwgPaperSizes_e pTaille = (swDwgPaperSizes_e)_SwFeuille.GetSize(ref pLargeur, ref pHauteur);

                if (Orientation != value)
                    _SwFeuille.SetSize(12, pHauteur, pLargeur);
            }
        }

        /// <summary>
        /// Retourne le gabarit de la feuille
        /// </summary>
        public String GabaritDeFeuille
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                return _SwFeuille.GetTemplateName();
            }
            set
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                _SwFeuille.SetTemplateName(value);
            }
        }

        /// <summary>
        /// Fonction interne.
        /// Test l'initialisation de l'objet ExtFeuille.
        /// </summary>
        internal Boolean EstInitialise { get { Debug.Print(MethodBase.GetCurrentMethod()); return _EstInitialise; } }

        #endregion

        #region "Méthodes"

        /// <summary>
        /// Méthode interne.
        /// Initialiser l'objet ExtFeuille.
        /// </summary>
        /// <param name="SwFeuille"></param>
        /// <param name="Dessin"></param>
        /// <returns></returns>
        internal Boolean Init(Sheet SwFeuille, eDessin Dessin)
        {
            Debug.Print(MethodBase.GetCurrentMethod());

            if ((SwFeuille != null) && (Dessin != null) && Dessin.EstInitialise)
            {
                _Dessin = Dessin;
                _SwFeuille = SwFeuille;

                Debug.Print(this.Nom);
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
        /// Initialiser l'objet ExtFeuille.
        /// </summary>
        /// <param name="SwFeuille"></param>
        /// <param name="Modele"></param>
        /// <returns></returns>
        internal Boolean Init(Sheet SwFeuille, eModele Modele)
        {
            Debug.Print(MethodBase.GetCurrentMethod());

            if ((SwFeuille != null) && (Modele != null) && Modele.EstInitialise && (Modele.TypeDuModele == TypeFichier_e.cDessin))
            {
                _Dessin = Modele.Dessin;
                _SwFeuille = SwFeuille;

                Debug.Print(this.Nom);
                _EstInitialise = true;
            }
            else
            {
                Debug.Print("!!!!! Erreur d'initialisation");
            }
            return _EstInitialise;
        }

        /// <summary>
        /// Activer la feuille.
        /// </summary>
        public void Activer()
        {
            Debug.Print(MethodBase.GetCurrentMethod());
            Dessin.SwDessin.ActivateSheet(Nom);
        }

        /// <summary>
        /// Supprimer la feuille.
        /// </summary>
        public void Supprimer()
        {
            Debug.Print(MethodBase.GetCurrentMethod());

            eModele pModeleDessin = Dessin.Modele;
            pModeleDessin.Activer();

            pModeleDessin.SwModele.Extension.SelectByID2(Nom, "SHEET", 0, 0, 0, false, 0, null, 0);
            pModeleDessin.SwModele.DeleteSelection(false);
            pModeleDessin.SwModele.ClearSelection2(true);
        }

        /// <summary>
        /// Zoom étendu de la feuille.
        /// </summary>
        public void ZoomEtendu()
        {
            Debug.Print(MethodBase.GetCurrentMethod());
            Dessin.Modele.SW.Modele().SwModele.ViewZoomtofit2();
        }

        /// <summary>
        /// Ajuster la feuille autour des vues
        /// </summary>
        public void AjusterAutourDesVues()
        {
            Debug.Print(MethodBase.GetCurrentMethod());

            eZone pEnveloppe = EnveloppeDesVues;

            if (pEnveloppe == null)
                return;

            Redimensionner(pEnveloppe.PointMax.X + pEnveloppe.PointMin.X, pEnveloppe.PointMax.Y + pEnveloppe.PointMin.Y);
        }

        /// <summary>
        /// Redimensionne la feuille
        /// </summary>
        public void Redimensionner(Double Largeur, Double Hauteur)
        {
            Debug.Print(MethodBase.GetCurrentMethod());

            if ((Largeur == 0) || (Hauteur == 0))
                return;

            _SwFeuille.SetSize((int)swDwgPaperSizes_e.swDwgPapersUserDefined, Largeur, Hauteur);

        }

        /// <summary>
        /// Creer la vue développée d'une tôle
        /// </summary>
        /// <param name="Piece"></param>
        /// <param name="Configuration"></param>
        /// <returns></returns>
        public eVue CreerVueToleDvp(ePiece Piece, eConfiguration Configuration, Boolean AfficherLesLignesDePliage)
        {
            eVue pVue = new eVue();
            View pSwVue = null;
            Configuration.Activer();

            // Si des corps autre que la tole dépliée sont encore visible dans la config, on les cache et on recontruit tout
            List<eCorps> pListeCorps = Piece.ListListeDesCorps("^((?!" + CONSTANTES.NOM_CORPS_DEPLIEE + ").)*$");
            foreach (eCorps pCorps in pListeCorps)
                pCorps.Visible = false;

            if (pListeCorps.Count > 0)
                Piece.Modele.ForcerAToutReconstruire();

            _Dessin.Modele.Activer();
            pSwVue = _Dessin.SwDessin.CreateFlatPatternViewFromModelView3(Piece.Modele.FichierSw.Chemin, Configuration.Nom, 0, 0, 0, AfficherLesLignesDePliage, false);

            if (pVue.Init(pSwVue, this))
            {
                Debug.Print("Vue dvp crée");
                return pVue;
            }

            Debug.Print("Vue dvp non crée");
            return null;
        }

        /// <summary>
        /// Exporte la feuille en DXF
        /// Le nom de la feuille correspond au nom du fichier
        /// </summary>
        /// <param name="CheminDossier"></param>
        /// <param name="NomDuFichierAlternatif"></param>
        public void ExporterEn(Extension_e TypeExport, String CheminDossier, String NomDuFichierAlternatif = "")
        {
            Debug.Print(MethodBase.GetCurrentMethod());

            Activer();
            ZoomEtendu();

            String Ext = "";
            ExportPdfData OptionsPDF = null;
            switch (TypeExport)
            {
                case Extension_e.cDXF:
                    Ext = ".dxf";
                    _Dessin.Modele.SW.GestOptions.DxfDwg_ExporterToutesLesFeuilles = swDxfMultisheet_e.swDxfActiveSheetOnly;
                    break;
                case Extension_e.cDWG:
                    Ext = ".dwg";
                    _Dessin.Modele.SW.GestOptions.DxfDwg_ExporterToutesLesFeuilles = swDxfMultisheet_e.swDxfActiveSheetOnly;
                    break;
                case Extension_e.cPDF:
                    Ext = ".pdf";
                    OptionsPDF = _Dessin.Modele.SW.SwSW.GetExportFileData((int)swExportDataFileType_e.swExportPdfData);
                    DispatchWrapper[] Wrapper = new DispatchWrapper[1];
                    Wrapper[0] = new DispatchWrapper((Object)SwFeuille);
                    OptionsPDF.SetSheets((int)swExportDataSheetsToExport_e.swExportData_ExportSpecifiedSheets, Wrapper);
                    break;
            }

            String CheminFichier = this.Nom;

            int Erreur = 0;
            int Warning = 0;

            if (!String.IsNullOrEmpty(NomDuFichierAlternatif))
                CheminFichier = NomDuFichierAlternatif;

            CheminFichier = Path.Combine(CheminDossier, CheminFichier + Ext);
            Dessin.Modele.SwModele.Extension.SaveAs(CheminFichier, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, OptionsPDF, Erreur, Warning);
            Debug.Print(CheminFichier);
        }

        /// <summary>
        /// Méthode interne.
        /// Renvoi la liste des vues de la feuille filtrée par les arguments.
        /// Si NomARechercher est vide, toutes les vues sont retournées.
        /// </summary>
        /// <param name="NomARechercher"></param>
        /// <returns></returns>
        internal List<eVue> ListListeDesVues(String NomARechercher = "")
        {
            Debug.Print(MethodBase.GetCurrentMethod());

            List<eVue> pListeVues = new List<eVue>();

            Object[] pTabVues = _SwFeuille.GetViews();

            if (pTabVues == null)
                return pListeVues;

            foreach (View pSwVue in pTabVues)
            {
                eVue pVue = new eVue();
                if (pVue.Init(pSwVue, this) && Regex.IsMatch(pSwVue.GetName2(), NomARechercher))
                    pListeVues.Add(pVue);
            }

            return pListeVues;
        }

        /// <summary>
        /// Renvoi la liste des vues de la feuille filtrée par les arguments.
        /// Si NomARechercher est vide, toutes les vues sont retournées.
        /// </summary>
        /// <param name="NomARechercher"></param>
        /// <returns></returns>
        public ArrayList ListeDesVues(String NomARechercher = "")
        {
            Debug.Print(MethodBase.GetCurrentMethod());

            List<eVue> pListeVues = ListListeDesVues(NomARechercher);
            ArrayList pArrayVues = new ArrayList();

            if (pListeVues.Count > 0)
                pArrayVues = new ArrayList(pListeVues);

            return pArrayVues;
        }

        /// <summary>
        /// Definit les paramètres pour l'impression
        /// </summary>
        /// <param name="Couleur"></param>
        /// <param name="HauteQualite"></param>
        public void MettreEnPagePourImpression(swPageSetupDrawingColor_e Couleur = swPageSetupDrawingColor_e.swPageSetup_AutomaticDrawingColor, Boolean HauteQualite = false)
        {
            _Dessin.Modele.SwModele.Extension.UsePageSetup = (int)swPageSetupInUse_e.swPageSetupInUse_Application;
            PageSetup pSetupApp = _Dessin.Modele.SwModele.Extension.AppPageSetup;
            pSetupApp.HighQuality = HauteQualite;
            pSetupApp.DrawingColor = (int)Couleur;
            pSetupApp.ScaleToFit = false;

            _Dessin.Modele.SwModele.Extension.UsePageSetup = (int)swPageSetupInUse_e.swPageSetupInUse_Document;
            PageSetup pSetupDoc = _Dessin.Modele.SwModele.PageSetup;
            pSetupDoc.HighQuality = HauteQualite;
            pSetupDoc.DrawingColor = (int)Couleur;
            pSetupDoc.ScaleToFit = false;

            _Dessin.Modele.SwModele.Extension.UsePageSetup = (int)swPageSetupInUse_e.swPageSetupInUse_DrawingSheet;
            PageSetup pSetupFeuille = _SwFeuille.PageSetup;
            pSetupFeuille.DrawingColor = (int)Couleur;
            pSetupFeuille.HighQuality = HauteQualite;
            pSetupFeuille.PrinterPaperSource = 15;
            pSetupFeuille.PrinterPaperSize = (int)Format;
            pSetupFeuille.ScaleToFit = false;

            if (Orientation == Orientation_e.cPaysage)
                pSetupFeuille.Orientation = (int)swPageSetupOrientation_e.swPageSetupOrient_Landscape;
            else
                pSetupFeuille.Orientation = (int)swPageSetupOrientation_e.swPageSetupOrient_Portrait;


        }

        #endregion

    }
}