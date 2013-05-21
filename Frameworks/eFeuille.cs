using System;
using System.Collections;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Reflection;
using SolidWorks.Interop.swconst;
using System.IO;

namespace Framework_SW2013
{

    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("17F1BCFD-2428-4DF1-8338-8FFA142E2A97")]
    public interface IeFeuille
    {
        Sheet SwFeuille { get; }
        eDessin Dessin { get; }
        String Nom { get; set; }
        eVue PremiereVue { get; }
        void Activer();
        void Supprimer();
        void ZoomEtendu();
        void Redimensionner();
        void ExporterEnDXF(String CheminDossier, String NomDuFichierAlternatif = "");
        ArrayList ListeDesVues(String NomARechercher = "");
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("AB11E456-34CF-4540-A7E3-E01D7C63E324")]
    [ProgId("Frameworks.eFeuille")]
    public class eFeuille : IeFeuille
    {
        #region "Variables locales"

        private Boolean _EstInitialise = false;

        private eDessin _Dessin;
        private Sheet _SwFeuille;
        #endregion

        #region "Constructeur\Destructeur"

        public eFeuille() { }

        #endregion

        #region "Propriétés"

        /// <summary>
        /// Retourne l'objet Sheet associé.
        /// </summary>
        public Sheet SwFeuille { get { Debug.Info(MethodBase.GetCurrentMethod()); return _SwFeuille; } }

        /// <summary>
        /// Retourne le parent ExtDessin.
        /// </summary>
        public eDessin Dessin { get { Debug.Info(MethodBase.GetCurrentMethod()); return _Dessin; } }

        /// <summary>
        /// Retourne ou défini le nom de la feuille.
        /// </summary>
        public String Nom { get { Debug.Info(MethodBase.GetCurrentMethod()); return _SwFeuille.GetName(); } set { Debug.Info(MethodBase.GetCurrentMethod()); _SwFeuille.SetName(value); } }

        /// <summary>
        /// Retourne la première vue du dessin
        /// </summary>
        public eVue PremiereVue
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                eVue pVue = new eVue();

                object[] pObjVues;
                pObjVues = _SwFeuille.GetViews();

                if ((pObjVues.Length > 0) && pVue.Init((View)pObjVues[0], this))
                    return pVue;

                return null;
            }
        }

        /// <summary>
        /// Fonction interne.
        /// Test l'initialisation de l'objet ExtFeuille.
        /// </summary>
        internal Boolean EstInitialise { get { Debug.Info(MethodBase.GetCurrentMethod()); return _EstInitialise; } }

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
            Debug.Info(MethodBase.GetCurrentMethod());

            if ((SwFeuille != null) && (Dessin != null) && Dessin.EstInitialise)
            {
                _Dessin = Dessin;
                _SwFeuille = SwFeuille;

                Debug.Info(this.Nom);
                _EstInitialise = true;
            }
            else
            {
                Debug.Info("!!!!! Erreur d'initialisation");
            }
            return _EstInitialise;
        }

        /// <summary>
        /// Activer la feuille.
        /// </summary>
        public void Activer()
        {
            Dessin.SwDessin.ActivateSheet(Nom);
        }

        /// <summary>
        /// Supprimer la feuille.
        /// </summary>
        public void Supprimer()
        {
            Dessin.Modele.SwModele.Extension.SelectByID2(Nom, "SHEET", 0, 0, 0, false, 0, null, 0);
            Dessin.Modele.SwModele.DeleteSelection(false);
            Dessin.Modele.SwModele.ClearSelection2(true);
        }

        /// <summary>
        /// Zoom étendu de la feuille.
        /// </summary>
        public void ZoomEtendu()
        {
            Dessin.Modele.SwModele.ViewZoomtofit2();
        }

        /// <summary>
        /// Redimensionne la feuille autour des vues
        /// </summary>
        public void Redimensionner()
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            eZone pEnveloppe = new eZone();

            pEnveloppe.PointMax.X = 0;
            pEnveloppe.PointMax.Y = 0;
            pEnveloppe.PointMin.X = 10000;
            pEnveloppe.PointMin.Y = 10000;

            List<eVue> pListeVues = ListListeDesVues();

            if (pListeVues.Count == 0)
                return;

            foreach (eVue Vue in pListeVues)
            {
                pEnveloppe.PointMax.X = Math.Max(pEnveloppe.PointMax.X, Vue.Dimensions.Zone.PointMax.X);
                pEnveloppe.PointMax.Y = Math.Max(pEnveloppe.PointMax.Y, Vue.Dimensions.Zone.PointMax.Y);
                pEnveloppe.PointMin.X = Math.Min(pEnveloppe.PointMin.X, Math.Max(0, Vue.Dimensions.Zone.PointMin.X));
                pEnveloppe.PointMin.Y = Math.Min(pEnveloppe.PointMin.Y, Math.Max(0, Vue.Dimensions.Zone.PointMin.Y));
            }

            _SwFeuille.SetSize((int)swDwgPaperSizes_e.swDwgPapersUserDefined,
                pEnveloppe.PointMax.X + pEnveloppe.PointMin.X,
                pEnveloppe.PointMax.Y + pEnveloppe.PointMin.Y);

        }

        /// <summary>
        /// Exporte la feuille en DXF
        /// Le nom de la feuille correspond au nom du fichier
        /// </summary>
        /// <param name="CheminDossier"></param>
        /// <param name="NomDuFichierAlternatif"></param>
        public void ExporterEnDXF(String CheminDossier, String NomDuFichierAlternatif = "")
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            Activer();
            ZoomEtendu();

            String CheminFichier = this.Nom;

            int Erreur = 0;
            int Warning = 0;

            if (!String.IsNullOrEmpty(NomDuFichierAlternatif))
                CheminFichier = NomDuFichierAlternatif;

            CheminFichier = Path.Combine(CheminDossier, CheminFichier + ".dxf");
            Dessin.Modele.SwModele.Extension.SaveAs(CheminFichier, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, null, Erreur, Warning);
            Debug.Info(CheminFichier);
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
            Debug.Info(MethodBase.GetCurrentMethod());

            List<eVue> pListeVues = new List<eVue>();

            object[] pObjVues;
            pObjVues = _SwFeuille.GetViews();

            if (pObjVues.Length == 0)
                return pListeVues;

            foreach (View pSwVue in pObjVues)
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
            Debug.Info(MethodBase.GetCurrentMethod());

            List<eVue> pListeVues = ListListeDesVues(NomARechercher);
            ArrayList pArrayVues = new ArrayList();

            if (pListeVues.Count > 0)
                pArrayVues = new ArrayList(pListeVues);

            return pArrayVues;
        }

        #endregion

    }
}
