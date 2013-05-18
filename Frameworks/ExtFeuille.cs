using System;
using System.Collections;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Reflection;

namespace Framework_SW2013
{

    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("17F1BCFD-2428-4DF1-8338-8FFA142E2A97")]
    public interface IExtFeuille
    {
        Sheet SwFeuille { get; }
        ExtDessin Dessin { get; }
        String Nom { get; set; }
        ExtVue PremiereVue { get; }
        void Activer();
        void Supprimer();
        void ZoomEtendu();
        ArrayList ListeDesVues(String NomARechercher = "");
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("AB11E456-34CF-4540-A7E3-E01D7C63E324")]
    [ProgId("Frameworks.ExtFeuille")]
    public class ExtFeuille : IExtFeuille
    {
        #region "Variables locales"
        
        private Boolean _EstInitialise = false;

        private ExtDessin _Dessin;
        private Sheet _SwFeuille;
        #endregion

        #region "Constructeur\Destructeur"

        public ExtFeuille() { }

        #endregion

        #region "Propriétés"

        /// <summary>
        /// Retourne l'objet Sheet associé.
        /// </summary>
        public Sheet SwFeuille { get { Debug.Info(MethodBase.GetCurrentMethod());  return _SwFeuille; } }

        /// <summary>
        /// Retourne le parent ExtDessin.
        /// </summary>
        public ExtDessin Dessin { get { Debug.Info(MethodBase.GetCurrentMethod());  return _Dessin; } }

        /// <summary>
        /// Retourne ou défini le nom de la feuille.
        /// </summary>
        public String Nom { get { Debug.Info(MethodBase.GetCurrentMethod());  return _SwFeuille.GetName(); } set { Debug.Info(MethodBase.GetCurrentMethod());  _SwFeuille.SetName(value); } }

        /// <summary>
        /// Retourne la première vue du dessin
        /// </summary>
        public ExtVue PremiereVue
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                ExtVue pVue = new ExtVue();

                object[] pObjVues;
                pObjVues = _SwFeuille.GetViews();

                if ((pObjVues.Length > 0) && pVue.Init((View)pObjVues[0],this))
                    return pVue;

                return null;
            }
        }

        /// <summary>
        /// Fonction interne.
        /// Test l'initialisation de l'objet ExtFeuille.
        /// </summary>
        internal Boolean EstInitialise { get { Debug.Info(MethodBase.GetCurrentMethod());  return _EstInitialise; } }

        #endregion

        #region "Méthodes"

        /// <summary>
        /// Méthode interne.
        /// Initialiser l'objet ExtFeuille.
        /// </summary>
        /// <param name="SwFeuille"></param>
        /// <param name="Dessin"></param>
        /// <returns></returns>
        internal Boolean Init(Sheet SwFeuille, ExtDessin Dessin)
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
        /// Méthode interne.
        /// Renvoi la liste des vues de la feuille filtrée par les arguments.
        /// Si NomARechercher est vide, toutes les vues sont retournées.
        /// </summary>
        /// <param name="NomARechercher"></param>
        /// <returns></returns>
        internal List<ExtVue> ListListeDesVues(String NomARechercher = "")
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            List<ExtVue> pListeVues = new List<ExtVue>();

            object[] pObjVues;
            pObjVues = _SwFeuille.GetViews();

            if (pObjVues.Length == 0)
                return pListeVues;

            foreach (View pSwVue in pObjVues)
            {
                ExtVue pVue = new ExtVue();
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

            List<ExtVue> pListeVues = ListListeDesVues(NomARechercher);
            ArrayList pArrayVues = new ArrayList();

            if (pListeVues.Count > 0)
                pArrayVues = new ArrayList(pListeVues);

            return pArrayVues;
        }

        #endregion

    }
}
