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

        public Sheet SwFeuille { get { Debug.Info(MethodBase.GetCurrentMethod());  return _SwFeuille; } }

        public ExtDessin Dessin { get { Debug.Info(MethodBase.GetCurrentMethod());  return _Dessin; } }

        public String Nom { get { Debug.Info(MethodBase.GetCurrentMethod());  return _SwFeuille.GetName(); } set { Debug.Info(MethodBase.GetCurrentMethod());  _SwFeuille.SetName(value); } }

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

        internal Boolean EstInitialise { get { Debug.Info(MethodBase.GetCurrentMethod());  return _EstInitialise; } }

        #endregion

        #region "Méthodes"

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
