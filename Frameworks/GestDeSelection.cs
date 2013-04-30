using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Reflection;
using SolidWorks.Interop.sldworks;

namespace Framework_SW2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("E01F2AAF-807F-4294-A428-240C26DF0269")]
    public interface IGestDeSelection
    {
        SelectionMgr SwGestDeSelection { get; }
        ExtModele Modele { get; }
        ArrayList ListeSelectionDesComposants(String NomComposant = "");
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("A3D09048-4B5D-4FCC-A1F2-F1ACAACB0E4C")]
    [ProgId("Frameworks.GestDeSelection")]
    public class GestDeSelection : IGestDeSelection
    {
        #region "Variables locales"

        private Boolean _EstInitialise = false;

        private ExtModele _Modele;
        private SelectionMgr _SwGestDeSelection;

        #endregion

        #region "Constructeur\Destructeur"

        public GestDeSelection() { }

        #endregion

        #region "Propriétés"

        /// <summary>
        /// Renvoi l'objet SelectionMgr.
        /// </summary>
        public SelectionMgr SwGestDeSelection { get { Debug.Info(MethodBase.GetCurrentMethod()); return _SwGestDeSelection; } }

        /// <summary>
        /// Retourne le parent ExtModele.
        /// </summary>
        public ExtModele Modele { get { Debug.Info(MethodBase.GetCurrentMethod()); return _Modele; } }

        /// <summary>
        /// Fonction interne
        /// Test l'initialisation de l'objet GestDeSelection
        /// </summary>
        internal Boolean EstInitialise { get { Debug.Info(MethodBase.GetCurrentMethod()); return _EstInitialise; } }

        #endregion

        #region "Méthodes"

        /// <summary>
        /// Méthode interne.
        /// Initialiser l'objet GestDeSelection.
        /// </summary>
        /// <param name="Modele"></param>
        /// <returns></returns>
        internal Boolean Init(SelectionMgr SwGestionnaire, ExtModele Modele)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            if ((SwGestionnaire != null) && (Modele != null) && Modele.EstInitialise)
            {
                _SwGestDeSelection = SwGestionnaire;
                _Modele = Modele;
                _EstInitialise = true;
            }
            else
            {
                Debug.Info("!!!!! Erreur d'initialisation");
            }

            return _EstInitialise;
        }

        internal List<ExtComposant> ListListeSelectionDesComposants(String NomComposant = "")
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            // Liste à renvoyer
            List<ExtComposant> pListeComposants = new List<ExtComposant>();

            int NbSel = _SwGestDeSelection.GetSelectedObjectCount2(-1);

            if (NbSel > 0)
            {
                for (int i = 1; i <= NbSel; i++)
                {
                    Component2 pSwComposant = _SwGestDeSelection.GetSelectedObjectsComponent4(i, -1);
                    // Pour intitialiser le composant correctement il faut un peu de bidouille
                    // sinon on à le droit à une belle reference circulaire
                    // Donc d'abord, on recherche le modele du SwComposant
                    Debug.Info(pSwComposant.GetPathName());
                    ExtModele pModele = _Modele.SW.Modele(pSwComposant.GetPathName());
                    // Ensuite, on créer un nouveau Composant avec la ref du SwComposant et du modele
                    ExtComposant pComposant = new ExtComposant();
                    // Et pour que les deux soit liés, on passe la ref du Composant que l'on vient de creer
                    // au modele. Comme ca, Modele.Composant pointe sur Composant et Composant.Modele pointe sur Modele,
                    // la boucle est bouclée
                    pComposant.Init(pSwComposant, pModele);
                    pModele.Composant = pComposant;
                    pListeComposants.Add(pComposant);
                }

                // On trie et c'est parti
                pListeComposants.Sort();
            }

            return pListeComposants;
        }

        public ArrayList ListeSelectionDesComposants(String NomComposant = "")
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            List<ExtComposant> pListeComps = ListListeSelectionDesComposants(NomComposant);
            ArrayList pArrayComps = new ArrayList();

            if (pListeComps.Count > 0)
                pArrayComps = new ArrayList(pListeComps);

            return pArrayComps;
        }

        #endregion
    }
}
