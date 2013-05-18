using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Reflection;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace Framework_SW2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("E01F2AAF-807F-4294-A428-240C26DF0269")]
    public interface IGestDeSelection
    {
        SelectionMgr SwGestDeSelection { get; }
        eModele Modele { get; }
        ArrayList ListeDesComposantsSelectionnes(String NomComposant = "", int Marque = -1);
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("A3D09048-4B5D-4FCC-A1F2-F1ACAACB0E4C")]
    [ProgId("Frameworks.GestDeSelection")]
    public class GestDeSelection : IGestDeSelection
    {
        #region "Variables locales"

        private Boolean _EstInitialise = false;

        private eModele _Modele;
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
        public eModele Modele { get { Debug.Info(MethodBase.GetCurrentMethod()); return _Modele; } }

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
        internal Boolean Init(SelectionMgr SwGestionnaire, eModele Modele)
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

        /// <summary>
        /// Fonction interne, renvoi la liste des composants sélectionnés
        /// </summary>
        /// <param name="NomComposant"></param>
        /// <returns></returns>
        internal List<ExtComposant> ListListeDesComposantsSelectionnes(String NomComposant = "", int Marque = -1)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            // Liste à renvoyer
            List<ExtComposant> pListeComposants = new List<ExtComposant>();

            if ((_Modele.TypeDuModele == TypeFichier_e.cAssemblage) || (_Modele.TypeDuModele == TypeFichier_e.cDessin))
            {

                for (int i = 1; i <= _SwGestDeSelection.GetSelectedObjectCount2(-1); i++)
                {
                    Component2 pSwComposant = _SwGestDeSelection.GetSelectedObjectsComponent4(i, Marque);
                    // Si le composant est null, on passe au suivant
                    // C'est le cas, quand on sélectionne le composant racine. A voir pour trouver une solution 
                    if ((pSwComposant == null) && (_SwGestDeSelection.GetSelectedObjectType3(i, Marque) == (int)swSelectType_e.swSelCOMPONENTS))
                    {
                        pSwComposant = _SwGestDeSelection.GetSelectedObject6(i, Marque);
                    }
                    // Pour intitialiser le composant correctement il faut un peu de bidouille
                    // sinon on à le droit à une belle reference circulaire
                    // Donc d'abord, on recherche le modele du SwComposant
                    Debug.Info(pSwComposant.GetPathName());
                    eModele pModele = _Modele.SW.Modele(pSwComposant.GetPathName());
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

        /// <summary>
        /// Renvoi la liste des composants sélectionnés dans l'assemblage
        /// </summary>
        /// <param name="NomComposant"></param>
        /// <returns></returns>
        public ArrayList ListeDesComposantsSelectionnes(String NomComposant = "", int Marque = -1)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            List<ExtComposant> pListeComps = ListListeDesComposantsSelectionnes(NomComposant, Marque);
            ArrayList pArrayComps = new ArrayList();

            if (pListeComps.Count > 0)
                pArrayComps = new ArrayList(pListeComps);

            return pArrayComps;
        }

        #endregion
    }
}
