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
    public interface IeGestDeSelection
    {
        SelectionMgr SwGestDeSelection { get; }
        eModele Modele { get; }
        int NbObjetsSelectionnes(int Marque = -1);
        ePoint PointDeSelectionObjet(int Index, int Marque = -1);
        swSelectType_e TypeObjet(int Index, int Marque = -1);
        int MarqueObjet(int Index);
        dynamic Objet(int Index, int Marque = -1);
        eComposant Composant(int Index, int Marque = -1);
        eVue Vue(int Index, int Marque = -1);
        ArrayList ListeDesObjetsSelectionnes(swSelectType_e TypeObjet, int Marque = -1);
        ArrayList ListeDesComposantsSelectionnes(String NomComposant = "", int Marque = -1);
        ArrayList ListeDesVuesSelectionnees(int Marque = -1);
        // GetSelectedObjectsFace
        // GetSelectedObjectLoop2
        // GetSelectionPointInSketchSpace2
        // IsInEditTarget2

        // SetSelectionPoint2

    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("A3D09048-4B5D-4FCC-A1F2-F1ACAACB0E4C")]
    [ProgId("Frameworks.eGestDeSelection")]
    public class eGestDeSelection : IeGestDeSelection
    {
        #region "Variables locales"

        private Boolean _EstInitialise = false;

        private eModele _Modele;
        private SelectionMgr _SwGestDeSelection;

        #endregion

        #region "Constructeur\Destructeur"

        public eGestDeSelection() { }

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
        /// Renvoi le nb d'objets selectionnés
        /// </summary>
        /// <param name="Marque"></param>
        /// <returns></returns>
        public int NbObjetsSelectionnes(int Marque = -1)
        {
            Debug.Info(MethodBase.GetCurrentMethod());
            return _SwGestDeSelection.GetSelectedObjectCount2(Marque);
        }

        /// <summary>
        /// Retourne le point de sélection de l'objet.
        /// </summary>
        /// <param name="Index"></param>
        /// <param name="Marque"></param>
        /// <returns></returns>
        public ePoint PointDeSelectionObjet(int Index, int Marque = -1)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            Double[] pSwPoint = _SwGestDeSelection.GetSelectionPoint2(Index, Marque);
            return new ePoint(pSwPoint[0], pSwPoint[1], pSwPoint[2]);
        }

        /// <summary>
        /// Retourne le type de l'objet sélectionné.
        /// </summary>
        /// <param name="Index"></param>
        /// <param name="Marque"></param>
        /// <returns></returns>
        public swSelectType_e TypeObjet(int Index, int Marque = -1)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            if (NbObjetsSelectionnes(Marque) > 0)
                return (swSelectType_e)_SwGestDeSelection.GetSelectedObjectType3(Index, Marque);

            return swSelectType_e.swSelNOTHING;
        }

        /// <summary>
        /// Retourne le marque de l'objet sélectionné.
        /// </summary>
        /// <param name="Index"></param>
        /// <returns></returns>
        public int MarqueObjet(int Index)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            if (NbObjetsSelectionnes() > 0)
                return _SwGestDeSelection.GetSelectedObjectMark(Index);

            return 0;
        }

        /// <summary>
        /// Retourne l'objet sélectionné.
        /// </summary>
        /// <param name="Index"></param>
        /// <param name="Marque"></param>
        /// <returns></returns>
        public dynamic Objet(int Index, int Marque = -1)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            if (NbObjetsSelectionnes() == 0)
                return null;

            eObjet pObjet = new eObjet();

            eComposant pComposant = Composant(Index, Marque);

            if ((pComposant != null) && pComposant.EstInitialise
                && (pObjet.Init(pComposant.Modele, _SwGestDeSelection.GetSelectedObject6(Index, Marque), TypeObjet(Index, Marque))))
                return pObjet.Objet;

            return null;
        }

        /// <summary>
        /// Retourne le composant associé à l'objet sélectionné.
        /// </summary>
        /// <param name="Index"></param>
        /// <param name="Marque"></param>
        /// <returns></returns>
        public eComposant Composant(int Index, int Marque = -1)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            if (NbObjetsSelectionnes() == 0)
                return null;

            Component2 pSwComposant = _SwGestDeSelection.GetSelectedObjectsComponent4(Index, Marque);

            // C'est le cas, quand on sélectionne le composant racine. A voir pour trouver une solution 
            if (pSwComposant == null)
            {
                pSwComposant = _SwGestDeSelection.GetSelectedObject6(Index, Marque);
            }

            if (pSwComposant == null)
                return null;

            // Pour intitialiser le composant correctement il faut un peu de bidouille
            // sinon on à le droit à une belle reference circulaire
            // Donc d'abord, on recherche le modele du SwComposant
            Debug.Info(pSwComposant.GetPathName());
            eModele pModele = _Modele.SW.Modele(pSwComposant.GetPathName());
            // Ensuite, on créer un nouveau Composant avec la ref du SwComposant et du modele
            eComposant pComposant = new eComposant();
            // Et pour que les deux soit liés, on passe la ref du Composant que l'on vient de creer
            // au modele. Comme ca, Modele.Composant pointe sur Composant et Composant.Modele pointe sur Modele,
            // la boucle est bouclée
            pComposant.Init(pSwComposant, pModele);
            pModele.Composant = pComposant;

            if (pComposant.EstInitialise)
                return pComposant;

            return null;
        }

        /// <summary>
        /// Retourne la vue associé à l'objet sélectionné.
        /// </summary>
        /// <param name="Index"></param>
        /// <param name="Marque"></param>
        /// <returns></returns>
        public eVue Vue(int Index, int Marque = -1)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            if (NbObjetsSelectionnes() == 0)
                return null;

            eVue pVue = new eVue();

            if (pVue.Init(_SwGestDeSelection.GetSelectedObjectsDrawingView2(Index, Marque), _Modele))
                return pVue;

            return null;
        }

        /// <summary>
        /// Fonction interne, renvoi la liste des objets sélectionnés.
        /// </summary>
        /// <param name="TypeObjet"></param>
        /// <param name="Marque"></param>
        /// <returns></returns>
        internal List<dynamic> ListListeDesObjetsSelectionnes(swSelectType_e TypeObjet, int Marque = -1)
        {
            List<dynamic> pListeObjets = new List<dynamic>();

            if (NbObjetsSelectionnes() > 0)
            {
                for (int i = 1; i <= _SwGestDeSelection.GetSelectedObjectCount2(-1); i++)
                {
                    if (this.TypeObjet(i, Marque) == TypeObjet)
                    {
                        dynamic pObjet = Objet(i, Marque);
                        if (pObjet != null)
                            pListeObjets.Add(pObjet);
                    }
                }
            }

            return pListeObjets;
        }

        /// <summary>
        /// Renvoi la liste des objets sélectionnés.
        /// </summary>
        /// <param name="Marque"></param>
        /// <returns></returns>
        public ArrayList ListeDesObjetsSelectionnes(swSelectType_e TypeObjet, int Marque = -1)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            List<dynamic> pListeObjets = ListListeDesObjetsSelectionnes(TypeObjet,Marque);
            ArrayList pArrayObjets = new ArrayList();

            if (pListeObjets.Count > 0)
                pArrayObjets = new ArrayList(pListeObjets);

            return pArrayObjets;
        }

        /// <summary>
        /// Fonction interne, renvoi la liste des composants sélectionnés.
        /// </summary>
        /// <param name="NomComposant"></param>
        /// <param name="Marque"></param>
        /// <returns></returns>
        internal List<eComposant> ListListeDesComposantsSelectionnes(String NomComposant = "", int Marque = -1)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            // Liste à renvoyer
            List<eComposant> pListeComposants = new List<eComposant>();

            for (int i = 1; i <= _SwGestDeSelection.GetSelectedObjectCount2(-1); i++)
            {

                eComposant pComposant = Composant(i, Marque);

                if ((pComposant != null) && pComposant.EstInitialise)
                    pListeComposants.Add(pComposant);
            }

            return pListeComposants;
        }

        /// <summary>
        /// Renvoi la liste des composants sélectionnés dans l'assemblage.
        /// </summary>
        /// <param name="NomComposant"></param>
        /// <param name="Marque"></param>
        /// <returns></returns>
        public ArrayList ListeDesComposantsSelectionnes(String NomComposant = "", int Marque = -1)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            List<eComposant> pListeComps = ListListeDesComposantsSelectionnes(NomComposant, Marque);
            ArrayList pArrayComps = new ArrayList();

            if (pListeComps.Count > 0)
                pArrayComps = new ArrayList(pListeComps);

            return pArrayComps;
        }

        /// <summary>
        /// Fonction interne, renvoi la liste des vues sélectionnées.
        /// </summary>
        /// <param name="Marque"></param>
        /// <returns></returns>
        internal List<eVue> ListListeDesVuesSelectionnes(int Marque = -1)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            // Liste à renvoyer
            List<eVue> pListeVues = new List<eVue>();

            if (_Modele.TypeDuModele == TypeFichier_e.cDessin)
            {

                for (int i = 1; i <= _SwGestDeSelection.GetSelectedObjectCount2(-1); i++)
                {
                    eVue pVue = Vue(i, Marque);

                    if ((pVue != null) && pVue.EstInitialise)
                        pListeVues.Add(pVue);
                }
            }

            return pListeVues;
        }

        /// <summary>
        /// Renvoi la liste des vues sélectionnées dans le dessin.
        /// </summary>
        /// <param name="Marque"></param>
        /// <returns></returns>
        public ArrayList ListeDesVuesSelectionnees(int Marque = -1)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            List<eVue> pListeVues = ListListeDesVuesSelectionnes(Marque);
            ArrayList pArrayVues = new ArrayList();

            if (pListeVues.Count > 0)
                pArrayVues = new ArrayList(pListeVues);

            return pArrayVues;
        }

        #endregion
    }
}
