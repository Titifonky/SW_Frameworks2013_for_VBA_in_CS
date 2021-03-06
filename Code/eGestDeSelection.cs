﻿using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections;
using System.Runtime.InteropServices;

namespace Framework
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
        dynamic Objet(int Index, int Marque = -1, Boolean RenvoyerObjet = false);
        eComposant Composant(int Index, int Marque = -1);
        eVue Vue(int Index, int Marque = -1);
        ArrayList ListeDesObjetsSelectionnes(swSelectType_e TypeObjet, int Marque = -1);
        ArrayList ListeDesComposantsSelectionnes(TypeFichier_e TypeDeFichier, String NomComposant = "", int Marque = -1);
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

        private static readonly String cNOMCLASSE = typeof(eGestDeSelection).Name;

        private Boolean _EstInitialise = false;

        private eModele _Modele = null;
        private SelectionMgr _SwGestDeSelection = null;

        #endregion

        #region "Constructeur\Destructeur"

        public eGestDeSelection() { }

        #endregion

        #region "Propriétés"

        /// <summary>
        /// Renvoi l'objet SelectionMgr.
        /// </summary>
        public SelectionMgr SwGestDeSelection { get { Log.Methode(cNOMCLASSE); return _SwGestDeSelection; } }

        /// <summary>
        /// Retourne le parent ExtModele.
        /// </summary>
        public eModele Modele { get { Log.Methode(cNOMCLASSE); return _Modele; } }

        /// <summary>
        /// Fonction interne
        /// Test l'initialisation de l'objet GestDeSelection
        /// </summary>
        internal Boolean EstInitialise { get { Log.Methode(cNOMCLASSE); return _EstInitialise; } }

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
            Log.Methode(cNOMCLASSE);

            if ((SwGestionnaire != null) && (Modele != null) && Modele.EstInitialise)
            {
                _SwGestDeSelection = SwGestionnaire;
                _Modele = Modele;
                _EstInitialise = true;
            }
            else
            {
                Log.Message("!!!!! Erreur d'initialisation");
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
            Log.Methode(cNOMCLASSE);
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
            Log.Methode(cNOMCLASSE);

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
            Log.Methode(cNOMCLASSE);

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
            Log.Methode(cNOMCLASSE);

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
        public dynamic Objet(int Index, int Marque = -1, Boolean RenvoyerObjet = false)
        {
            Log.Methode(cNOMCLASSE);

            if (NbObjetsSelectionnes() == 0)
                return null;

            eModele pModele = _Modele;

            if (_Modele.TypeDuModele != TypeFichier_e.cDessin)
                pModele = Composant(Index, Marque).Modele;

            dynamic pSwObjet = _SwGestDeSelection.GetSelectedObject6(Index, Marque);
            swSelectType_e pType = TypeObjet(Index, Marque);

            if ((pModele != null) && pModele.EstInitialise && !RenvoyerObjet)
            {

                switch (pType)
                {
                    case swSelectType_e.swSelCOMPONENTS:
                        Component2 pSwComposant = pSwObjet;
                        eComposant pComposant = new eComposant();
                        if (pComposant.Init(pSwComposant, pModele))
                        {
                            Modele.Composant = pComposant;
                            return pComposant;
                        }
                        break;

                    case swSelectType_e.swSelCONFIGURATIONS:
                        Configuration pSwConfiguration = pSwObjet;
                        eConfiguration pConfiguration = new eConfiguration();
                        if (pConfiguration.Init(pSwConfiguration, pModele))
                            return pConfiguration;
                        break;

                    case swSelectType_e.swSelDRAWINGVIEWS:
                        View pSwVue = pSwObjet;
                        eVue pVue = new eVue();
                        if (pVue.Init(pSwVue, pModele))
                            return pVue;
                        break;

                    case swSelectType_e.swSelSHEETS:
                        Sheet pSwFeuille = pSwObjet;
                        eFeuille pFeuille = new eFeuille();
                        if (pFeuille.Init(pSwFeuille, pModele))
                            return pFeuille;
                        break;

                    case swSelectType_e.swSelSOLIDBODIES:
                        Body2 pSwCorps = pSwObjet;
                        eCorps pCorps = new eCorps();
                        if (pCorps.Init(pSwCorps, pModele))
                            return pCorps;
                        break;

                    case swSelectType_e.swSelDATUMPLANES:
                    case swSelectType_e.swSelDATUMAXES:
                    case swSelectType_e.swSelDATUMPOINTS:
                    case swSelectType_e.swSelATTRIBUTES:
                    case swSelectType_e.swSelSKETCHES:
                    case swSelectType_e.swSelSECTIONLINES:
                    case swSelectType_e.swSelDETAILCIRCLES:
                    case swSelectType_e.swSelMATES:
                    case swSelectType_e.swSelBODYFEATURES:
                    case swSelectType_e.swSelREFCURVES:
                    case swSelectType_e.swSelREFERENCECURVES:
                    case swSelectType_e.swSelREFSILHOUETTE:
                    case swSelectType_e.swSelCAMERAS:
                    case swSelectType_e.swSelSWIFTANNOTATIONS:
                    case swSelectType_e.swSelSWIFTFEATURES:
                    case swSelectType_e.swSelCTHREADS:
                        eFonction pFonction = new eFonction();
                        if (pFonction.Init(pSwObjet, pModele))
                            return pFonction;
                        break;

                    default:
                        eObjet pObjet = new eObjet();

                        eModele pInitModele;
                        if ((pModele != null) && pModele.EstInitialise)
                            pInitModele = pModele;
                        else
                            pInitModele = _Modele;

                        if (pObjet.Init(pInitModele, pSwObjet, pType))
                            return pObjet;
                        break;

                }
            }
            else if (RenvoyerObjet)
            {
                eObjet pObjet = new eObjet();

                eModele pInitModele = _Modele;
                if ((pModele != null) && pModele.EstInitialise)
                    pInitModele = pModele;

                pObjet.Init(pInitModele, pSwObjet, pType);

                if (pObjet.EstInitialise)
                    return pObjet;
            }

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
            Log.Methode(cNOMCLASSE);

            if (NbObjetsSelectionnes() == 0)
                return null;

            Component2 pSwComposant = _SwGestDeSelection.GetSelectedObjectsComponent4(Index, Marque);
            Log.Message("OKKKKKK");
            // Si le composant racine est sélectionné et que l'on est dans un assemblage, rien n'est renvoyé.
            // Donc on le récupère.
            if ((pSwComposant == null) && (Modele.TypeDuModele == TypeFichier_e.cAssemblage))
            {
                Log.Message("OKKKKKK222222222222222");
                pSwComposant = _SwGestDeSelection.GetSelectedObject6(Index, Marque);
                Log.Message("OKKKKKK3333333333333333");
            }
            else if ((pSwComposant == null) && (Modele.TypeDuModele == TypeFichier_e.cPiece))
                pSwComposant = Modele.Composant.SwComposant;
            // Si c'est un dessin, pas de composant
            else if ((pSwComposant == null) && (Modele.TypeDuModele == TypeFichier_e.cDessin))
                return null;

            if (pSwComposant == null)
                Log.Message(" ========================= Erreur de composant");

            // Pour intitialiser le composant correctement il faut un peu de bidouille
            // sinon on à le droit à une belle reference circulaire
            // Donc d'abord, on recherche le modele du SwComposant
            Log.Message(pSwComposant.GetPathName());
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
            Log.Methode(cNOMCLASSE);

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
        public ArrayList ListeDesObjetsSelectionnes(swSelectType_e TypeObjet, int Marque = -1)
        {
            Log.Methode(cNOMCLASSE);

            ArrayList pListeObjets = new ArrayList();

            if (NbObjetsSelectionnes() > 0)
            {
                for (int i = 1; i <= _SwGestDeSelection.GetSelectedObjectCount2(Marque); i++)
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
        /// Fonction interne, renvoi la liste des composants sélectionnés.
        /// </summary>
        /// <param name="NomComposant"></param>
        /// <param name="Marque"></param>
        /// <returns></returns>
        public ArrayList ListeDesComposantsSelectionnes(TypeFichier_e TypeDeFichier, String NomComposant = "", int Marque = -1)
        {
            Log.Methode(cNOMCLASSE);

            // Liste à renvoyer
            ArrayList pListeComposants = new ArrayList();

            for (int i = 1; i <= _SwGestDeSelection.GetSelectedObjectCount2(Marque); i++)
            {
                eComposant pComposant = Composant(i, Marque);

                if ((pComposant != null) && pComposant.EstInitialise && TypeDeFichier.HasFlag(pComposant.TypeDuModele))
                    pListeComposants.Add(pComposant);
            }

            return pListeComposants;
        }

        /// <summary>
        /// Fonction interne, renvoi la liste des vues sélectionnées.
        /// </summary>
        /// <param name="Marque"></param>
        /// <returns></returns>
        public ArrayList ListeDesVuesSelectionnees(int Marque = -1)
        {
            Log.Methode(cNOMCLASSE);

            // Liste à renvoyer
            ArrayList pListeVues = new ArrayList();

            if (_Modele.TypeDuModele == TypeFichier_e.cDessin)
            {

                for (int i = 1; i <= _SwGestDeSelection.GetSelectedObjectCount2(Marque); i++)
                {
                    eVue pVue = Vue(i, Marque);

                    if ((pVue != null) && pVue.EstInitialise)
                        pListeVues.Add(pVue);
                }
            }

            return pListeVues;
        }

        #endregion
    }
}
