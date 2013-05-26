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
    [Guid("57E51812-5820-11E2-82D5-34046188709B")]
    public interface IePiece
    {
        PartDoc SwPiece { get; }
        eModele Modele { get; }
        Boolean Contient(TypeCorps_e T);
        ArrayList ListeDesCorps(TypeCorps_e TypeDeCorps = TypeCorps_e.cTousLesTypesDeCorps, Boolean PrendreEnCompteCache = false);
        ArrayList ListeDesDossiersDePiecesSoudees(TypeCorps_e TypeDeCorps = TypeCorps_e.cTousLesTypesDeCorps, Boolean PrendreEnCompteExclus = false);
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("5E46FC3E-5820-11E2-86E5-38046188709B")]
    [ProgId("Frameworks.ePiece")]
    public class ePiece : IePiece
    {
        #region "Variables locales"
        
        private Boolean _EstInitialise = false;

        private eModele _Modele = null;
        private PartDoc _SwPiece = null;

        #endregion

        #region "Constructeur\Destructeur"

        public ePiece() { }

        #endregion

        #region "Propriétés"

        /// <summary>
        /// Renvoi l'objet PartDoc.
        /// </summary>
        public PartDoc SwPiece { get { Debug.Info(MethodBase.GetCurrentMethod());  return _SwPiece; } }

        /// <summary>
        /// Renvoi l'objet ExtModele.
        /// </summary>
        public eModele Modele { get { Debug.Info(MethodBase.GetCurrentMethod());  return _Modele; } }

        /// <summary>
        /// Renvoi la valeur de l'initialisation.
        /// </summary>
        internal Boolean EstInitialise { get { Debug.Info(MethodBase.GetCurrentMethod());  return _EstInitialise; } }

        #endregion

        #region "Méthodes"

        /// <summary>
        /// Méthode interne.
        /// Initialiser l'objet pièce.
        /// </summary>
        /// <param name="Modele"></param>
        /// <returns></returns>
        internal Boolean Init(eModele Modele)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            if ((Modele != null) && Modele.EstInitialise && (Modele.TypeDuModele == TypeFichier_e.cPiece))
            {
                Debug.Info(Modele.FichierSw.Chemin);

                _Modele = Modele;
                _SwPiece = Modele.SwModele as PartDoc;
                _EstInitialise = true;
            }
            else
            {
                Debug.Info("!!!!! Erreur d'initialisation");
            }

            return _EstInitialise;
        }

        /// <summary>
        /// Renvoi la fonction du dossier contenant les corps.
        /// </summary>
        /// <returns></returns>
        internal Feature DossierDesCorps()
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            Feature pFonctionPiecesSoudees = _SwPiece.FirstFeature();

            while (pFonctionPiecesSoudees != null)
            {
                if (pFonctionPiecesSoudees.GetTypeName2() == "SolidBodyFolder")
                {
                    BodyFolder pDossier = pFonctionPiecesSoudees.GetSpecificFeature2();
                    pDossier.UpdateCutList();
                    return pFonctionPiecesSoudees;
                }

                pFonctionPiecesSoudees = pFonctionPiecesSoudees.GetNextFeature();
            }

            return null;
        }
        
        /// <summary>
        /// Renvoi Vrai si la piece contient des corps du type T
        /// A tester. Peut dans certain cas renvoyer un resultat erroné. ex :
        /// Si un corps de tolerie ou un profil a été créé puis supprimé,
        /// la fonction existe mais plus le corps. Pb également dans le cas de corps combiné
        /// </summary>
        /// <param name="T"></param>
        /// <returns></returns>
        public Boolean Contient(TypeCorps_e T)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            if (Convert.ToBoolean(T & TypeCorps_e.cTole))
            {
                foreach (eFonction Fonction in _Modele.ListListeDesFonctions())
                {
                    if ((Fonction.TypeDeLaFonction == "SMBaseFlange")
                        || (Fonction.TypeDeLaFonction == "SolidToSheetMetal")
                        || (Fonction.TypeDeLaFonction == "SheetMetal")
                        || (Fonction.TypeDeLaFonction == "FlatPattern"))
                        return true;
                }

                T = (TypeCorps_e)( T - TypeCorps_e.cTole);
            }

            if (Convert.ToBoolean(T & TypeCorps_e.cBarre))
            {
                foreach (eFonction Fonction in _Modele.ListListeDesFonctions())
                {
                    if (Fonction.TypeDeLaFonction == "WeldMemberFeat")
                        return true;
                }

                T = (TypeCorps_e)(T - TypeCorps_e.cBarre);
            }

            if (Convert.ToBoolean(T & TypeCorps_e.cAutre))
            {
                if (ListListeDesDossiersDePiecesSoudees(TypeCorps_e.cAutre, false).Count > 0)
                    return true;
            }

            return false;
        }

        /// <summary>
        /// Méthode interne.
        /// Renvoi la liste des corps de la pièces filtrée par les arguments.
        /// </summary>
        /// <param name="TypeDeCorps"></param>
        /// <param name="PrendreEnCompteCache"></param>
        /// <returns></returns>
        internal List<eCorps> ListListeDesCorps(TypeCorps_e TypeDeCorps = TypeCorps_e.cTousLesTypesDeCorps, Boolean PrendreEnCompteCache = false)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            List<eCorps> Liste = new List<eCorps>();

            Object[] TableauDesCorps = _SwPiece.GetBodies2((int)swBodyType_e.swAllBodies, PrendreEnCompteCache);

            if (TableauDesCorps.Length > 0)
            {
                foreach (Object ObjetCorps in TableauDesCorps)
                {
                    Body2 pSwCorps = (Body2)ObjetCorps;
                    eCorps pCorps = new eCorps();
                    if (pCorps.Init(pSwCorps, this) && Convert.ToBoolean(pCorps.TypeDeCorps & TypeDeCorps))
                        {
                            Liste.Add(pCorps);
                        }
                }
            }

            return Liste;
        }

        /// <summary>
        /// Renvoi la liste des corps de la pièces filtrée par les arguments.
        /// </summary>
        /// <param name="TypeDeCorps"></param>
        /// <param name="PrendreEnCompteCache"></param>
        /// <returns></returns>
        public ArrayList ListeDesCorps(TypeCorps_e TypeDeCorps = TypeCorps_e.cTousLesTypesDeCorps, Boolean PrendreEnCompteCache = false)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            List<eCorps> pListeCorps = ListListeDesCorps(TypeDeCorps, PrendreEnCompteCache);
            ArrayList pArrayCorps = new ArrayList();

            if (pListeCorps.Count > 0)
                pArrayCorps = new ArrayList(pListeCorps);

            return pArrayCorps;
        }

        /// <summary>
        /// Méthode interne
        /// Renvoi la liste des dossiers de pièces soudées de la pièce filtrée par les arguments.
        /// </summary>
        /// <param name="TypeDeCorps"></param>
        /// <param name="PrendreEnCompteExclus"></param>
        /// <returns></returns>
        internal List<eDossier> ListListeDesDossiersDePiecesSoudees(TypeCorps_e TypeDeCorps = TypeCorps_e.cTousLesTypesDeCorps, Boolean PrendreEnCompteExclus = false)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            List<eDossier> Liste = new List<eDossier>();

            Feature pFonction = DossierDesCorps();

            // S'il n'y a pas de liste, on arrete là
            if (pFonction == null)
                return null;

            pFonction = pFonction.GetFirstSubFeature();

            while (pFonction != null)
            {
                if (pFonction.GetTypeName2() == "CutListFolder")
                {
                    BodyFolder pSwDossier = pFonction.GetSpecificFeature2();
                    eDossier Dossier = new eDossier();

                    if (Dossier.Init(pSwDossier, this) && Convert.ToBoolean(Dossier.TypeDeCorps & TypeDeCorps) && (!Dossier.EstExclu | PrendreEnCompteExclus))
                        Liste.Add(Dossier);

                    Dossier = null;
                }

                pFonction = pFonction.GetNextFeature();
            }

            return Liste;

        }

        /// <summary>
        /// Renvoi la liste des dossiers de pièces soudées de la pièce filtrée par les arguments.
        /// </summary>
        /// <param name="TypeDeCorps"></param>
        /// <param name="PrendreEnCompteExclus"></param>
        /// <returns></returns>
        public ArrayList ListeDesDossiersDePiecesSoudees(TypeCorps_e TypeDeCorps = TypeCorps_e.cTousLesTypesDeCorps, Boolean PrendreEnCompteExclus = false)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            List<eDossier> pListeDossier = ListListeDesDossiersDePiecesSoudees(TypeDeCorps, PrendreEnCompteExclus);
            ArrayList pArrayDossiers = new ArrayList();

            if (pListeDossier.Count > 0)
                pArrayDossiers = new ArrayList(pListeDossier);

            return pArrayDossiers;
        }

        #endregion

    }
}
