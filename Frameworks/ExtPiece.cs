using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using SolidWorks.Interop.sldworks;

/////////////////////////// Implementation terminée ///////////////////////////

namespace Framework_SW2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("57E51812-5820-11E2-82D5-34046188709B")]
    public interface IExtPiece
    {
        PartDoc SwPiece { get; }
        ExtModele Modele { get; }
        Boolean Contient(TypeCorps_e T);
        ArrayList ListeDesDossiers(TypeCorps_e TypeDeCorps = TypeCorps_e.cTousLesTypesDeCorps, Boolean PrendreEnCompteExclus = false);
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("5E46FC3E-5820-11E2-86E5-38046188709B")]
    [ProgId("Frameworks.ExtPiece")]
    public class ExtPiece : IExtPiece
    {
        #region "Variables locales"
        private Debug _Debug = Debug.Instance;
        private Boolean _EstInitialise = false;

        private ExtModele _Modele;
        private PartDoc _SwPiece;

        #endregion

        #region "Constructeur\Destructeur"

        public ExtPiece() { }

        #endregion

        #region "Propriétés"

        public PartDoc SwPiece { get { return _SwPiece; } }

        public ExtModele Modele { get { return _Modele; } }

        internal Boolean EstInitialise { get { return _EstInitialise; } }

        #endregion

        #region "Méthodes"

        internal Boolean Init(ExtModele Modele)
        {
            _MethodBase Methode = System.Reflection.MethodBase.GetCurrentMethod();

            if ((Modele != null) && Modele.EstInitialise && (Modele.TypeDuModele == TypeFichier_e.cPiece))
            {
                _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name + " : " + Modele.Chemin);

                _Modele = Modele;
                _SwPiece = Modele.SwModele as PartDoc;
                _EstInitialise = true;
            }
            else
            {
                _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name + " : Erreur d'initialisation");
            }

            return _EstInitialise;
        }

        internal Feature ListeDesPiecesSoudees()
        {
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
            if (Convert.ToBoolean(T & TypeCorps_e.cTole))
            {
                foreach (ExtFonction Fonction in _Modele.ListListeDesFonctions())
                {
                    if ((Fonction.TypeDeLaFonction == "SMBaseFlange") || (Fonction.TypeDeLaFonction == "SolidToSheetMetal"))
                        return true;
                }

                T = (TypeCorps_e)( T - TypeCorps_e.cTole);
            }

            if (Convert.ToBoolean(T & TypeCorps_e.cProfil))
            {
                foreach (ExtFonction Fonction in _Modele.ListListeDesFonctions())
                {
                    if (Fonction.TypeDeLaFonction == "WeldMemberFeat")
                        return true;
                }

                T = (TypeCorps_e)(T - TypeCorps_e.cProfil);
            }

            if (Convert.ToBoolean(T & TypeCorps_e.cAutre))
            {
                if (ListListeDesDossiers(TypeCorps_e.cAutre, false).Count > 0)
                    return true;
            }

            return false;
        }

        internal List<ExtDossier> ListListeDesDossiers(TypeCorps_e TypeDeCorps = TypeCorps_e.cTousLesTypesDeCorps, Boolean PrendreEnCompteExclus = false)
        {
            _MethodBase Methode = System.Reflection.MethodBase.GetCurrentMethod();
            _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name);

            List<ExtDossier> Liste = new List<ExtDossier>();

            Feature pFonction = ListeDesPiecesSoudees();

            // S'il n'y a pas de liste, on arrete là
            if (pFonction == null)
                return null;

            pFonction = pFonction.GetFirstSubFeature();

            while (pFonction != null)
            {
                if (pFonction.GetTypeName2() == "CutListFolder")
                {
                    BodyFolder pSwDossier = pFonction.GetSpecificFeature2();
                    ExtDossier Dossier = new ExtDossier();

                    if (Dossier.Init(pSwDossier, this) && Convert.ToBoolean(Dossier.TypeDeCorps | TypeDeCorps) && (!Dossier.EstExclu | PrendreEnCompteExclus))
                        Liste.Add(Dossier);

                    Dossier = null;
                }

                pFonction = pFonction.GetNextFeature();
            }

            return Liste;

        }

        public ArrayList ListeDesDossiers(TypeCorps_e TypeDeCorps = TypeCorps_e.cTousLesTypesDeCorps, Boolean PrendreEnCompteExclus = false)
        {
            _MethodBase Methode = System.Reflection.MethodBase.GetCurrentMethod();
            _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name);

            List<ExtDossier> pListeDossier = ListListeDesDossiers(TypeDeCorps, PrendreEnCompteExclus);
            ArrayList pArrayDossiers = new ArrayList();

            if (pListeDossier.Count > 0)
                pArrayDossiers = new ArrayList(pListeDossier);

            return pArrayDossiers;
        }

        #endregion

    }
}
