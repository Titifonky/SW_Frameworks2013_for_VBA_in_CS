using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using SolidWorks.Interop.sldworks;

namespace Framework_SW2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("57E51812-5820-11E2-82D5-34046188709B")]
    public interface IExtPiece
    {
        PartDoc SwPiece { get; }
        ExtModele Modele { get; }
        ArrayList ListeDesDossiers(TypeCorps_e TypeDeCorps = TypeCorps_e.cTousLesTypesDeCorps, Boolean PrendreEnCompteExclus = false);
        ArrayList ListeDesFonctions(String NomARechercher = "", Boolean AvecLesSousFonctions = false);
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
        private PartDoc _swPiece;

        #endregion

        #region "Constructeur\Destructeur"

        public ExtPiece() { }

        #endregion

        #region "Propriétés"

        public PartDoc SwPiece { get { return _swPiece; } }

        public ExtModele Modele { get { return _Modele; } }

        internal Boolean EstInitialise { get { return _EstInitialise; } }

        #endregion

        #region "Méthodes"

        internal Boolean Init(ExtModele Modele)
        {
            _MethodBase Methode = System.Reflection.MethodBase.GetCurrentMethod();

            if ((Modele != null) && Modele.EstInitialise && (Modele.TypeDuModele == TypeFichier_e.cPiece))
            {
                _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name);

                _Modele = Modele;
                _swPiece = Modele.SwModele as PartDoc;
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
            Feature pFonctionPiecesSoudees = _Modele.SwModele.FirstFeature();

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

        internal List<ExtDossier> ListListeDesDossiers(TypeCorps_e TypeDeCorps = TypeCorps_e.cTousLesTypesDeCorps, Boolean PrendreEnCompteExclus = false)
        {
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

                    if (Dossier.Init(pSwDossier, this) && (Dossier.TypeDeCorps == TypeDeCorps) && (!Dossier.EstExclu | PrendreEnCompteExclus))
                        Liste.Add(Dossier);

                    Dossier = null;
                }

                pFonction = pFonction.GetNextFeature();
            }

            return Liste;

        }

        internal List<ExtFonction> ListListeDesFonctions(String NomARechercher = "", Boolean AvecLesSousFonctions = false)
        {
            List<ExtFonction> Liste = new List<ExtFonction>();

            Feature pSwFonction = _Modele.SwModele.FirstFeature();

            while (pSwFonction != null)
            {
                ExtFonction pFonction = new ExtFonction();

                if ((Regex.IsMatch(pSwFonction.Name, NomARechercher)) && pFonction.Init(pSwFonction, this))
                    Liste.Add(pFonction);

                pFonction = null;

                pSwFonction = pSwFonction.GetNextFeature();
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
