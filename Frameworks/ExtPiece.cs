using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;

namespace Frameworks2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("57E51812-5820-11E2-82D5-34046188709B")]
    public interface IExtPiece
    {
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
        private PartDoc _swPiece;

        #endregion

        #region "Constructeur\Destructeur"

        public ExtPiece()
        {
        }

        #endregion

        #region "Propriétés"
        #endregion

        #region "Méthodes"

        internal ExtPiece Init(ExtModele Modele)
        {
            _MethodBase Methode = System.Reflection.MethodBase.GetCurrentMethod();

            if ((Modele != null) && (Modele.TypeDuModele == TypeFichier_e.cPiece))
            {
                
                _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name);
                
                _Modele = Modele;
                _swPiece = Modele.SwModele as PartDoc;
                _EstInitialise = true;
                return this;
            }

            _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name + " : Erreur d'initialisation");
            _EstInitialise = false;
            return null;
        }

        internal ExtPiece Init()
        {
            if (_EstInitialise)
                return this;
            else
                return null;
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

        internal List<ExtDossier> ListeDesDossiers(TypeCorps_e TypeDeCorps = TypeCorps_e.cTousLesTypesDeCorps, Boolean PrendreEnCompteExclus = false)
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

                    if ((Dossier.Init(pSwDossier, this) != null) && (Dossier.Est(TypeDeCorps)) && (!Dossier.EstExclu | PrendreEnCompteExclus))
                        Liste.Add(Dossier);
                }
            }

            return Liste;

        }

        //Public Function ListeDesDossiers(Optional TypeDeCorps As TypeCorps_e = cTousLesTypesDeCorps, Optional PrendreEnCompteExclus As Boolean = False) As Collection
        //    Dim Dossier     As ExtDossier
        //    Dim Fonction    As Feature
        //    Dim DossierSw   As BodyFolder
    
        //    Set ListeDesDossiers = New Collection
        //    Set Fonction = ListeDesPiecesSoudees.GetFirstSubFeature
    
        //    Do Until Fonction Is Nothing
    
        //        'Si c'est un dossier de pièces soudées
        //        If Fonction.GetTypeName2 = "CutListFolder" Then
            
        //            Set DossierSw = Fonction.GetSpecificFeature2
        //            Set Dossier = New ExtDossier
            
        //            If Dossier.SetDossier(DossierSw, Me) Then
        //                If Dossier.Est(TypeDeCorps) And (Dossier.Exclu Imp PrendreEnCompteExclus) Then
        //                    ListeDesDossiers.Add Dossier
        //                End If
        //            End If
        //            Set Dossier = Nothing
        //        End If
        //        Set Fonction = Fonction.GetNextSubFeature
        //    Loop
    
        //    Set Dossier = Nothing
        //    Set Fonction = Nothing
        //    Set DossierSw = Nothing
        
        //End Function

        #endregion

    }
}
