using System;
using System.Reflection;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace Framework
{

    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("3E082AF3-8F22-4687-A8AB-5DBE544FA5D8")]
    public interface IeParametreTolerie
    {
        ePiece Piece { get; }
        eTole Tole { get; }
        Double Epaisseur { get; set; }
        Double Rayon { get; set; }
        Double FacteurK { get; set; }
        Boolean EcraserEpaisseur { get; set; }
        Boolean EcraserRayon { get; set; }
        Boolean EcraserFacteurK { get; set; }
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("F699797C-AD86-4E2C-957E-2BE8EE033941")]
    [ProgId("Frameworks.eParametreTolerie")]
    public class eParametreTolerie : IeParametreTolerie
    {
        #region "Variables locales"

        private Boolean _EstInitialise = false;

        private ePiece _Piece = null;
        private eTole _Tole = null;
        private String _DimRayon = "D1";
        private String _DimK = "D2";

        #endregion

        #region "Constructeur\Destructeur"

        public eParametreTolerie() { }

        #endregion

        #region "Propriétés"

        /// <summary>
        /// Retourne le parent ExtPiece.
        /// </summary>
        public ePiece Piece { get { Debug.Print(MethodBase.GetCurrentMethod()); return _Piece; } }

        /// <summary>
        /// Retourne le parent ExtPiece.
        /// </summary>
        public eTole Tole { get { Debug.Print(MethodBase.GetCurrentMethod()); return _Tole; } }

        /// <summary>
        /// Retourne ou défini l'épaisseur de la tole
        /// </summary>
        public Double Epaisseur
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());

                SheetMetalFeatureData pParam = FonctionTolerie.SwFonction.GetDefinition();
                return Math.Round(pParam.Thickness * 1000, 5);
            }
            set
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                eFonction pFonctionTolerie = FonctionTolerie;
                Dimension pSwDim;

                if (_Tole != null)
                {
                    pSwDim = FonctionToleDeBase.SwFonction.Parameter("D7");
                    if (pSwDim != null)
                    {
                        pSwDim.SetSystemValue3(value * 0.001, (int)swInConfigurationOpts_e.swThisConfiguration, null);
                    }

                    pSwDim = pFonctionTolerie.SwFonction.Parameter("D7");
                    if (pSwDim != null)
                    {
                        pSwDim.SetSystemValue3(value * 0.001, (int)swInConfigurationOpts_e.swThisConfiguration, null);
                    }
                }

                pSwDim = pFonctionTolerie.SwFonction.Parameter("Epaisseur");
                if (pSwDim != null)
                {
                    pSwDim.SetSystemValue3(value * 0.001, (int)swInConfigurationOpts_e.swThisConfiguration, null);
                }
            }
        }

        /// <summary>
        /// Retourne ou défini le rayon intérieur de pliage
        /// </summary>
        public Double Rayon
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());

                BaseFlangeFeatureData pParam = FonctionTolerie.SwFonction.GetDefinition();
                return Math.Round(pParam.BendRadius * 1000, 5);
            }
            set
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                Dimension pSwDim = FonctionTolerie.SwFonction.Parameter(_DimRayon);
                if (pSwDim == null) return;
                pSwDim.SetSystemValue3(value * 0.001, (int)swInConfigurationOpts_e.swThisConfiguration, null).ToString();
            }
        }

        /// <summary>
        /// Retourne ou défini le facteur K pour le developpé
        /// </summary>
        public Double FacteurK
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());

                BaseFlangeFeatureData pParam = FonctionTolerie.SwFonction.GetDefinition();
                return Math.Round(pParam.KFactor, 5);
            }
            set
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                Dimension pSwDim = FonctionTolerie.SwFonction.Parameter(_DimK);
                if (pSwDim == null) return;
                pSwDim.SetSystemValue3(value, (int)swInConfigurationOpts_e.swThisConfiguration, null);
            }
        }

        public Boolean EcraserEpaisseur
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                if (_Tole == null)
                    return false;
                BaseFlangeFeatureData pParam = FonctionToleDeBase.SwFonction.GetDefinition();
                return pParam.OverrideThickness;
            }
            set
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                if (_Tole == null)
                    return;
                ModelDoc2 pSwModele = SwModele;
                Component2 pSwComposant = SwComposant;
                Feature pSwFonction = FonctionToleDeBase.SwFonction;
                BaseFlangeFeatureData pParam = pSwFonction.GetDefinition();
                pParam.AccessSelections(pSwModele, pSwComposant);
                pParam.OverrideThickness = value;
                pSwFonction.ModifyDefinition(pParam, pSwModele, pSwComposant);
                pParam.ReleaseSelectionAccess();
            }
        }

        public Boolean EcraserRayon
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                if (_Tole == null)
                    return false;
                BaseFlangeFeatureData pParam = FonctionToleDeBase.SwFonction.GetDefinition();
                return pParam.OverrideRadius;
            }
            set
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                if (_Tole == null)
                    return;
                ModelDoc2 pSwModele = SwModele;
                Component2 pSwComposant = SwComposant;
                Feature pSwFonction = FonctionToleDeBase.SwFonction;
                BaseFlangeFeatureData pParam = pSwFonction.GetDefinition();
                pParam.AccessSelections(pSwModele, pSwComposant);
                pParam.OverrideRadius = value;
                pSwFonction.ModifyDefinition(pParam, pSwModele, pSwComposant);
                pParam.ReleaseSelectionAccess();
            }
        }

        public Boolean EcraserFacteurK
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                if (_Tole == null)
                    return false;
                BaseFlangeFeatureData pParam = FonctionToleDeBase.SwFonction.GetDefinition();
                return pParam.OverrideKFactor;
            }
            set
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                if (_Tole == null)
                    return;
                ModelDoc2 pSwModele = SwModele;
                Component2 pSwComposant = SwComposant;
                Feature pSwFonction = FonctionToleDeBase.SwFonction;
                BaseFlangeFeatureData pParam = pSwFonction.GetDefinition();
                pParam.AccessSelections(pSwModele, pSwComposant);
                pParam.OverrideKFactor = value;
                pSwFonction.ModifyDefinition(pParam, pSwModele, pSwComposant);
                pParam.ReleaseSelectionAccess();
            }
        }

        /// <summary>
        /// Renvoi la fonction Tolerie
        /// </summary>
        /// <returns></returns>
        private eFonction FonctionTolerie
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());

                if (_Piece != null)
                {
                    return DossierTolerie();
                }
                else
                {
                    return _Tole.FonctionTolerie;
                }
            }
        }

        /// <summary>
        /// Renvoi la fonction Tolerie
        /// </summary>
        /// <returns></returns>
        private eFonction FonctionToleDeBase
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());

                if (_Tole != null)
                    return _Tole.FonctionToleDeBase;

                return null;
            }
        }

        private ModelDoc2 SwModele
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                if (_Tole.Corps.Piece.Modele.Equals(_Tole.Corps.Piece.Modele.SW.Modele()))
                    return _Tole.Corps.Piece.Modele.SwModele;

                return null;
            }
        }

        private Component2 SwComposant
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                if (SwModele == null)
                    return _Tole.Corps.Piece.Modele.Composant.SwComposant;

                return null;
            }
        }

        /// <summary>
        /// Fonction interne.
        /// Test l'initialisation de l'objet eTole.
        /// </summary>
        internal Boolean EstInitialise { get { Debug.Print(MethodBase.GetCurrentMethod()); return _EstInitialise; } }

        #endregion

        #region "Méthodes"

        /// <summary>
        /// Méthode interne.
        /// Initialiser
        /// </summary>
        /// <param name="SwCorps"></param>
        /// <param name="Tole"></param>
        /// <returns></returns>
        internal Boolean Init(eTole Tole)
        {
            Debug.Print(MethodBase.GetCurrentMethod());

            if ((Tole != null) && Tole.EstInitialise)
            {
                _Tole = Tole;
                _EstInitialise = true;
            }
            else
            {
                Debug.Print("!!!!! Erreur d'initialisation");
            }
            return _EstInitialise;
        }

        /// <summary>
        /// Méthode interne.
        /// Initialiser
        /// </summary>
        /// <param name="SwCorps"></param>
        /// <param name="Piece"></param>
        /// <returns></returns>
        internal Boolean Init(ePiece Piece)
        {
            Debug.Print(MethodBase.GetCurrentMethod());

            if ((Piece != null) && Piece.EstInitialise && (Piece.Contient(TypeCorps_e.cTole)))
            {
                _Piece = Piece;
                _EstInitialise = true;
            }
            else
            {
                Debug.Print("!!!!! Erreur d'initialisation");
            }
            return _EstInitialise;
        }

        /// <summary>
        /// Renvoi le dossier Tôlerie à partir du FeatureManager
        /// </summary>
        /// <returns></returns>
        private eFonction DossierTolerie()
        {
            Debug.Print(MethodBase.GetCurrentMethod());

            TreeControlItem pNoeud = Piece.Modele.GestDeFonction_NoeudRacine().GetFirstChild();

            while (pNoeud != null)
            {
                eFonction pFonction = new eFonction();
                if (pNoeud.ObjectType == (int)swTreeControlItemType_e.swFeatureManagerItem_Feature)
                {
                    if (pFonction.Init(pNoeud.Object, Piece.Modele)
                        && (pFonction.TypeDeLaFonction == "TemplateSheetMetal"))
                        return pFonction;
                }

                pNoeud = pNoeud.GetNext();
            }

            return null;
        }

        #endregion
    }

}
