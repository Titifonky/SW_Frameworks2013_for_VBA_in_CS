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
        Boolean MethodeModification { set; }
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
        private Boolean _Methode = true;

        #endregion

        #region "Constructeur\Destructeur"

        public eParametreTolerie() { }

        #endregion

        #region "Propriétés"

        public Boolean MethodeModification { set { _Methode = value; } }

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
                Dimension pSwDim;
                if (_Tole != null)
                {
                    if (_Methode)
                    {
                        pSwDim = FonctionToleDeBase.SwFonction.Parameter("D7");
                        if (pSwDim != null)
                        {
                            pSwDim.SetSystemValue3(value * 0.001, (int)swInConfigurationOpts_e.swThisConfiguration, null);
                        }

                        pSwDim = FonctionTolerie.SwFonction.Parameter("D7");
                        if (pSwDim != null)
                        {
                            pSwDim.SetSystemValue3(value * 0.001, (int)swInConfigurationOpts_e.swThisConfiguration, null);
                        }
                    }
                    else
                    {
                        Feature F = FonctionToleDeBase.SwFonction;
                        BaseFlangeFeatureData pParam = F.GetDefinition();
                        pParam.AccessSelections(SwModele, SwComposant);
                        pParam.Thickness = value * 0.001;
                        F.ModifyDefinition(pParam, SwModele, SwComposant);
                    }
                }
                else
                {
                    pSwDim = FonctionTolerie.SwFonction.Parameter("Epaisseur");
                    if (pSwDim != null)
                    {
                        pSwDim.SetSystemValue3(value * 0.001, (int)swInConfigurationOpts_e.swThisConfiguration, null);
                    }
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

                SheetMetalFeatureData pParam = FonctionTolerie.SwFonction.GetDefinition();
                return Math.Round(pParam.BendRadius * 1000, 5);
            }
            set
            {
                Debug.Print(MethodBase.GetCurrentMethod());

                if (_Methode)
                {
                    String pDimRayon = "D1";
                    Dimension pSwDim = FonctionTolerie.SwFonction.Parameter(pDimRayon);
                    if (pSwDim == null) return;
                    pSwDim.SetSystemValue3(value * 0.001, (int)swInConfigurationOpts_e.swThisConfiguration, null).ToString();
                }
                else
                {
                    Feature F = FonctionTolerie.SwFonction;
                    SheetMetalFeatureData pParam = F.GetDefinition();
                    pParam.AccessSelections(SwModele, SwComposant);
                    pParam.BendRadius = value * 0.001;
                    F.ModifyDefinition(pParam, SwModele, SwComposant);
                }
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

                SheetMetalFeatureData pParam = FonctionTolerie.SwFonction.GetDefinition();
                return Math.Round(pParam.KFactor, 5);
            }
            set
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                if (_Methode)
                {
                    String pDimK = "D2";
                    Dimension pSwDim = FonctionTolerie.SwFonction.Parameter(pDimK);
                    if (pSwDim == null) return;
                    pSwDim.SetSystemValue3(value, (int)swInConfigurationOpts_e.swThisConfiguration, null);
                }
                else
                {
                    Feature F = FonctionTolerie.SwFonction;
                    SheetMetalFeatureData pParam = F.GetDefinition();
                    pParam.AccessSelections(SwModele, SwComposant);
                    pParam.KFactor = value;
                    F.ModifyDefinition(pParam, SwModele, SwComposant);
                }
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

                if (_Tole != null)
                {
                    return _Tole.FonctionTolerie;
                }
                else
                {
                    return DossierTolerie();
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

                if (Piece.Modele.Equals(Piece.Modele.SW.Modele()))
                    return Piece.Modele.SwModele;

                return null;
            }
        }

        private Component2 SwComposant
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                if (SwModele == null)
                    return Piece.Modele.Composant.SwComposant;

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
                _Piece = _Tole.Corps.Piece;
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
