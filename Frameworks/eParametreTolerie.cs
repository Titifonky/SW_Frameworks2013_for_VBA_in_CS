using System;
using System.Reflection;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;

namespace Framework_SW2013
{
#if SW2013
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("3E082AF3-8F22-4687-A8AB-5DBE544FA5D8")]
    public interface IeParametreTolerie
    {
        ePiece Piece { get; }
        eCorps Corps { get; }
        Double Epaisseur { get; set; }
        Double Rayon { get; set; }
        Double FacteurK { get; set; }
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("F699797C-AD86-4E2C-957E-2BE8EE033941")]
    [ProgId("Frameworks.eParametreTolerie")]
    public class eParametreTolerie : IeParametreTolerie
    {
#region "Variables locales"

        private Boolean _EstInitialise = false;

        private ePiece _Piece = null;
        private eCorps _Corps = null;

#endregion

#region "Constructeur\Destructeur"

        public eParametreTolerie() { }

#endregion

#region "Propriétés"

        /// <summary>
        /// Retourne le parent ExtPiece.
        /// </summary>
        public ePiece Piece { get { Debug.Info(MethodBase.GetCurrentMethod()); return _Piece; } }

        /// <summary>
        /// Retourne le parent ExtPiece.
        /// </summary>
        public eCorps Corps { get { Debug.Info(MethodBase.GetCurrentMethod()); return _Corps; } }

        private ModelDoc2 SwModele
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());

                eModele pModele;
                if (_Corps == null)
                    pModele = _Piece.Modele;
                else
                    pModele = _Corps.Piece.Modele;

                if (pModele.EstActif)
                    return pModele.SwModele;

                return null;
            }
        }

        private Component2 SwComposant
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());

                eComposant pComposant;
                if (_Corps == null)
                    pComposant = _Piece.Modele.Composant;
                else
                    pComposant = _Corps.Piece.Modele.Composant;

                if (SwModele == null)
                    return pComposant.SwComposant;

                return null;
            }
        }

        /// <summary>
        /// Retourne ou défini l'épaisseur de la tole
        /// Ne fonctionne pas !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
        /// </summary>
        public Double Epaisseur
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                SheetMetalFeatureData pParam = FonctionTolerie.SwFonction.GetDefinition();
                return pParam.Thickness * 1000;
            }
            set
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                ModelDoc2 pSwModele = SwModele;
                Component2 pSwComposant = SwComposant;
                Feature pSwFonction = FonctionTolerie.SwFonction;
                SheetMetalFeatureData pParam = pSwFonction.GetDefinition();
                pParam.AccessSelections(pSwModele, pSwComposant);
                pParam.Thickness = value * 0.001;
                pSwFonction.ModifyDefinition(pParam, pSwModele, pSwComposant);
                pParam.ReleaseSelectionAccess();
            }
        }

        /// <summary>
        /// Retourne ou défini le rayon intérieur de pliage
        /// </summary>
        public Double Rayon
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                SheetMetalFeatureData pParam = FonctionTolerie.SwFonction.GetDefinition();
                return pParam.BendRadius * 1000;
            }
            set
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                ModelDoc2 pSwModele = SwModele;
                Component2 pSwComposant = SwComposant;
                Feature pSwFonction = FonctionTolerie.SwFonction;
                SheetMetalFeatureData pParam = pSwFonction.GetDefinition();
                pParam.AccessSelections(pSwModele, pSwComposant);
                pParam.BendRadius = value * 0.001;
                pSwFonction.ModifyDefinition(pParam, pSwModele, pSwComposant);
                pParam.ReleaseSelectionAccess();
            }
        }

        /// <summary>
        /// Retourne ou défini le facteur K pour le developpé
        /// </summary>
        public Double FacteurK
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                SheetMetalFeatureData pParam = FonctionTolerie.SwFonction.GetDefinition();
                return pParam.GetCustomBendAllowance().KFactor;
            }
            set
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                ModelDoc2 pSwModele = SwModele;
                Component2 pSwComposant = SwComposant;
                Feature pSwFonction = FonctionTolerie.SwFonction;
                SheetMetalFeatureData pParam = pSwFonction.GetDefinition();
                CustomBendAllowance pParamPli = pParam.GetCustomBendAllowance();
                pParam.AccessSelections(pSwModele, pSwComposant);
                pParamPli.KFactor = value;
                pParam.SetCustomBendAllowance(pParamPli);
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
                Debug.Info(MethodBase.GetCurrentMethod());

                swFeatureType_e TypeFonc = new swFeatureType_e();

                foreach (eFonction pFonc in _Piece.Modele.ListListeDesFonctions("^Tôlerie$"))
                {
                    return pFonc;
                }

                return null;
            }
        }

        /// <summary>
        /// Fonction interne.
        /// Test l'initialisation de l'objet eTole.
        /// </summary>
        internal Boolean EstInitialise { get { Debug.Info(MethodBase.GetCurrentMethod()); return _EstInitialise; } }

#endregion

#region "Méthodes"

        /// <summary>
        /// Méthode interne.
        /// Initialiser
        /// </summary>
        /// <param name="SwCorps"></param>
        /// <param name="Corps"></param>
        /// <returns></returns>
        internal Boolean Init(eCorps Corps)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            if ((Corps != null) && Corps.EstInitialise && (Corps.TypeDeCorps == TypeCorps_e.cTole))
            {
                _Corps = Corps;
                _EstInitialise = true;
            }
            else
            {
                Debug.Info("!!!!! Erreur d'initialisation");
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
            Debug.Info(MethodBase.GetCurrentMethod());

            if ((Piece != null) && Piece.EstInitialise && (Piece.Contient(TypeCorps_e.cTole)))
            {
                _Piece = Piece;
                _EstInitialise = true;
            }
            else
            {
                Debug.Info("!!!!! Erreur d'initialisation");
            }
            return _EstInitialise;
        }

#endregion
    }

#endif
}
