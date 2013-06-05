using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using System.Collections;
using System.Reflection;
using SolidWorks.Interop.swconst;

namespace Framework_SW2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("44D349FF-E4E1-4C9A-86EB-197EC4AE0DBE")]
    public interface IeTole
    {
        eCorps Corps { get; }
        Double Epaisseur { get; set; }
        Double Rayon { get; set; }
        Double FacteurK { get; set; }
        Boolean EcraserEpaisseur { get; set; }
        Boolean EcraserRayon { get; set; }
        Boolean EcraserFacteurK { get; set; }
        eFonction FonctionTolerie { get; }
        eFonction FonctionToleDeBase { get; }
        eFonction FonctionDeplie { get; }
        eFonction FonctionCubeDeVisualisation { get; }
        eConfiguration ConfigurationDepliee { get; }
        void Deplier(Boolean T);
        eConfiguration CreerConfigurationDepliee(Boolean Ecraser = false);
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("DCE24670-A986-4606-88FE-577980D34685")]
    [ProgId("Frameworks.eTole")]
    public class eTole : IeTole
    {
        #region "Variables locales"
        
        private Boolean _EstInitialise = false;

        private eCorps _Corps = null;

        #endregion

        #region "Constructeur\Destructeur"

        public eTole() { }

        #endregion

        #region "Propriétés"

        /// <summary>
        /// Retourne le parent ExtPiece.
        /// </summary>
        public eCorps Corps { get { Debug.Info(MethodBase.GetCurrentMethod());  return _Corps; } }

        private ModelDoc2 SwModele
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                if(_Corps.Piece.Modele.Equals(_Corps.Piece.Modele.SW.Modele()))
                    return _Corps.Piece.Modele.SwModele;

                return null;
            }
        }

        private Component2 SwComposant
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                if (SwModele == null)
                    return _Corps.Piece.Modele.Composant.SwComposant;

                return null;
            }
        }

        /// <summary>
        /// Retourne ou défini l'épaisseur de la tole
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
                Feature pSwFonction = FonctionToleDeBase.SwFonction;
                BaseFlangeFeatureData pParam = pSwFonction.GetDefinition();
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
                Feature pSwFonctionTolerie = FonctionTolerie.SwFonction;
                SheetMetalFeatureData pParam = pSwFonctionTolerie.GetDefinition();
                pParam.AccessSelections(pSwModele, pSwComposant);
                pParam.BendRadius = value * 0.001;
                pSwFonctionTolerie.ModifyDefinition(pParam, pSwModele, pSwComposant);
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
                Feature pSwFonctionTolerie = FonctionTolerie.SwFonction;
                SheetMetalFeatureData pParam = pSwFonctionTolerie.GetDefinition();
                CustomBendAllowance pParamPli = pParam.GetCustomBendAllowance();
                pParam.AccessSelections(pSwModele, pSwComposant);
                pParamPli.KFactor = value;
                pParam.SetCustomBendAllowance(pParamPli);
                pSwFonctionTolerie.ModifyDefinition(pParam, pSwModele, pSwComposant);
                pParam.ReleaseSelectionAccess();
            }
        }

        public Boolean EcraserEpaisseur
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                BaseFlangeFeatureData pParam = FonctionToleDeBase.SwFonction.GetDefinition();
                return pParam.OverrideThickness;
            }
            set
            {
                Debug.Info(MethodBase.GetCurrentMethod());
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
                Debug.Info(MethodBase.GetCurrentMethod());
                BaseFlangeFeatureData pParam = FonctionToleDeBase.SwFonction.GetDefinition();
                return pParam.OverrideRadius;
            }
            set
            {
                Debug.Info(MethodBase.GetCurrentMethod());
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
                Debug.Info(MethodBase.GetCurrentMethod());
                BaseFlangeFeatureData pParam = FonctionToleDeBase.SwFonction.GetDefinition();
                return pParam.OverrideKFactor;
            }
            set
            {
                Debug.Info(MethodBase.GetCurrentMethod());
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
        /// Renvoi la fonction Tolerie du corps
        /// </summary>
        /// <returns></returns>
        public eFonction FonctionTolerie
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());

                swFeatureType_e TypeFonc = new swFeatureType_e();

                foreach (eFonction pFonc in _Corps.ListListeDesFonctions())
                {
                    if (pFonc.TypeDeLaFonction == TypeFonc.swTnSheetMetal)
                        return pFonc;
                }

                return null;
            }
        }

        /// <summary>
        /// Renvoi la fonction Tolerie du corps
        /// </summary>
        /// <returns></returns>
        public eFonction FonctionToleDeBase
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());

                swFeatureType_e TypeFonc = new swFeatureType_e();

                foreach (eFonction pFonc in _Corps.ListListeDesFonctions())
                {
                    if ((pFonc.TypeDeLaFonction == TypeFonc.swTnBaseFlange) || (pFonc.TypeDeLaFonction == TypeFonc.swTnSolidToSheetMetal))
                        return pFonc;

                }

                return null;
            }
        }

        /// <summary>
        /// Renvoi la fonction EtatDeplie du corps
        /// </summary>
        /// <returns></returns>
        public eFonction FonctionDeplie
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());

                swFeatureType_e TypeFonc = new swFeatureType_e();

                foreach (eFonction pFonc in _Corps.ListListeDesFonctions())
                {
                    if (pFonc.TypeDeLaFonction == TypeFonc.swTnFlatPattern)
                        return pFonc;
                }

                return null;
            }
        }

        /// <summary>
        /// Renvoi la fonction CubeDeVisualisation du corps
        /// </summary>
        /// <returns></returns>
        public eFonction FonctionCubeDeVisualisation
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());

                return this.FonctionDeplie.ListListeDesSousFonctions(CONSTANTES.CUBE_DE_VISUALISATION)[0];
            }
        }

        /// <summary>
        /// Plier ou deplier la tole
        /// </summary>
        public void Deplier (Boolean T)
        {
                Debug.Info(MethodBase.GetCurrentMethod());
                if (T)
                    FonctionDeplie.Activer();

                FonctionDeplie.Desactiver();
        }

        private String NomConfigDepliee
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                eGestDeConfigurations pGestConfig = _Corps.Piece.Modele.GestDeConfigurations;
                eGestDeProprietes pGestProps = _Corps.Dossier.GestDeProprietes;

                String pNoDossier = "";

                if (pGestProps.ProprieteExiste(CONSTANTES.NO_DOSSIER))
                {
                    pNoDossier = pGestProps.RecupererPropriete(CONSTANTES.NO_DOSSIER).Valeur;
                    return pGestConfig.ConfigurationActive.Nom + CONSTANTES.CONFIG_DEPLIEE + pNoDossier;
                }

                return "";
            }
        }

        /// <summary>
        /// Revoi la configuration dépliée associée au corps de tôlerie
        /// </summary>
        public eConfiguration ConfigurationDepliee
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                return _Corps.Piece.Modele.GestDeConfigurations.ConfigurationAvecLeNom(NomConfigDepliee);
            }
        }

        /// <summary>
        /// Fonction interne.
        /// Test l'initialisation de l'objet eTole.
        /// </summary>
        internal Boolean EstInitialise { get { Debug.Info(MethodBase.GetCurrentMethod());  return _EstInitialise; } }

        #endregion

        #region "Méthodes"

        /// <summary>
        /// Méthode interne.
        /// Initialiser l'objet ExtCorps.
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
        /// Ajoute une configuration dépliée.
        /// </summary>
        /// <param name="NomConfigDerivee"></param>
        /// <returns></returns>
        public eConfiguration CreerConfigurationDepliee(Boolean Ecraser = false)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            eGestDeConfigurations pGestConfig = _Corps.Piece.Modele.GestDeConfigurations;
            eConfiguration pConfigActive = pGestConfig.ConfigurationActive;
            String pNomConfigDepliee = NomConfigDepliee;

            if (pConfigActive.Est(TypeConfig_e.cPliee))
            {
                if (Ecraser)
                {
                    pGestConfig.SupprimerConfiguration(pNomConfigDepliee);
                    _Corps.Piece.Modele.Sauver();
                }

                eConfiguration pConfigDepliee = ConfigurationDepliee;

                if (pConfigDepliee == null)
                    pConfigDepliee = pConfigActive.AjouterUneConfigurationDerivee(pNomConfigDepliee);

                Debug.Info(" ==========================   " + (pConfigDepliee != null).ToString());

                if (pConfigDepliee != null)
                {
                    pConfigDepliee.GestDeProprietes.AjouterPropriete(CONSTANTES.NO_CONFIG, swCustomInfoType_e.swCustomInfoText, pConfigActive.Nom, true);
                    pConfigDepliee.RenommerEtatAffichage();
                    return pConfigDepliee;
                }
            }
            return null;
        }

        #endregion
    }
}
