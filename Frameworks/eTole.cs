using System;
using System.Reflection;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Collections.Generic;

namespace Framework_SW2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("44D349FF-E4E1-4C9A-86EB-197EC4AE0DBE")]
    public interface IeTole
    {
        eCorps Corps { get; }
        eParametreTolerie ParametresDeTolerie { get; }
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
        private eParametreTolerie _ParamTolerie = null;

        #endregion

        #region "Constructeur\Destructeur"

        public eTole() { }

        #endregion

        #region "Propriétés"

        /// <summary>
        /// Retourne le parent ExtPiece.
        /// </summary>
        public eCorps Corps { get { Debug.Print(MethodBase.GetCurrentMethod()); return _Corps; } }

        public eParametreTolerie ParametresDeTolerie
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());

                if (_ParamTolerie == null)
                {
                    _ParamTolerie = new eParametreTolerie();
                    _ParamTolerie.Init(this);
                }
                    
                if(_ParamTolerie.EstInitialise)
                    return _ParamTolerie;

                return null;
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
                Debug.Print(MethodBase.GetCurrentMethod());

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
                Debug.Print(MethodBase.GetCurrentMethod());

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
                Debug.Print(MethodBase.GetCurrentMethod());

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
                Debug.Print(MethodBase.GetCurrentMethod());

                return this.FonctionDeplie.ListListeDesSousFonctions(CONSTANTES.CUBE_DE_VISUALISATION)[0];
            }
        }

        /// <summary>
        /// Plier ou deplier la tole
        /// </summary>
        public void Deplier(Boolean T)
        {
            Debug.Print(MethodBase.GetCurrentMethod());
            if (T)
                FonctionDeplie.Activer();

            FonctionDeplie.Desactiver();
        }

        /// <summary>
        /// Retourne le nom de la configuration dépliée associée au corps.
        /// </summary>
        private String NomConfigDepliee
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());
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
        /// Renvoi la configuration dépliée associée au corps de tôlerie
        /// </summary>
        public eConfiguration ConfigurationDepliee
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                return _Corps.Piece.Modele.GestDeConfigurations.ConfigurationAvecLeNom(NomConfigDepliee);
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
        /// Initialiser l'objet ExtCorps.
        /// </summary>
        /// <param name="SwCorps"></param>
        /// <param name="Corps"></param>
        /// <returns></returns>
        internal Boolean Init(eCorps Corps)
        {
            Debug.Print(MethodBase.GetCurrentMethod());

            if ((Corps != null) && Corps.EstInitialise && (Corps.TypeDeCorps == TypeCorps_e.cTole))
            {
                _Corps = Corps;
                _EstInitialise = true;
            }
            else
            {
                Debug.Print("!!!!! Erreur d'initialisation");
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
            Debug.Print(MethodBase.GetCurrentMethod());

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

                Debug.Print(" ==========================   " + (pConfigDepliee != null).ToString());

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
