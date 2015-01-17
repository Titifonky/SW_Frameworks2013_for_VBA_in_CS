using SolidWorks.Interop.swconst;
using System;
using System.Collections;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace Framework
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("44D349FF-E4E1-4C9A-86EB-197EC4AE0DBE")]
    public interface IeTole
    {
        eCorps Corps { get; }
        eParametreTolerie ParametresDeTolerie { get; }
        eFonction FonctionTolerie { get; }
        eFonction FonctionToleDeBase { get; }
        eFonction FonctionDepliee { get; }
        eFonction FonctionCubeDeVisualisation { get; }
        eConfiguration ConfigurationDepliee { get; }
        Boolean EstDepliee { get; }
        String NomVueDepliee { get; }
        void Deplier(Boolean T);
        eConfiguration CreerConfigurationDepliee(Boolean Ecraser = false);
        void MajConfigurationDepliee();
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("DCE24670-A986-4606-88FE-577980D34685")]
    [ProgId("Frameworks.eTole")]
    public class eTole : IeTole
    {
        #region "Variables locales"

        private static readonly String cNOMCLASSE = typeof(eTole).Name;

        private Boolean _EstInitialise = false;

        private eCorps _Corps = null;
        private ePiece _Piece = null;
        private eParametreTolerie _ParamTolerie = null;

        #endregion

        #region "Constructeur\Destructeur"

        public eTole() { }

        #endregion

        #region "Propriétés"

        /// <summary>
        /// Retourne le parent eCorps.
        /// </summary>
        public eCorps Corps { get { Log.Methode(cNOMCLASSE); return _Corps; } }

        public eParametreTolerie ParametresDeTolerie
        {
            get
            {
                Log.Methode(cNOMCLASSE);

                if (_ParamTolerie == null)
                {
                    _ParamTolerie = new eParametreTolerie();
                    _ParamTolerie.Init(this);
                }

                if (_ParamTolerie.EstInitialise)
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
                Log.Methode(cNOMCLASSE);

                swFeatureType_e TypeFonc = new swFeatureType_e();

                foreach (eFonction pFonc in _Corps.ListeDesFonctions())
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
                Log.Methode(cNOMCLASSE);

                swFeatureType_e TypeFonc = new swFeatureType_e();

                foreach (eFonction pFonc in _Corps.ListeDesFonctions())
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
        public eFonction FonctionDepliee
        {
            get
            {
                Log.Methode(cNOMCLASSE);

                swFeatureType_e TypeFonc = new swFeatureType_e();

                foreach (eFonction pFonc in _Corps.ListeDesFonctions())
                {
                    if (pFonc.TypeDeLaFonction == TypeFonc.swTnFlatPattern)
                        return pFonc;
                }

                return null;
            }
        }

        public String NomVueDepliee
        {
            get
            {
                return _Corps.Piece.Modele.FichierSw.NomDuFichierSansExt + " - " + NomConfigDepliee;
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
                Log.Methode(cNOMCLASSE);

                return (eFonction)this.FonctionDepliee.ListeDesSousFonctions(CONSTANTES.CUBE_DE_VISUALISATION)[0];
            }
        }

        /// <summary>
        /// Plier ou deplier la tole
        /// </summary>
        public void Deplier(Boolean T)
        {
            Log.Methode(cNOMCLASSE);
            if (T)
            {
                FonctionDepliee.Activer();
                return;
            }

            FonctionDepliee.Desactiver();
        }

        /// <summary>
        /// Retourne le nom de la configuration dépliée associée au corps.
        /// </summary>
        private String NomConfigDepliee
        {
            get
            {
                Log.Methode(cNOMCLASSE);

                eConfiguration pConfigActive = _Piece.Modele.GestDeConfigurations.ConfigurationActive;

                if (EstDepliee && pConfigActive.Est(TypeConfig_e.cDepliee))
                    return pConfigActive.Nom;

                eGestDeProprietes pGestProps = _Corps.Dossier.GestDeProprietes;

                String pNoDossier = "";

                if (pGestProps.ProprieteExiste(CONSTANTES.NO_DOSSIER))
                {
                    pNoDossier = pGestProps.RecupererPropriete(CONSTANTES.NO_DOSSIER).Valeur;

                    if (pConfigActive.Est(TypeConfig_e.cPliee))
                        return pConfigActive.Nom + CONSTANTES.CONFIG_DEPLIEE + pNoDossier;
                    else if (pConfigActive.Est(TypeConfig_e.cDepliee))
                        return pConfigActive.ConfigurationParent.Nom + CONSTANTES.CONFIG_DEPLIEE + pNoDossier;
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
                Log.Methode(cNOMCLASSE);
                return _Piece.Modele.GestDeConfigurations.ConfigurationAvecLeNom(NomConfigDepliee);
            }
        }

        public Boolean EstDepliee
        {
            get
            {

                eCorps pCorps = _Piece.CorpsDeplie();

                if ((pCorps != null) && (_Piece.Modele.SwModele.Extension.IsSamePersistentID(_Corps.PID, pCorps.PID) == 1))
                    return true;

                return false;
            }
        }

        /// <summary>
        /// Fonction interne.
        /// Test l'initialisation de l'objet eTole.
        /// </summary>
        internal Boolean EstInitialise { get { Log.Methode(cNOMCLASSE); return _EstInitialise; } }

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
            Log.Methode(cNOMCLASSE);

            if ((Corps != null) && Corps.EstInitialise && (Corps.TypeDeCorps == TypeCorps_e.cTole))
            {
                _Corps = Corps;
                _Piece = Corps.Piece;
                _EstInitialise = true;
            }
            else
            {
                Log.Message("!!!!! Erreur d'initialisation");
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
            Log.Methode(cNOMCLASSE);

            eGestDeConfigurations pGestConfig = _Piece.Modele.GestDeConfigurations;
            eConfiguration pConfigActive = pGestConfig.ConfigurationActive;
            String pNomConfigDepliee = NomConfigDepliee;

            if (pConfigActive.Est(TypeConfig_e.cPliee))
            {
                if (Ecraser)
                {
                    pGestConfig.SupprimerConfiguration(pNomConfigDepliee);
                    _Piece.Modele.Sauver();
                }

                Log.Message("------------------------------> " + pNomConfigDepliee);
                eConfiguration pConfigDepliee = ConfigurationDepliee;

                if (pConfigDepliee == null)
                    pConfigDepliee = pConfigActive.AjouterUneConfigurationDerivee(pNomConfigDepliee);

                Log.Message(" ==========================   " + (pConfigDepliee != null).ToString());

                if (pConfigDepliee != null)
                {
                    pConfigDepliee.GestDeProprietes.AjouterPropriete(CONSTANTES.NO_CONFIG, swCustomInfoType_e.swCustomInfoText, pConfigActive.Nom, true);
                    pConfigDepliee.RenommerEtatAffichage();

                    MajConfigurationDepliee();

                    return pConfigDepliee;
                }
            }
            return null;
        }

        public void MajConfigurationDepliee()
        {
            ConfigurationDepliee.Activer();
            FonctionDepliee.Activer(ConfigurationDepliee, true);

            foreach (eFonction sF in FonctionDepliee.ListeDesSousFonctions())
            {
                sF.Activer(ConfigurationDepliee);
                if ((Regex.IsMatch(sF.Nom, CONSTANTES.LIGNES_DE_PLIAGE)) ||
                    (Regex.IsMatch(sF.Nom, CONSTANTES.CUBE_DE_VISUALISATION)))
                {
                    sF.Selectionner(false);
                    _Piece.Modele.SwModele.UnblankSketch();
                    _Piece.Modele.EffacerLesSelections();
                }
            }
        }

        #endregion
    }
}
