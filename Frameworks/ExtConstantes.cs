using System;
using System.Runtime.InteropServices;

namespace Frameworks2013
{

    #region "Enumérations"

    public enum TypeFichier_e
    {
        cAucun = 0,
        cAssemblage = 1,
        cPiece = 2,
        cDessin = 4,
        cTousLesTypesDeFichier = 7,
    }

    public enum TypeCorps_e
    {
    cTole = 1,
    cProfil = cTole * 2,
    cAutre = cProfil * 2,
    cTousLesTypesDeCorps = cTole + cProfil + cAutre
    }

    public enum TypeConfig_e
    {
        cDeBase = 1,
        cDerivee = 2,
        cDepliee = 4,
        cPliee = 8,
        cToutesLesTypesDeConfig = 15
    }

    public enum EtatFonction_e
    {
        cDesactivee = 0,
        cActivee = 1
    }

    public enum Orientation_e
    {
        cPortrait = 1,
        cPaysage = 2
    }

    #endregion

    #region "Structures"

    public struct Point
    {
        public Double X;
        public Double Y; 
        public Double Z;
    }

    public struct Dimensions
    {
        public Double Lg; 
        public Double Ht; 
        }

    public struct Rectangle
    {
        public Double MinX;
        public Double MinY;
        public Double MaxX;
        public Double MaxY;
    }

    #endregion

    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("78322EBA-506F-11E2-AD84-26056188709B")]
    public interface IExtConstantes
    {
        String CONFIG_DEPLIEE { get; set; }
        String CONFIG_PLIEE { get; set; }
        String ARTICLE_LISTE_DES_PIECES_SOUDEES { get; set; }
        String EPAISSEUR_DE_TOLE { get; set; }
        String NO_DOSSIER { get; set; }
        String NOM_ELEMENT { get; set; }
        String CUBE_DE_VISUALISATION { get; set; }
        String MODELE_DE_DESSIN_LASER { get; set; }
        String NOM_CORPS_DEPLIEE { get; set; }
        String ETAT_D_AFFICHAGE { get; set; }
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("9853F1B0-506F-11E2-AE27-2E056188709B")]
    [ProgId("Frameworks.ExtConstantes")]
    public class ExtConstantes : IExtConstantes
    {
        #region "Variables locales"

        private String _CONFIG_DEPLIEE = "SM-FLAT-PATTERN";
        private String _CONFIG_PLIEE = "#";
        private String _ARTICLE_LISTE_DES_PIECES_SOUDEES = "Article-liste-des-pièces-soudées";
        private String _EPAISSEUR_DE_TOLE = "Epaisseur de la tôle";
        private String _NO_DOSSIER = "NoDossier";
        private String _NOM_ELEMENT = "Element";
        private String _CUBE_DE_VISUALISATION = "Cube de visualisation";
        private String _MODELE_DE_DESSIN_LASER = "MacroLaser";
        private String _NOM_CORPS_DEPLIEE = "Etat déplié";
        private String _ETAT_D_AFFICHAGE = "Etat d'affichage-";

        #endregion

        #region "Constructeur\Destructeur"

        public ExtConstantes()
        {
        }

        #endregion

        #region "Propriétés"

        public String CONFIG_DEPLIEE
        {
            get { return _CONFIG_DEPLIEE; }
            set { _CONFIG_DEPLIEE = value; }
        }

        public String CONFIG_PLIEE
        {
            get { return _CONFIG_PLIEE; }
            set { _CONFIG_PLIEE = value; }
        }

        public String ARTICLE_LISTE_DES_PIECES_SOUDEES
        {
            get { return _ARTICLE_LISTE_DES_PIECES_SOUDEES; }
            set { _ARTICLE_LISTE_DES_PIECES_SOUDEES = value; }
        }

        public String EPAISSEUR_DE_TOLE
        {
            get { return _EPAISSEUR_DE_TOLE; }
            set { _EPAISSEUR_DE_TOLE = value; }
        }

        public String NO_DOSSIER
        {
            get { return _NO_DOSSIER; }
            set { _NO_DOSSIER = value; }
        }

        public String NOM_ELEMENT
        {
            get { return _NOM_ELEMENT; }
            set { _NOM_ELEMENT = value; }
        }

        public String CUBE_DE_VISUALISATION
        {
            get { return _CUBE_DE_VISUALISATION; }
            set { _CUBE_DE_VISUALISATION = value; }
        }

        public String MODELE_DE_DESSIN_LASER
        {
            get { return _MODELE_DE_DESSIN_LASER; }
            set { _MODELE_DE_DESSIN_LASER = value; }
        }

        public String NOM_CORPS_DEPLIEE
        {
            get { return _NOM_CORPS_DEPLIEE; }
            set { _NOM_CORPS_DEPLIEE = value; }
        }

        public String ETAT_D_AFFICHAGE
        {
            get { return _ETAT_D_AFFICHAGE; }
            set { _ETAT_D_AFFICHAGE = value; }
        }

        #endregion

        #region "Méthodes"

        #endregion
    }
}
