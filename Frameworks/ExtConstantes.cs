using System;
using System.Runtime.InteropServices;

namespace Framework2013
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

    public static class ExtConstantes
    {
        #region "Variables locales"

        public static String _CONFIG_DEPLIEE = "SM-FLAT-PATTERN";
        public static String _CONFIG_PLIEE = "#";
        public static String _ARTICLE_LISTE_DES_PIECES_SOUDEES = "Article-liste-des-pièces-soudées";
        public static String _EPAISSEUR_DE_TOLE = "Epaisseur de la tôle";
        public static String _NO_DOSSIER = "NoDossier";
        public static String _NOM_ELEMENT = "Element";
        public static String _CUBE_DE_VISUALISATION = "Cube de visualisation";
        public static String _MODELE_DE_DESSIN_LASER = "MacroLaser";
        public static String _NOM_CORPS_DEPLIEE = "Etat déplié";
        public static String _ETAT_D_AFFICHAGE = "Etat d'affichage-";

        #endregion

    }
}
