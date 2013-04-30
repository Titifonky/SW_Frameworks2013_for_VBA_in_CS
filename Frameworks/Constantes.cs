using System;
using System.Runtime.InteropServices;
using SolidWorks.Interop.swconst;

namespace Framework_SW2013
{

    #region "Enumérations"

    //Cet attribut permet de combiner les valeurs d'enumération
    [Flags]
    public enum TypeFichier_e
    {
        cAucun = 0,
        cAssemblage = 1,
        cPiece = 2,
        cDessin = 4
    }

    //Cet attribut permet de combiner les valeurs d'enumération
    [Flags]
    public enum TypeCorps_e
    {
        cAucun = 0,
        cTole = 1,
        cProfil = 2,
        cAutre = 4,
        cTousLesTypesDeCorps = cTole + cProfil + cAutre
    }

    //Cet attribut permet de combiner les valeurs d'enumération
    [Flags]
    public enum TypeConfig_e
    {
        cDeBase = 1,
        cDerivee = 2,
        cDepliee = 4,
        cPliee = 8,
        cToutesLesTypesDeConfig = cDeBase + cDerivee + cDepliee + cPliee
    }

    public enum EtatFonction_e
    {
        cDesactivee = 0,
        cActivee = 1
    }

    public enum Orientation_e
    {
        cPortrait = 0,
        cPaysage = 1
    }

    #endregion

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("5015C3D8-9DC1-4B38-BA40-E04AEA31A45A")]
    [ProgId("Frameworks.Constantes")]
    public static class Constantes
    {
        #region "Variables locales"

        public static String CONFIG_DEPLIEE = "^*SM-FLAT-PATTERN*";
        public static String CONFIG_PLIEE = "^[0-9]";
        public static String ARTICLE_LISTE_DES_PIECES_SOUDEES = "Article-liste-des-pièces-soudées";
        public static String EPAISSEUR_DE_TOLE = "Epaisseur de la tôle";
        public static String NO_DOSSIER = "NoDossier";
        public static String NOM_ELEMENT = "Element";
        public static String CUBE_DE_VISUALISATION = "Cube de visualisation";
        public static String MODELE_DE_DESSIN_LASER = "MacroLaser";
        public static String NOM_CORPS_DEPLIEE = "Etat déplié";
        public static String ETAT_D_AFFICHAGE = "Etat d'affichage-";
        public static String BIB_MATERIAUX = "C:/Users/billiet-pao/Solidworks/Materiaux/Matériaux Billiet.sldmat";

        #endregion

    }
}
