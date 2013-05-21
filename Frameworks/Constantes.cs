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

    public enum Extension_e
    {
        cDXF = 1,
        cDWG = 2,
        cPDF = 4
    }

    #endregion

    internal static class CONSTANTES
    {
        #region "Variables locales"

        internal static String CONFIG_DEPLIEE = "^*SM-FLAT-PATTERN*";
        internal static String CONFIG_PLIEE = "^[0-9]";
        internal static String ARTICLE_LISTE_DES_PIECES_SOUDEES = "Article-liste-des-pièces-soudées";
        internal static String EPAISSEUR_DE_TOLE = "Epaisseur de la tôle";
        internal static String NO_DOSSIER = "NoDossier";
        internal static String NOM_ELEMENT = "Element";
        internal static String CUBE_DE_VISUALISATION = "Cube de visualisation";
        internal static String MODELE_DE_DESSIN_LASER = "MacroLaser";
        internal static String NOM_CORPS_DEPLIEE = "Etat déplié";
        internal static String ETAT_D_AFFICHAGE = "Etat d'affichage-";
        internal static String BIB_MATERIAUX = "C:/Users/billiet-pao/Solidworks/Materiaux/Matériaux Billiet.sldmat";

        #endregion
    }

    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("EE3D32C7-E632-4DCE-83F7-CF58F569F523")]
    public interface IConstantes
    {
        String ARTICLE_LISTE_DES_PIECES_SOUDEES { get; set; }
        String EPAISSEUR_DE_TOLE { get; set; }
        String NO_DOSSIER { get; set; }
        String NOM_ELEMENT { get; set; }
        String MODELE_DE_DESSIN_LASER { get; set; }
        String NOM_CORPS_DEPLIEE { get; set; }
        String ETAT_D_AFFICHAGE { get; set; }
        String BIB_MATERIAUX { get; set; }
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("5015C3D8-9DC1-4B38-BA40-E04AEA31A45A")]
    [ProgId("Frameworks.Constantes")]
    public class Constantes : IConstantes
    {
        #region "Propriétés"

        public String ARTICLE_LISTE_DES_PIECES_SOUDEES { get { return CONSTANTES.ARTICLE_LISTE_DES_PIECES_SOUDEES; } set { CONSTANTES.ARTICLE_LISTE_DES_PIECES_SOUDEES = value; } }
        public String EPAISSEUR_DE_TOLE { get { return CONSTANTES.EPAISSEUR_DE_TOLE; } set { CONSTANTES.EPAISSEUR_DE_TOLE = value; } }
        public String NO_DOSSIER { get { return CONSTANTES.NO_DOSSIER; } set { CONSTANTES.NO_DOSSIER = value; } }
        public String NOM_ELEMENT { get { return CONSTANTES.NOM_ELEMENT; } set { CONSTANTES.NOM_ELEMENT = value; } }
        public String MODELE_DE_DESSIN_LASER { get { return CONSTANTES.MODELE_DE_DESSIN_LASER; } set { CONSTANTES.MODELE_DE_DESSIN_LASER = value; } }
        public String NOM_CORPS_DEPLIEE { get { return CONSTANTES.NOM_CORPS_DEPLIEE; } set { CONSTANTES.NOM_CORPS_DEPLIEE = value; } }
        public String ETAT_D_AFFICHAGE { get { return CONSTANTES.ETAT_D_AFFICHAGE; } set { CONSTANTES.ETAT_D_AFFICHAGE = value; } }
        public String BIB_MATERIAUX { get { return CONSTANTES.BIB_MATERIAUX; } set { CONSTANTES.BIB_MATERIAUX = value; } }

        #endregion
    }
}
