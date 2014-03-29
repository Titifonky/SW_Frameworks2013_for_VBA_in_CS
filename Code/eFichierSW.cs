using System;
using System.IO;
using System.Runtime.InteropServices;

namespace Framework
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("FAB06846-3E7C-421E-B20A-395906EDEEB0")]
    public interface IeFichierSW
    {
        eSldWorks SW { get; }
        String Chemin { get; }
        String Configuration { get; set; }
        int Nb { get; }
        TypeFichier_e TypeDuFichier { get; }
        String NomDuFichier { get; }
        String NomDuFichierSansExt { get; }
        String NomDuDossier { get; }
        eModele Ouvrir();
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("C2D869E0-06B4-4E07-BFA3-96F67307ABEB")]
    [ProgId("Frameworks.eFichierSW")]
    public class eFichierSW : IeFichierSW
    {
        #region "Variables locales"

        private static readonly String cNOMCLASSE = typeof(eFichierSW).Name;
        private Boolean _EstInitialise = false;

        private eSldWorks _SW = null;
        private String _Chemin = "";
        private String _NomConfiguration = "";
        private int _Nb = 1;

        #endregion

        #region "Constructeur\Destructeur"

        public eFichierSW() { }

        #endregion

        #region "Propriétés"

        public eSldWorks SW { get { return _SW; } }
        public String Chemin { get { return _Chemin; } internal set { _Chemin = value; } }
        public String Configuration { get { return _NomConfiguration; } set { _NomConfiguration = value; } }
        public int Nb { get { return _Nb; } internal set { _Nb = value; } }

        /// <summary>
        /// Retourne le type du fichier.
        /// </summary>
        public TypeFichier_e TypeDuFichier
        {
            get
            {
                Log.Methode(cNOMCLASSE);
                switch (Path.GetExtension(_Chemin).ToUpper())
                {
                    case ".SLDASM":
                        return TypeFichier_e.cAssemblage;

                    case ".SLDPRT":
                        return TypeFichier_e.cPiece;

                    case ".SLDLFP":
                        return TypeFichier_e.cLibrairie;

                    case ".SLDDRW":
                        return TypeFichier_e.cDessin;

                    default:
                        return 0;
                }
            }
        }

        /// <summary>
        /// Retourne le nom du fichier avec extension.
        /// </summary>
        public String NomDuFichier
        {
            get
            {
                Log.Methode(cNOMCLASSE);
                if (!String.IsNullOrEmpty(_Chemin))
                    return Path.GetFileName(_Chemin);

                return "";
            }
        }

        /// <summary>
        /// Retourne le nom du fichier sans extension.
        /// </summary>
        public String NomDuFichierSansExt
        {
            get
            {
                Log.Methode(cNOMCLASSE);
                if (!String.IsNullOrEmpty(_Chemin))
                    return Path.GetFileNameWithoutExtension(_Chemin);

                return "";
            }
        }

        /// <summary>
        /// Retourne le chemin du dossier.
        /// </summary>
        public String NomDuDossier
        {
            get
            {
                Log.Methode(cNOMCLASSE);
                if (!String.IsNullOrEmpty(_Chemin))
                    return Path.GetDirectoryName(_Chemin);

                return "";
            }
        }


        /// <summary>
        /// Fonction interne.
        /// Test l'initialisation de l'objet ExtFichierSW.
        /// </summary>
        internal Boolean EstInitialise { get { Log.Methode(cNOMCLASSE); return _EstInitialise; } }

        #endregion

        #region "Méthodes"

        /// <summary>
        /// Méthode interne.
        /// Initialiser l'objet ExtFichierSW.
        /// </summary>
        /// <param name="Sw"></param>
        /// <returns></returns>
        internal Boolean Init(eSldWorks Sw)
        {
            Log.Methode(cNOMCLASSE);

            if ((Sw != null) && Sw.EstInitialise)
            {
                _SW = Sw;
                // On valide l'initialisation
                _EstInitialise = true;
            }
            else
            {
                Log.Message("!!!!! Erreur d'initialisation");
            }

            return _EstInitialise;
        }

        public eModele Ouvrir()
        {
            Log.Methode(cNOMCLASSE);

            if ((TypeFichier_e.cAssemblage | TypeFichier_e.cPiece | TypeFichier_e.cDessin).HasFlag(TypeDuFichier))
            {
                eModele pModele = _SW.Modele(_Chemin);
                if (!String.IsNullOrEmpty(_NomConfiguration))
                {
                    eConfiguration pConfig = pModele.GestDeConfigurations.ConfigurationAvecLeNom(_NomConfiguration);
                    pConfig.Activer();
                    pModele.ReinitialiserComposant();
                    pModele.Composant.Configuration = pConfig;
                    pModele.Composant.Nb = Nb;
                }
                return pModele;
            }

            return null;
        }

        #endregion
    }
}
