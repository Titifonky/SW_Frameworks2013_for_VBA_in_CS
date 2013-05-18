using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Reflection;
using System.IO;
using SolidWorks.Interop.swconst;

namespace Framework_SW2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("FAB06846-3E7C-421E-B20A-395906EDEEB0")]
    public interface IExtFichierSW
    {
        ExtSldWorks SW { get; }
        String Chemin { get; }
        String Configuration { get; set; }
        long Nb { get; set; }
        TypeFichier_e TypeDuFichier { get; }
        String NomDuFichier { get; }
        String NomDuFichierSansExt { get; }
        String NomDuDossier { get; }
        eModele Ouvrir();
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("C2D869E0-06B4-4E07-BFA3-96F67307ABEB")]
    [ProgId("Frameworks.ExtFichierSW")]
    public class ExtFichierSW : IExtFichierSW
    {
        #region "Variables locales"
        private Boolean _EstInitialise = false;

        private ExtSldWorks _SW;
        private String _Chemin;
        private String _Configuration;
        private long _Nb;

        #endregion

        #region "Constructeur\Destructeur"

        public ExtFichierSW() { }

        #endregion

        #region "Propriétés"

        public ExtSldWorks SW { get { return _SW; } }
        public String Chemin { get { return _Chemin; } internal set { _Chemin = value; } }
        public String Configuration { get { return _Configuration; } set { _Configuration = value; } }
        public long Nb { get { return _Nb; } set { _Nb = value; } }

        /// <summary>
        /// Retourne le type du fichier.
        /// </summary>
        public TypeFichier_e TypeDuFichier
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                switch (Path.GetExtension(_Chemin).ToUpper())
                {
                    case ".SLDASM":
                        return TypeFichier_e.cAssemblage;

                    case ".SLDPRT":
                        return TypeFichier_e.cPiece;

                    case ".SLDDRW":
                        return TypeFichier_e.cDessin;

                    default:
                        return TypeFichier_e.cAucun;
                }
            }
        }

        /// <summary>
        /// Retourne le nom du fichier avec extension.
        /// </summary>
        public String NomDuFichier { get { Debug.Info(MethodBase.GetCurrentMethod()); return Path.GetFileName(_Chemin); } }

        /// <summary>
        /// Retourne le nom du fichier sans extension.
        /// </summary>
        public String NomDuFichierSansExt { get { Debug.Info(MethodBase.GetCurrentMethod()); return Path.GetFileNameWithoutExtension(_Chemin); } }

        /// <summary>
        /// Retourne le chemin du dossier.
        /// </summary>
        public String NomDuDossier { get { Debug.Info(MethodBase.GetCurrentMethod()); return Path.GetDirectoryName(_Chemin); } }


        /// <summary>
        /// Fonction interne.
        /// Test l'initialisation de l'objet ExtFichierSW.
        /// </summary>
        internal Boolean EstInitialise { get { Debug.Info(MethodBase.GetCurrentMethod()); return _EstInitialise; } }

        #endregion

        #region "Méthodes"

        /// <summary>
        /// Méthode interne.
        /// Initialiser l'objet ExtFichierSW.
        /// </summary>
        /// <param name="Sw"></param>
        /// <returns></returns>
        internal Boolean Init(ExtSldWorks Sw)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            if ((Sw != null) && Sw.EstInitialise)
            {
                _SW = Sw;

                // On valide l'initialisation
                _EstInitialise = true;
            }
            else
            {
                Debug.Info("!!!!! Erreur d'initialisation");
            }

            return _EstInitialise;
        }

        public eModele Ouvrir()
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            eModele pModele = _SW.Modele(_Chemin);
            eConfiguration pConfig = pModele.GestDeConfigurations.ConfigurationAvecLeNom(_Configuration);
            pConfig.Activer();
            return pModele;
        }

        #endregion
    }
}
