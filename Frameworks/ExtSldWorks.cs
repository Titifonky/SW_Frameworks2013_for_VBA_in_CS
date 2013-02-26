using System;
using System.IO;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

/////////////////////////// Implementation terminée ///////////////////////////

namespace Framework_SW2013
{

    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("9ED6BE92-5820-11E2-9D5D-93046188709B")]
    public interface IExtSldWorks
    {
        SldWorks SwSW { get; }
        TypeFichier_e TypeDuModeleActif { get; }
        String VersionDeBase { get; }
        String VersionCourante { get; }
        String Hotfixe { get; }
        String Revision { get; }
        Boolean Init(SldWorks SldWks);
        ExtModele Modele(String Chemin = "");
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("A2E9795C-5820-11E2-9CA7-94046188709B")]
    [ProgId("Frameworks.ExtSldWorks")]
    public class ExtSldWorks : IExtSldWorks
    {
        #region "Variables locales"
        private Debug _Debug = Debug.Instance;
        private Boolean _EstInitialise = false;

        private SldWorks _SwSW;
        private String _VersionDeBase;
        private String _VersionCourante;
        private String _Hotfixe;
        private String _Revision;
        private int Erreur = 0;
        private int Warning = 0;

        #endregion

        #region "Constructeur\Destructeur"

        public ExtSldWorks() { }

        #endregion

        #region "Propriétés"

        /// <summary>
        /// Initialisation de l'objet ExtSldWorks pour commencer
        /// </summary>
        public SldWorks SwSW { get { return _SwSW; } }

        /// <summary>
        /// Retourner le type du document actif
        /// </summary>
        public TypeFichier_e TypeDuModeleActif
        {
            get
            {
                ExtModele Modele = new ExtModele() ;
                Modele.Init(_SwSW.ActiveDoc(), this);
                return Modele.TypeDuModele;
            }
        }

        public String VersionDeBase { get { return _VersionDeBase; } }
        public String VersionCourante { get { return _VersionCourante; } }
        public String Hotfixe { get { return _Hotfixe; } }
        public String Revision { get { return _Revision; } }

        internal Boolean EstInitialise { get { return _EstInitialise; } }

        #endregion

        #region "Méthodes"

        public Boolean Init(SldWorks SldWks)
        {
            try
            {

                if (SldWks != null)
                {
                    _SwSW = SldWks;
                    _Debug.Init(this);
                    _Debug.DebugAjouterLigne(this.GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name);
                    /// A chaque initialisation de l'objet SW, on vide le debug et on inscrit la version de SW
                    /// Ca evite de chercher trop loin

                    _SwSW.GetBuildNumbers2(out _VersionDeBase, out _VersionCourante, out _Hotfixe);
                    _Revision = _SwSW.RevisionNumber();
                    _Debug.DebugAjouterLigne("    ");
                    _Debug.DebugAjouterLigne("================================================================================================");
                    _Debug.DebugAjouterLigne("SOLIDWORKS");
                    _Debug.DebugAjouterLigne("Version de base : " + VersionDeBase + "    Version courante : " + VersionCourante + "    Hotfixe : " + Hotfixe);
                    _Debug.DebugAjouterLigne("------------------------------------------------------------------------------------------------");
                    _EstInitialise = true;
                }

                return _EstInitialise;
            }
            catch (Exception ex)
            {
                _Debug.DebugAjouterLigne(ex.Source.ToString());
                return false;
            }
        }
        
        /// <summary>
        /// Renvoi le modele actif ou ouvre le fichier à partir du chemin passé en parametre.
        /// Le fichier est ouvert en mode silencieux, il ne devrait donc pas être visible.
        /// </summary>
        /// <param name="Chemin"></param>
        /// <returns></returns>
        public ExtModele Modele(String Chemin = "")
        {
            _Debug.DebugAjouterLigne(this.GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name);

            ExtModele pModele = new ExtModele();
            if (String.IsNullOrEmpty(Chemin))
            {
                _Debug.DebugAjouterLigne("\t -> " + "SldWorks.ActiveDoc");
                pModele.Init(_SwSW.ActiveDoc, this);
            }
            else
            {
                _Debug.DebugAjouterLigne("\t -> " + "Ouvrir " + Chemin);
                pModele.Init(Ouvrir(Chemin), this);
            }

            if (pModele.EstInitialise)
                return pModele;
            else
                return null;
        }

        /// <summary>
        /// Ouvre un fichier à partir de son chemin
        /// Verifie s'il est déjà ouvert, auquel cas ce dernier est renvoyé
        /// </summary>
        /// <param name="Chemin"></param>
        /// <returns></returns>
        private ModelDoc2 Ouvrir(String Chemin)
        {
            _Debug.DebugAjouterLigne(this.GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name);

            foreach (ModelDoc2 pSwModele in _SwSW.GetDocuments())
            {
                if (pSwModele.GetPathName() == Chemin)
                {
                    _Debug.DebugAjouterLigne("\t -> " + "Fichier déjà ouvert : " + Chemin);
                    return pSwModele;
                }
            }

            swDocumentTypes_e Type = 0;

            switch (Path.GetExtension(Chemin))
            {
                case ".SLDASM" :
                    Type = swDocumentTypes_e.swDocASSEMBLY;
                    break;
                case ".SLDPRT":
                    Type = swDocumentTypes_e.swDocPART;
                    break;
                case ".SLDDRW":
                    Type = swDocumentTypes_e.swDocDRAWING;
                    break;
                default:
                    return null;
            }

            _Debug.DebugAjouterLigne("\t -> " + "Ouvre le fichier : " + Chemin);

            return _SwSW.OpenDoc6(Chemin, (int)Type, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref Erreur, ref Warning);
        }

        #endregion
    }
}