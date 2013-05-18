using System;
using System.IO;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Reflection;

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
        Boolean ActiverDebug { get; set; }
        Boolean Init(SldWorks SldWks);
        eModele Modele(String Chemin = "");
        eModele ModeleEnCoursEdition();
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("A2E9795C-5820-11E2-9CA7-94046188709B")]
    [ProgId("Frameworks.ExtSldWorks")]
    public class ExtSldWorks : IExtSldWorks
    {
        #region "Variables locales"
        
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
        /// Renvoi l'objet SldWorks associé.
        /// </summary>
        public SldWorks SwSW { get { Debug.Info(MethodBase.GetCurrentMethod()); return _SwSW; } }

        /// <summary>
        /// Retourner le type du document actif.
        /// </summary>
        public TypeFichier_e TypeDuModeleActif
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                eModele Modele = new eModele();
                Modele.Init(_SwSW.ActiveDoc(), this);
                return Modele.TypeDuModele;
            }
        }

        /// <summary>
        /// Retourne le numero de la version de base.
        /// </summary>
        public String VersionDeBase { get { return _VersionDeBase; } }
        
        /// <summary>
        /// Retourne le numero de version courant.
        /// </summary>
        public String VersionCourante { get { return _VersionCourante; } }

        /// <summary>
        /// Retourne le numero du hotfixe.
        /// </summary>
        public String Hotfixe { get { return _Hotfixe; } }

        /// <summary>
        /// Retourne le numero de la révision.
        /// </summary>
        public String Revision { get { return _Revision; } }

        /// <summary>
        /// Activer ou désactiver l'ecriture dans le fichier Debug.
        /// </summary>
        public Boolean ActiverDebug { get { return Debug.Actif; } set { Debug.Actif = value; } }

        /// <summary>
        /// Fonction interne.
        /// Test l'initialisation de l'objet ExtSldWorks.
        /// </summary>
        internal Boolean EstInitialise { get { Debug.Info(MethodBase.GetCurrentMethod()); return _EstInitialise; } }

        #endregion

        #region "Méthodes"

        /// <summary>
        /// Initialiser l'objet ExtSldWorks.
        /// </summary>
        /// <param name="SldWks"></param>
        /// <returns></returns>
        public Boolean Init(SldWorks SldWks)
        {
            try
            {

                if (SldWks != null)
                {
                    _SwSW = SldWks;
                    Debug.Init(SldWks);
                    Debug.Info(MethodBase.GetCurrentMethod());
                    SldWks.GetBuildNumbers2(out _VersionDeBase, out _VersionCourante, out _Hotfixe);
                    _Revision = SldWks.RevisionNumber();
                    _EstInitialise = true;
                }

                return _EstInitialise;
            }
            catch (Exception ex)
            {
                Debug.Info(ex.Source.ToString());
                return false;
            }
        }

        /// <summary>
        /// Renvoi le modele actif ou ouvre le fichier à partir du chemin passé en parametre.
        /// Le fichier est ouvert en mode silencieux, il ne devrait donc pas être visible.
        /// </summary>
        /// <param name="Chemin"></param>
        /// <returns></returns>
        public eModele Modele(String Chemin = "")
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            eModele pModele = new eModele();
            if (String.IsNullOrEmpty(Chemin))
            {
                Debug.Info("Document actif");
                pModele.Init(_SwSW.ActiveDoc, this);
            }
            else
            {
                Debug.Info("Ouvrir " + Chemin);
                pModele.Init(Ouvrir(Chemin), this);
            }

            if (pModele.EstInitialise)
                return pModele;
            else
                return null;
        }

        /// <summary>
        /// Renvoi le modele en cours d'edition
        /// </summary>
        /// <returns></returns>
        public eModele ModeleEnCoursEdition()
        {
            Debug.Info(MethodBase.GetCurrentMethod());
            eModele pModeleActif = this.Modele();
            eModele pModeleEdite = new eModele();
            if (pModeleActif.EstInitialise && (pModeleActif.TypeDuModele == TypeFichier_e.cAssemblage))
            {
                if (pModeleEdite.Init(pModeleActif.Assemblage.SwAssemblage.GetEditTarget(), this))
                {
                    ExtComposant pComposant = new ExtComposant();
                    if (pComposant.Init(pModeleActif.Assemblage.SwAssemblage.GetEditTargetComponent(), pModeleEdite))
                    {
                        pModeleEdite.Composant = pComposant;
                        return pModeleEdite;
                    }

                }
            }

            return pModeleActif;
        }

        /// <summary>
        /// Ouvre un fichier à partir de son chemin.
        /// Verifie s'il est déjà ouvert, auquel cas ce dernier est renvoyé.
        /// </summary>
        /// <param name="Chemin"></param>
        /// <returns></returns>
        private ModelDoc2 Ouvrir(String Chemin)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            if (_SwSW.GetDocumentCount() > 0)
            {
                foreach (ModelDoc2 pSwModele in _SwSW.GetDocuments())
                {
                    if (pSwModele.GetPathName() == Chemin)
                    {
                        Debug.Info("Fichier déjà ouvert : " + Chemin);
                        return pSwModele;
                    }
                }
            }

            swDocumentTypes_e Type = 0;

            switch (Path.GetExtension(Chemin))
            {
                case ".SLDASM":
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

            Debug.Info("Ouvre le fichier : " + Chemin);

            return _SwSW.OpenDoc6(Chemin, (int)Type, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref Erreur, ref Warning);
        }

        #endregion
    }
}