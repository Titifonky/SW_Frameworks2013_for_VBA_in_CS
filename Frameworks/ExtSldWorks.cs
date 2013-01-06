using System;
using System.IO;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace Framework2013
{

    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("9ED6BE92-5820-11E2-9D5D-93046188709B")]
    public interface IExtSldWorks
    {
        SldWorks swSW { get; }
        ExtConstantes Constantes { get; }
        ExtDebug Debug { get; }
        TypeFichier_e TypeDuModeleActif { get; }
        String VersionDeBase { get; }
        String VersionCourante { get; }
        String Hotfixe { get; }
        Boolean Init(SldWorks SldWks);
        ExtModele Modele(String Chemin = "");
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("A2E9795C-5820-11E2-9CA7-94046188709B")]
    [ProgId("Frameworks.ExtSldWorks")]
    public class ExtSldWorks : IExtSldWorks
    {
        #region "Variables locales"

        private SldWorks _swSW;
        private ExtConstantes _Constantes = new ExtConstantes();
        private ExtDebug _Debug = new ExtDebug();
        private String _VersionDeBase;
        private String _VersionCourante;
        private String _Hotfixe;
        private int Erreur = 0;
        private int Warning = 0;

        #endregion

        #region "Constructeur\Destructeur"

        public ExtSldWorks()
        {
        }

        ~ExtSldWorks()
        {
            _swSW = null;
            _Constantes = null;
            _Debug = null;
        }

        #endregion

        #region "Propriétés"

        /// <summary>
        /// Initialisation de l'objet ExtSldWorks pour commencer
        /// </summary>
        public SldWorks swSW
        {
            get { return _swSW; }
        }

        /// <summary>
        /// Retourne l'objet contenant les constantes
        /// </summary>
        public ExtConstantes Constantes
        {
            get { return _Constantes; }
        }

        /// <summary>
        /// Retourne l'objet de debug
        /// </summary>
        public ExtDebug Debug
        {
            get { return _Debug; }
        }

        /// <summary>
        /// Retourner le type du document actif
        /// </summary>
        public TypeFichier_e TypeDuModeleActif
        {
            get
            {
                ExtModele Modele = new ExtModele() ;
                Modele.Init(_swSW.ActiveDoc(), this);
                return Modele.TypeDuModele;
            }
        }

        public String VersionDeBase
        {
            get { return _VersionDeBase; }
        }

        public String VersionCourante
        {
            get { return _VersionCourante; }
        }

        public String Hotfixe
        {
            get { return _Hotfixe; }
        }

        #endregion

        #region "Méthodes"

        public Boolean Init(SldWorks SldWks)
        {
            _MethodBase Methode = System.Reflection.MethodBase.GetCurrentMethod();

            if (!(SldWks.Equals(null)))
            {
                _swSW = SldWks;
                _Debug.Init(this);
                _Debug.ErreurAjouterLigne(this.GetType().Name + "." + Methode.Name);
                /// A chaque initialisation de l'objet SW, on vide le debug et on inscrit la version de SW
                /// Ca evite de chercher trop loin

                _swSW.GetBuildNumbers2(out _VersionDeBase, out _VersionCourante, out _Hotfixe);
                _Debug.ExecutionAjouterLigne("    ");
                _Debug.ExecutionAjouterLigne("================================================================================================");
                _Debug.ExecutionAjouterLigne("SOLIDWORKS");
                _Debug.ExecutionAjouterLigne("Version de base : " + VersionDeBase + "    Version courante : " + VersionCourante + "    Hotfixe : " + Hotfixe);
                _Debug.ExecutionAjouterLigne("------------------------------------------------------------------------------------------------");
                return true;
            }

            
            return false;
            
        }
        
        /// <summary>
        /// Renvoi le modele actif ou ouvre le fichier à partir du chemin passé en parametre.
        /// Le fichier est ouvert en mode silencieux, il ne devrait donc pas être visible.
        /// </summary>
        /// <param name="Chemin"></param>
        /// <returns></returns>
        public ExtModele Modele(String Chemin = "")
        {
            _MethodBase Methode = System.Reflection.MethodBase.GetCurrentMethod();
            _Debug.ErreurAjouterLigne(this.GetType().Name + "." + Methode.Name);

            ExtModele pModele = new ExtModele();
            if (String.IsNullOrEmpty(Chemin))
            {
                _Debug.ErreurAjouterLigne("\t" + this.GetType().Name + " -> " + "SldWorks.ActiveDoc");
                pModele.Init(_swSW.ActiveDoc, this);
            }
            else
            {
                _Debug.ErreurAjouterLigne("\t" + this.GetType().Name + " -> " + "Ouvrir");
                pModele.Init(Ouvrir(Chemin), this);
            }

            return pModele;
        }

        /// <summary>
        /// Ouvre un fichier à partir de son chemin
        /// Verifie s'il est déjà ouvert, auquel cas ce dernier est renvoyé
        /// </summary>
        /// <param name="Chemin"></param>
        /// <returns></returns>
        private ModelDoc2 Ouvrir(String Chemin)
        {
            _MethodBase Methode = System.Reflection.MethodBase.GetCurrentMethod();
            _Debug.ErreurAjouterLigne(this.GetType().Name + "." + Methode.Name);

            foreach (ModelDoc2 Modele in _swSW.GetDocuments())
            {
                if (Modele.GetPathName() == Chemin)
                    return Modele;
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

            return _swSW.OpenDoc6(Chemin, (int)Type, (int)swOpenDocOptions_e.swOpenDocOptions_Silent,"", ref Erreur, ref Warning);
        }

        #endregion
    }
}