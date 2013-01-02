using System;
using System.IO;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace Frameworks2013
{

    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("01A31113-9353-44cc-A1F4-C6F1210E4B30")]
    public interface IExtSldWorks
    {
        SldWorks swSW { get; set; }
        TypeFichier_e TypeDuModeleActif { get; }
        ExtConstantes Constantes();
        ExtDebug Debug();
        ExtModele Modele(String Chemin = "");
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("E2F07CD4-CE73-4102-B35D-119362624C47")]
    [ProgId("Frameworks.ExtSldWorks")]
    public class ExtSldWorks : IExtSldWorks
    {
        #region "Variables locales"

        private SldWorks _swSW;
        private ExtConstantes _Constantes = new ExtConstantes();
        private ExtDebug _Debug = new ExtDebug();
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
            set
            {
                _swSW = value;

                /// A chaque initialisation de l'objet SW, on vide le debug et on inscrit la version de SW
                /// Ca evite de chercher trop loin

                _Debug.Effacer();

                String VersionDeBase;
                String VersionCourante;
                String Hotfixe;

                _swSW.GetBuildNumbers2(out VersionDeBase, out VersionCourante, out Hotfixe);
                _Debug.AjouterLigne("    ");
                _Debug.AjouterLigne("================================================================================================");
                _Debug.AjouterLigne("SOLIDWORKS");
                _Debug.AjouterLigne("Version de base : " + VersionDeBase + "    Version courante : " + VersionCourante + "    Hotfixe : " + Hotfixe);
                _Debug.AjouterLigne("------------------------------------------------------------------------------------------------");
            }
        }

        /// <summary>
        /// Retourner le type du document actif
        /// </summary>
        public TypeFichier_e TypeDuModeleActif
        {
            get
            {
                ModelDoc2 Modele = _swSW.ActiveDoc() ;
                TypeFichier_e Type;

                switch (Modele.GetType())
                {
                    case (int)swDocumentTypes_e.swDocASSEMBLY :
                        Type = TypeFichier_e.cAssemblage ;
                        break;

                    case (int)swDocumentTypes_e.swDocPART:
                        Type = TypeFichier_e.cPiece;
                        break;

                    case (int)swDocumentTypes_e.swDocDRAWING:
                        Type = TypeFichier_e.cDessin;
                        break;

                    default :
                        Type = TypeFichier_e.cAucun;
                        break;
                }

                Modele = null;
                return Type;
            }
        }

        #endregion

        #region "Méthodes"
        
        /// <summary>
        /// Renvoi le modele actif ou ouvre le fichier à partir du chemin passé en parametre.
        /// Le fichier est ouvert en mode silencieux, il ne devrait donc pas être visible.
        /// </summary>
        /// <param name="Chemin"></param>
        /// <returns></returns>
        public ExtModele Modele(String Chemin = "")
        {
            ExtModele pModele = new ExtModele();
            pModele.SW = this;

            if (String.IsNullOrEmpty(Chemin))
                pModele.swModele = _swSW.ActiveDoc;
            else
                pModele.swModele = Ouvrir(Chemin);

            return pModele;
        }

        /// <summary>
        /// Ouvre un fichier à partir de son chemin
        /// </summary>
        /// <param name="Chemin"></param>
        /// <returns></returns>
        private ModelDoc2 Ouvrir(String Chemin)
        {
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

        /// <summary>
        /// Retourne l'objet contenant les constantes
        /// </summary>
        /// <returns></returns>
        public ExtConstantes Constantes()
        {
            return _Constantes;
        }

        /// <summary>
        /// Retourne l'objet de debug
        /// </summary>
        /// <returns></returns>
        public ExtDebug Debug()
        {
            return _Debug;
        }

        #endregion
    }
}