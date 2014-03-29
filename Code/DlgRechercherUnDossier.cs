using System;
using System.Collections;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Framework
{

    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("64A83C8C-997C-4266-83EC-E6B99A58FE94")]
    public interface IDlgRechercherUnDossier
    {
        String Description { get; set; }
        Environment.SpecialFolder DossierRacine { get; set; }
        Boolean BoutonNouveauDossier { get; set; }
        Boolean EstInitialise { get; }

        Boolean Init(eSldWorks Sw);

        String SelectionnerUnDossier();
        ArrayList SelectionnerUnDossierEtRenvoyerFichierSW(TypeFichier_e TypeDesFichiers, String NomARechercher = "*", Boolean ParcourirLesSousDossier = false);
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("BADDD2D2-9D97-4D6B-8F49-C9BC161E3E06")]
    [ProgId("Frameworks.DlgRechercherUnDossier")]
    public class DlgRechercherUnDossier : IDlgRechercherUnDossier, IDisposable
    {
        #region "Variables locales"

        private static readonly String cNOMCLASSE = typeof(DlgRechercherUnDossier).Name;

        private Boolean _EstInitialise = false;

        private eSldWorks _SW = null;

        #endregion

        #region "Constructeur\Destructeur"

        public DlgRechercherUnDossier() { }

        #endregion

        #region "Propriétés"

        private FolderBrowserDialog _Dialogue = new FolderBrowserDialog();

        public String Description { get { return _Dialogue.Description; } set { _Dialogue.Description = value; } }
        public Environment.SpecialFolder DossierRacine { get { return _Dialogue.RootFolder; } set { _Dialogue.RootFolder = value; } }
        public Boolean BoutonNouveauDossier { get { return _Dialogue.ShowNewFolderButton; } set { _Dialogue.ShowNewFolderButton = value; } }

        public Boolean EstInitialise { get { Log.Methode(cNOMCLASSE); return _EstInitialise; } }

        #endregion

        #region "Méthodes"

        /// <summary>
        /// Initialiser l'objet.
        /// </summary>
        /// <param name="Sw"></param>
        /// <returns></returns>
        public Boolean Init(eSldWorks Sw)
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

        /// <summary>
        /// Retourne le chemin du dossier sélectionné
        /// </summary>
        /// <returns></returns>
        public String SelectionnerUnDossier()
        {
            if (_Dialogue.ShowDialog() == DialogResult.OK)
                return _Dialogue.SelectedPath;

            return "";
        }

        /// <summary>
        /// Retourne la liste des fichiers du dossier sélectionné
        /// </summary>
        /// <param name="TypeDesFichiers"></param>
        /// <param name="NomARechercher"></param>
        /// <param name="ParcourirLesSousDossier"></param>
        /// <returns></returns>
        public ArrayList SelectionnerUnDossierEtRenvoyerFichierSW(TypeFichier_e TypeDesFichiers, String NomARechercher = "*", Boolean ParcourirLesSousDossier = false)
        {
            ArrayList pArrayFichiers = new ArrayList();

            String CheminDossier = SelectionnerUnDossier();

            SearchOption Options = SearchOption.TopDirectoryOnly;
            if (ParcourirLesSousDossier)
                Options = SearchOption.AllDirectories;

            if (Directory.Exists(CheminDossier))
            {
                foreach (String CheminFichier in Directory.EnumerateFiles(CheminDossier, NomARechercher, Options))
                {
                    eFichierSW pFichierSW = new eFichierSW();
                    if (pFichierSW.Init(_SW))
                    {
                        pFichierSW.Chemin = CheminFichier;
                        if (TypeDesFichiers.HasFlag(pFichierSW.TypeDuFichier))
                            pArrayFichiers.Add(pFichierSW);
                    }
                }
            }

            return pArrayFichiers;
        }

        #endregion

        void IDisposable.Dispose()
        {
            _Dialogue.Dispose();
        }
    }
}
