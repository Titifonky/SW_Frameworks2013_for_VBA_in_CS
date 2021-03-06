﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Framework
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("5F43729A-1E72-48F7-AAA5-0474D859E79B")]
    public interface IDlgRechercherUnFichier
    {
        String Titre { get; set; }
        Boolean AjouterExtension { get; set; }
        Boolean TestFichierExiste { get; set; }
        Boolean TestCheminExiste { get; set; }
        String ExtensionParDefaut { get; set; }
        String Filtre { get; set; }
        int IndexDuFiltre { get; set; }
        String DossierInitial { get; set; }
        Boolean RestaurerLeDossierActif { get; set; }
        Boolean AfficherAide { get; set; }
        Boolean AfficherLectureSeule { get; set; }
        Boolean ExtensionsMultiple { get; set; }
        Boolean LectureSeuleSelectionne { get; set; }

        Boolean EstInitialise { get; }

        Boolean Init(eSldWorks Sw);

        void FiltreSW(TypeFichier_e TypeDesFichiersFiltres, Boolean FiltreDistinct = true);
        String SelectionnerUnFichier(Boolean CheminComplet = true);
        eFichierSW SelectionnerUnFichierSW();
        ArrayList SelectionnerPlusieursFichiers(Boolean CheminComplet = true);
        ArrayList SelectionnerPlusieursFichierSW();
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("736A58F3-6B78-4826-A919-82E817206398")]
    [ProgId("Frameworks.DlgRechercherUnFichier")]
    public class DlgRechercherUnFichier : IDlgRechercherUnFichier, IDisposable
    {
        #region "Variables locales"

        private static readonly String cNOMCLASSE = typeof(DlgRechercherUnFichier).Name;

        private Boolean _EstInitialise = false;

        private eSldWorks _SW = null;

        #endregion

        #region "Constructeur\Destructeur"

        public DlgRechercherUnFichier() { }

        #endregion

        #region "Propriétés"

        private OpenFileDialog _Dialogue = new OpenFileDialog();

        public String Titre { get { return _Dialogue.Title; } set { _Dialogue.Title = value; } }
        public Boolean AjouterExtension { get { return _Dialogue.AddExtension; } set { _Dialogue.AddExtension = value; } }
        public Boolean TestFichierExiste { get { return _Dialogue.CheckFileExists; } set { _Dialogue.CheckFileExists = value; } }
        public Boolean TestCheminExiste { get { return _Dialogue.CheckPathExists; } set { _Dialogue.CheckPathExists = value; } }
        public String ExtensionParDefaut { get { return _Dialogue.DefaultExt; } set { _Dialogue.DefaultExt = value; } }
        public String Filtre { get { return _Dialogue.Filter; } set { _Dialogue.Filter = value; } }
        public int IndexDuFiltre { get { return _Dialogue.FilterIndex; } set { _Dialogue.FilterIndex = value; } }
        public String DossierInitial { get { return _Dialogue.InitialDirectory; } set { _Dialogue.InitialDirectory = value; } }
        public Boolean RestaurerLeDossierActif { get { return _Dialogue.RestoreDirectory; } set { _Dialogue.RestoreDirectory = value; } }
        public Boolean AfficherAide { get { return _Dialogue.ShowHelp; } set { _Dialogue.ShowHelp = value; } }
        public Boolean AfficherLectureSeule { get { return _Dialogue.ShowReadOnly; } set { _Dialogue.ShowReadOnly = value; } }
        public Boolean ExtensionsMultiple { get { return _Dialogue.SupportMultiDottedExtensions; } set { _Dialogue.SupportMultiDottedExtensions = value; } }
        public Boolean LectureSeuleSelectionne { get { return _Dialogue.ReadOnlyChecked; } set { _Dialogue.ReadOnlyChecked = value; } }

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
        /// Filtrer les fichiers SW
        /// </summary>
        /// <param name="TypeDesFichiersFiltres"></param>
        public void FiltreSW(TypeFichier_e TypeDesFichiersFiltres, Boolean FiltreDistinct = true)
        {
            Log.Methode(cNOMCLASSE);

            String TxtFiltre;

            List<String> Filtre = new List<string>();

            if (FiltreDistinct)
            {
                foreach (TypeFichier_e T in Enum.GetValues(typeof(TypeFichier_e)))
                {
                    if (TypeDesFichiersFiltres.HasFlag(T))
                        Filtre.Add(CONSTANTES.InfoFichier(T, InfoFichier_e.cNom) + " (*" + CONSTANTES.InfoFichier(T) + "|*" + CONSTANTES.InfoFichier(T));
                }

                TxtFiltre = String.Join("|", Filtre);
            }
            else
            {
                foreach (TypeFichier_e T in Enum.GetValues(typeof(TypeFichier_e)))
                {
                    if (TypeDesFichiersFiltres.HasFlag(T))
                        Filtre.Add("*" + CONSTANTES.InfoFichier(T));
                }

                TxtFiltre = "Fichier SolidWorks (" + String.Join(", ", Filtre) + ")" + "|" + String.Join("; ", Filtre);
            }

            _Dialogue.Filter = TxtFiltre;
        }

        /// <summary>
        /// Selectionner un fichier et renvoi le chemin
        /// </summary>
        /// <param name="CheminComplet"></param>
        /// <returns></returns>
        public String SelectionnerUnFichier(Boolean CheminComplet = true)
        {
            Log.Methode(cNOMCLASSE);

            _Dialogue.Multiselect = false;

            if (_Dialogue.ShowDialog() == DialogResult.OK)
            {
                if (CheminComplet)
                    return _Dialogue.FileName;
                else
                    return _Dialogue.SafeFileName;
            }
            else
                return "";
        }

        /// <summary>
        /// Selectionner un fichier et renvoi l'objet FichierSW
        /// </summary>
        /// <param name="CheminComplet"></param>
        /// <returns></returns>
        public eFichierSW SelectionnerUnFichierSW()
        {
            Log.Methode(cNOMCLASSE);

            eFichierSW pFichierSW = new eFichierSW();
            if (pFichierSW.Init(_SW))
            {
                pFichierSW.Chemin = SelectionnerUnFichier(true);
                if ((pFichierSW.TypeDuFichier == TypeFichier_e.cAssemblage) || (pFichierSW.TypeDuFichier == TypeFichier_e.cPiece) || (pFichierSW.TypeDuFichier == TypeFichier_e.cDessin))
                    return pFichierSW;
            }

            return null;
        }

        /// <summary>
        /// Selectionner des fichiers et renvoi un tableau de chemins
        /// </summary>
        /// <param name="CheminComplet"></param>
        /// <returns></returns>
        public ArrayList SelectionnerPlusieursFichiers(Boolean CheminComplet = true)
        {
            Log.Methode(cNOMCLASSE);

            _Dialogue.Multiselect = true;
            ArrayList pArrayFichiers = new ArrayList();

            if (_Dialogue.ShowDialog() == DialogResult.OK)
            {
                if (CheminComplet)
                    pArrayFichiers = new ArrayList(_Dialogue.FileNames);
                else
                    pArrayFichiers = new ArrayList(_Dialogue.SafeFileNames);
            }

            return pArrayFichiers;
        }

        /// <summary>
        /// Selectionner des fichiers et renvoi un tableau d'objet FichierSW
        /// </summary>
        /// <param name="CheminComplet"></param>
        /// <returns></returns>
        public ArrayList SelectionnerPlusieursFichierSW()
        {
            Log.Methode(cNOMCLASSE);

            ArrayList pArrayFichiers = new ArrayList();

            foreach (String CheminFichier in SelectionnerPlusieursFichiers(true))
            {
                eFichierSW pFichierSW = new eFichierSW();
                if (pFichierSW.Init(_SW))
                {
                    pFichierSW.Chemin = CheminFichier;
                    if ((pFichierSW.TypeDuFichier == TypeFichier_e.cAssemblage) || (pFichierSW.TypeDuFichier == TypeFichier_e.cPiece) || (pFichierSW.TypeDuFichier == TypeFichier_e.cDessin))
                        pArrayFichiers.Add(pFichierSW);
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
