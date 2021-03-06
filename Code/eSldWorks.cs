﻿using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.IO;
using System.Runtime.InteropServices;

namespace Framework
{

    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("9ED6BE92-5820-11E2-9D5D-93046188709B")]
    public interface IeSldWorks
    {
        SldWorks SwSW { get; }
        TypeFichier_e TypeDuModeleActif { get; }
        String VersionDeBase { get; }
        String VersionCourante { get; }
        String Hotfixe { get; }
        String Revision { get; }
        Boolean ActiverLog { get; set; }
        Boolean ActiverLesConfigurations { get; set; }
        eGestOptions GestOptions { get; }
        void EcrireLog(String Message = "");
        Boolean Init(SldWorks SldWks);
        eModele Modele(String Chemin = "");
        eModele ModeleEnCoursEdition();
        eModele CreerDocument(String Dossier, String NomDuDocument, TypeFichier_e TypeDeDocument, String Gabarit = "");
        Boolean EstInitialise { get; }
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("A2E9795C-5820-11E2-9CA7-94046188709B")]
    [ProgId("Frameworks.eSldWorks")]
    public class eSldWorks : IeSldWorks
    {
        #region "Variables locales"

        private static readonly String cNOMCLASSE = typeof(eSldWorks).Name;

        private Boolean _EstInitialise = false;

        private SldWorks _SwSW = null;
        private String _VersionDeBase = "";
        private String _VersionCourante = "";
        private String _Hotfixe = "";
        private String _Revision = "";
        private int Erreur = 0;
        private int Warning = 0;
        private eGestOptions _GestOptions = null;

        private Boolean _ActiverLesConfigurations = false;

        #endregion

        #region "Constructeur\Destructeur"

        public eSldWorks() { }

        ~eSldWorks()
        {
            Log.Stopper();
        }

        #endregion

        #region "Propriétés"

        /// <summary>
        /// Renvoi l'objet SldWorks associé.
        /// </summary>
        public SldWorks SwSW
        {
            get
            {
                Log.Methode(cNOMCLASSE);
                return _SwSW;
            }
        }

        /// <summary>
        /// Retourner le type du document actif.
        /// </summary>
        public TypeFichier_e TypeDuModeleActif
        {
            get
            {
                Log.Methode(cNOMCLASSE);
                return Modele().TypeDuModele;
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

        public Boolean ActiverLog
        {
            get
            {
                return Log.Activer;
            }
            set
            {
                Log.Activer = value;
            }
        }

        public Boolean ActiverLesConfigurations { get { return _ActiverLesConfigurations; } set { _ActiverLesConfigurations = value; } }

        /// <summary>
        /// Retourne le gestionnaire d'options
        /// </summary>
        public eGestOptions GestOptions
        {
            get
            {
                Log.Methode(cNOMCLASSE);

                if (_GestOptions == null)
                {
                    _GestOptions = new eGestOptions();
                    _GestOptions.Init(this);
                }

                if (_GestOptions.EstInitialise)
                    return _GestOptions;

                return null;
            }
        }

        /// <summary>
        /// Fonction interne.
        /// Test l'initialisation de l'objet ExtSldWorks.
        /// </summary>
        public Boolean EstInitialise
        {
            get
            {
                Log.Methode(cNOMCLASSE);
                return _EstInitialise;
            }
        }

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
                    Log.Entete();
                    Log.Methode(cNOMCLASSE);
                    Log.Activer = false;
                    _SwSW.GetBuildNumbers2(out _VersionDeBase, out _VersionCourante, out _Hotfixe);
                    _Revision = SldWks.RevisionNumber();
                    _EstInitialise = true;
                }

                return _EstInitialise;
            }
            catch (Exception ex)
            {
                Log.Message(ex.Source.ToString());
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
            Log.Methode(cNOMCLASSE);

            eModele pModele = new eModele();
            if (String.IsNullOrEmpty(Chemin))
            {
                Log.Message("Document actif");
                ModelDoc2 pModeleActif = _SwSW.ActiveDoc;
                pModele.Init(pModeleActif, this);
                pModele.ReinitialiserComposant();
            }
            else
            {
                Log.Message("Ouvrir " + Chemin);
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
            Log.Methode(cNOMCLASSE);
            eModele pModeleActif = this.Modele();
            eModele pModeleEdite = new eModele();
            if (pModeleActif.EstInitialise && (pModeleActif.TypeDuModele == TypeFichier_e.cAssemblage))
            {
                if (pModeleEdite.Init(pModeleActif.Assemblage.SwAssemblage.GetEditTarget(), this))
                {
                    eComposant pComposant = new eComposant();
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
            Log.Methode(cNOMCLASSE);

            if (_SwSW.GetDocumentCount() > 0)
            {
                foreach (ModelDoc2 pSwModele in _SwSW.GetDocuments())
                {
                    if (pSwModele.GetPathName() == Chemin)
                    {
                        Log.Message("Fichier déjà ouvert : " + Chemin);
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

            Log.Message("Ouvre le fichier : " + Chemin);

            return _SwSW.OpenDoc6(Chemin, (int)Type, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref Erreur, ref Warning);
        }

        /// <summary>
        /// Creer un document
        /// </summary>
        /// <param name="Dossier"></param>
        /// <param name="NomDuDocument"></param>
        /// <param name="TypeDeDocument"></param>
        /// <param name="Gabarit"></param>
        /// <returns></returns>
        public eModele CreerDocument(String Dossier, String NomDuDocument, TypeFichier_e TypeDeDocument, String Gabarit = "")
        {
            eModele pModele = new eModele();
            ModelDoc2 pSwModele;

            String pCheminGabarit = "";

            if (String.IsNullOrEmpty(Gabarit))
            {
                switch (TypeDeDocument)
                {
                    case TypeFichier_e.cAssemblage:
                        pCheminGabarit = _SwSW.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplateAssembly);
                        break;
                    case TypeFichier_e.cPiece:
                        pCheminGabarit = _SwSW.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
                        break;
                    case TypeFichier_e.cDessin:
                        pCheminGabarit = _SwSW.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplateDrawing);
                        break;
                }
            }
            else
            {
                String[] pTabCheminsGabarit = _SwSW.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swFileLocationsDocumentTemplates).Split(';');

                foreach (String Chemin in pTabCheminsGabarit)
                {
                    pCheminGabarit = Chemin + @"\" + Gabarit + CONSTANTES.InfoFichier(TypeDeDocument, InfoFichier_e.cGabarit);
                    if (File.Exists(pCheminGabarit))
                        break;
                }
            }

            int Format = 0;
            Double Lg = 0;
            Double Ht = 0;

            if (TypeDeDocument == TypeFichier_e.cDessin)
            {
                Double[] pTab = _SwSW.GetTemplateSizes(pCheminGabarit);
                Format = (int)pTab[0];
                Lg = pTab[1];
                Ht = pTab[2];
            }

            Log.Message(Format.ToString());
            Log.Message(Lg.ToString());
            Log.Message(Ht.ToString());

            pSwModele = _SwSW.NewDocument(pCheminGabarit, Format, Lg, Ht);
            pSwModele.Extension.SaveAs(Dossier + @"\" + NomDuDocument + CONSTANTES.InfoFichier(TypeDeDocument),
                                        (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                                        (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                                        null,
                                        ref Erreur,
                                        ref Warning);

            if (pModele.Init(pSwModele, this))
                return pModele;

            return null;
        }

        public void EcrireLog(String Message = "")
        {
            Log.Write("EXTERNE ====> " + Message);
        }

        #endregion
    }
}