﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace Framework_SW2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("BF2ED17A-5820-11E2-8160-9F046188709B")]
    public interface IeComposant
    {
        Component2 SwComposant { get; }
        eModele Modele { get; }
<<<<<<< HEAD
<<<<<<< HEAD
<<<<<<< HEAD
        eConfiguration Configuration { get;}
=======
        eConfiguration Configuration { get; }
>>>>>>> 305ac273011a596c941cdda96677e845e4d8fd03
=======
        eConfiguration Configuration { get; }
>>>>>>> 305ac273011a596c941cdda96677e845e4d8fd03
=======
        eConfiguration Configuration { get; }
>>>>>>> 305ac273011a596c941cdda96677e845e4d8fd03
        String Nom { get; }
        Boolean EstExclu { get; set; }
        Boolean EstSupprime { get; set; }
        Boolean EstVisible { get; set; }
        Boolean EstFixe { get; set; }
        Boolean EstUneRepetition { get; }
        Boolean EstUneSymetrie { get; }
        Boolean EstCharge { get; }
        int NoOccurence { get; }
        int Nb { get; }
        TypeFichier_e TypeDuModele { get; }
        eRecherche NouvelleRecherche { get; }
        eRepere Repere { get; }
        eComposant Parent { get; }
        void Selectionner(Boolean Ajouter = true);
        void DeSelectionner();
        ArrayList ComposantsEnfants(String NomComposant = "", Boolean PrendreEnCompteSupprime = false);
        ArrayList ListeDesCorps(TypeCorps_e TypeDeCorps = TypeCorps_e.cTous);
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("C46318AE-5820-11E2-A863-A3046188709B")]
    [ProgId("Frameworks.eComposant")]
    public class eComposant : IeComposant, IComparable<eComposant>, IComparer<eComposant>, IEquatable<eComposant>
    {
#region "Variables locales"

        private Boolean _EstInitialise = false;

        private Component2 _SwComposant = null;
        private eModele _Modele = null;
        private eConfiguration _Configuration = null;
        private int _Nb = 0;

#endregion

#region "Constructeur\Destructeur"

        public eComposant() { }

#endregion

#region "Propriétés"

        /// <summary>
        /// Retourne l'objet Component2 associé.
        /// </summary>
        public Component2 SwComposant { get { Debug.Info(MethodBase.GetCurrentMethod()); return _SwComposant; } }

        /// <summary>
        /// Retourne le modele ExtModele associé.
        /// Active la configuration du composant si celle-ci ne l'est pas
        /// Cela permet de récupérer les corps et les fonctions de la configuration
        /// </summary>
        public eModele Modele
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                if (!_Modele.GestDeConfigurations.ConfigurationActive.Equals(_Configuration))
                    _Configuration.Activer();

                return _Modele;
            }
        }

        /// <summary>
        /// Retourne le type du modele.
        /// </summary>
        public TypeFichier_e TypeDuModele
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                ModelDoc2 pSwModele = _SwComposant.GetModelDoc2();

                switch (pSwModele.GetType())
                {
                    case (int)swDocumentTypes_e.swDocASSEMBLY:
                        return TypeFichier_e.cAssemblage;

                    case (int)swDocumentTypes_e.swDocPART:
                        return TypeFichier_e.cPiece;

                    case (int)swDocumentTypes_e.swDocDRAWING:
                        return TypeFichier_e.cDessin;

                    default:
                        return 0;
                }
            }
        }

        /// <summary>
        /// Retourne la configuration ExtConfiguration associée.
        /// </summary>
        public eConfiguration Configuration
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                return _Configuration;
            }
            internal set
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                if (value.Modele.Equals(_Modele))
                {
                    _Configuration = value;
                    _SwComposant.ReferencedConfiguration = _Configuration.Nom;
                }
            }
        }
        
        /// <summary>
        /// Retourne le nom du composant tel qu'il est dans l'arbre de création.
        /// </summary>
        public String Nom { get { Debug.Info(MethodBase.GetCurrentMethod()); return _SwComposant.Name2; } }
        
        /// <summary>
        /// Retourne le nonbre de composant.
        /// </summary>
        public int Nb { get { Debug.Info(MethodBase.GetCurrentMethod()); return _Nb; } internal set { Debug.Info(MethodBase.GetCurrentMethod()); _Nb = value; } }

        /// <summary>
        /// Retourne ou défini si le composant est exclu de la nomenclature.
        /// </summary>
        public Boolean EstExclu
        {
            get { Debug.Info(MethodBase.GetCurrentMethod()); return Convert.ToBoolean(SwComposant.ExcludeFromBOM); }
            set { Debug.Info(MethodBase.GetCurrentMethod()); SwComposant.ExcludeFromBOM = value; }
        }

        /// <summary>
        /// Retourne ou défini si le composant est supprimé.
        /// </summary>
        public Boolean EstSupprime
        {
            get { Debug.Info(MethodBase.GetCurrentMethod()); return Convert.ToBoolean(SwComposant.IsSuppressed()); }
            set
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                
                if (value == true)
                    SwComposant.SetSuppression2((int)swComponentSuppressionState_e.swComponentSuppressed);
                else
                    SwComposant.SetSuppression2((int)swComponentSuppressionState_e.swComponentResolved);
            }
        }

        /// <summary>
        /// Retourne ou défini si le composant est visible.
        /// </summary>
        public Boolean EstVisible
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                return Convert.ToBoolean(_SwComposant.Visible);
            }
            set
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                if (value)
                    _SwComposant.Visible = (int)swComponentVisibilityState_e.swComponentVisible;
                else
                    _SwComposant.Visible = (int)swComponentVisibilityState_e.swComponentHidden;
            }
        }

        /// <summary>
        /// Retourne ou défini si le composant est fixé.
        /// </summary>
        public Boolean EstFixe
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                return _SwComposant.IsFixed();
            }
            set
            {
                Debug.Info(MethodBase.GetCurrentMethod());

                if (_Modele.SW.TypeDuModeleActif == TypeFichier_e.cAssemblage)
                {
                    Selectionner(false);
                    if (value)
                        _Modele.SW.Modele().Assemblage.SwAssemblage.FixComponent();
                    else
                        _Modele.SW.Modele().Assemblage.SwAssemblage.UnfixComponent();
                }
             }
        }

        /// <summary>
        /// Retourne si le composant est issue d'une fonction répétition.
        /// </summary>
        public Boolean EstUneRepetition
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                return _SwComposant.IsPatternInstance();
            }
        }

        /// <summary>
        /// Retourne si le composant est chargé.
        /// </summary>
        public Boolean EstCharge
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                return _SwComposant.IsLoaded();
            }
        }

        /// <summary>
        /// Retourne si le composant est issue d'une fonction symetrie.
        /// </summary>
        public Boolean EstUneSymetrie
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                return _SwComposant.IsMirrored();
            }
        }

        /// <summary>
        /// Retourne le no d'occurence du composant.
        /// </summary>
        public int NoOccurence
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());

                int pNo = 0;
                String[] pTab = _SwComposant.Name2.Split('-');
                Int32.TryParse(pTab[pTab.GetUpperBound(0)], out pNo);

                return pNo;
            }
        }

        /// <summary>
        /// Renvoi un nouvel objet Recherche
        /// </summary>
        public eRecherche NouvelleRecherche
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());

                eRecherche pNouvelleRecherche = new eRecherche();

                if (pNouvelleRecherche.Init(this))
                    return pNouvelleRecherche;

                return null;
            }
        }

        /// <summary>
        /// Renvoi la matrice de transformation du composant
        /// </summary>
        public eRepere Repere
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());

                Double[] pMatrice = _SwComposant.Transform2.ArrayData;
                eRepere pRepere = new eRepere();
                pRepere.VecteurX.X = pMatrice[0];
                pRepere.VecteurX.Y = pMatrice[1];
                pRepere.VecteurX.Y = pMatrice[2];
                pRepere.VecteurY.X = pMatrice[3];
                pRepere.VecteurY.Y = pMatrice[4];
                pRepere.VecteurY.Y = pMatrice[5];
                pRepere.VecteurZ.X = pMatrice[6];
                pRepere.VecteurZ.Y = pMatrice[7];
                pRepere.VecteurZ.Y = pMatrice[8];
                pRepere.Origine.X = pMatrice[9];
                pRepere.Origine.Y = pMatrice[10];
                pRepere.Origine.Z = pMatrice[11];
                pRepere.Echelle = pMatrice[12];
                return pRepere;
            }
        }

        /// <summary>
        /// Retourne le composant parent
        /// </summary>
        public eComposant Parent
        {
            get
            {
                Component2 pSwComposant = _SwComposant.GetParent();

                if (pSwComposant != null)
                {
                    eModele pModele = _Modele.SW.Modele(pSwComposant.GetPathName());
                    // Ensuite, on créer un nouveau Composant avec la ref du SwComposant et du modele
                    eComposant pComposant = new eComposant();
                    // Et pour que les deux soit liés, on passe la ref du Composant que l'on vient de creer
                    // au modele. Comme ca, Modele.Composant pointe sur Composant et Composant.Modele pointe sur Modele,
                    // la boucle est bouclée
                    pComposant.Init(pSwComposant, pModele);
                    pModele.Composant = pComposant;

                    return pComposant;
                }

                return null;
            }
        }

        /// <summary>
        /// Fonction interne.
        /// Test l'initialisation de l'objet ExtComposant.
        /// </summary>
        public Boolean EstInitialise { get { Debug.Info(MethodBase.GetCurrentMethod()); return _EstInitialise; } }

#endregion

#region "Méthodes"

        /// <summary>
        /// Méthode interne.
        /// Initialiser l'objet ExtComposant.
        /// </summary>
        /// <param name="SwComposant"></param>
        /// <param name="Modele"></param>
        /// <returns></returns>
        internal Boolean Init(Component2 SwComposant, eModele Modele)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            // On teste si le Modele est valide
            if ((SwComposant != null) && (Modele != null) && Modele.EstInitialise)
            {
                // On valide l'initialisation avant de recupérer la configuration
                _EstInitialise = true;

                Configuration pSwConfig;

                // Si la configuration referencé est vide, on recupère la configuration active.
                if (String.IsNullOrEmpty(SwComposant.ReferencedConfiguration))
                    pSwConfig = Modele.SwModele.GetActiveConfiguration();
                else
                    pSwConfig = Modele.SwModele.GetConfigurationByName(SwComposant.ReferencedConfiguration);

                _Configuration = new eConfiguration();

                // Si la config est ok
                if (_Configuration.Init(pSwConfig, Modele))
                {
                    _SwComposant = SwComposant;
                    _Modele = Modele;
                    _Nb = 1;

                    Debug.Info(this.Modele.FichierSw.Chemin);
                }
                else
                {
                    _EstInitialise = false;
                }
            }
            else // Sinon, on envoi pour le debug
            {
                Debug.Info("!!!!! Erreur d'initialisation");
            }

            return _EstInitialise;
        }

        /// <summary>
        /// Selectionne le composant
        /// </summary>
        /// <param name="Ajouter"></param>
        public void Selectionner(Boolean Ajouter = true)
        {
            Debug.Info(MethodBase.GetCurrentMethod());
            _SwComposant.Select4(Ajouter, null, false);
        }

        /// <summary>
        /// DeSelectionne le composant
        /// </summary>
        public void DeSelectionner()
        {
            Debug.Info(MethodBase.GetCurrentMethod());
            _SwComposant.DeSelect();
        }

        /// <summary>
        /// Méthode interne.
        /// Renvoi la liste des composants enfants filtrée par les arguments.
        /// </summary>
        /// <param name="NomComposant"></param>
        /// <param name="PrendreEnCompteSupprime"></param>
        /// <returns></returns>
        internal List<eComposant> ListComposantsEnfants(String NomComposant = "", Boolean PrendreEnCompteSupprime = false)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            List<eComposant> pListe = new List<eComposant>();

            if (SwComposant.IGetChildrenCount() == 0)
                return pListe;

            foreach (Component2 pSwComposant in SwComposant.GetChildren())
            {
                /// Si le composant est supprimé mais qu'on a decidé de le prendre en compte, c'est bon
                if (((pSwComposant.IsSuppressed() == false) | PrendreEnCompteSupprime) && Regex.IsMatch(pSwComposant.Name2, NomComposant))
                {
                    // Pour intitialiser le composant correctement il faut un peu de bidouille
                    // sinon on à le droit à une belle reference circulaire
                    // Donc d'abord, on recherche le modele du SwComposant
                    eModele pModele = _Modele.SW.Modele(pSwComposant.GetPathName());
                    // Ensuite, on créer un nouveau Composant avec la ref du SwComposant et du modele
                    eComposant pComposant = new eComposant();
                    // Et pour que les deux soit liés, on passe la ref du Composant que l'on vient de creer
                    // au modele. Comme ca, Modele.Composant pointe sur Composant et Composant.Modele pointe sur Modele,
                    // la boucle est bouclée
                    pComposant.Init(pSwComposant, pModele);
                    pModele.Composant = pComposant;
                    pListe.Add(pComposant);
                }
            }

            pListe.Sort();
            return pListe;

        }

        /// <summary>
        /// Renvoi la liste des composants enfants filtrée par les arguments.
        /// </summary>
        /// <param name="NomComposant"></param>
        /// <param name="PrendreEnCompteSupprime"></param>
        /// <returns></returns>
        public ArrayList ComposantsEnfants(String NomComposant = "", Boolean PrendreEnCompteSupprime = false)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            List<eComposant> pListeComps = ListComposantsEnfants(NomComposant, PrendreEnCompteSupprime);
            ArrayList pArrayComps = new ArrayList();

            if (pListeComps.Count > 0)
                pArrayComps = new ArrayList(pListeComps);

            return pArrayComps;
        }

        /// <summary>
        /// Méthode interne.
        /// Renvoi la liste des corps de la pièces filtrée par les arguments.
        /// </summary>
        /// <param name="TypeDeCorps"></param>
        /// <param name="PrendreEnCompteCache"></param>
        /// <returns></returns>
        internal List<eCorps> ListListeDesCorps(TypeCorps_e TypeDeCorps = TypeCorps_e.cTous)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            _Modele.Composant.Configuration.Activer();

            List<eCorps> Liste = new List<eCorps>();
            Object pInfosCorps = null;
            Object[] TableauDesCorps = _SwComposant.GetBodies3((int)swBodyType_e.swAllBodies, out pInfosCorps);

            if (TableauDesCorps.Length > 0)
            {
                foreach (Object ObjetCorps in TableauDesCorps)
                {
                    Body2 pSwCorps = (Body2)ObjetCorps;
                    eCorps pCorps = new eCorps();
                    if (pCorps.Init(pSwCorps, _Modele) && TypeDeCorps.HasFlag(pCorps.TypeDeCorps))
                    {
                        Liste.Add(pCorps);
                    }
                }
            }

            return Liste;
        }

        /// <summary>
        /// Renvoi la liste des corps de la pièces filtrée par les arguments.
        /// </summary>
        /// <param name="TypeDeCorps"></param>
        /// <param name="PrendreEnCompteCache"></param>
        /// <returns></returns>
        public ArrayList ListeDesCorps(TypeCorps_e TypeDeCorps = TypeCorps_e.cTous)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            List<eCorps> pListeCorps = ListListeDesCorps(TypeDeCorps);
            ArrayList pArrayCorps = new ArrayList();

            if (pListeCorps.Count > 0)
                pArrayCorps = new ArrayList(pListeCorps);

            return pArrayCorps;
        }

#endregion

#region "Interfaces génériques"

        public int CompareTo(eComposant Comp)
        {
            String Nom1 = _SwComposant.GetPathName() + _Configuration.Nom + "_" + NoOccurence.ToString().PadLeft(10,'0');
            String Nom2 = Comp.SwComposant.GetPathName() + Comp.Configuration.Nom + "_" + Comp.NoOccurence.ToString().PadLeft(10,'0');
            return Nom1.CompareTo(Nom2);
        }

        public int Compare(eComposant Comp1, eComposant Comp2)
        {
            String Nom1 = Comp1.SwComposant.GetPathName() + Comp1.Configuration.Nom + "_" + Comp1.NoOccurence.ToString().PadLeft(10, '0');
            String Nom2 = Comp2.SwComposant.GetPathName() + Comp2.Configuration.Nom + "_" + Comp2.NoOccurence.ToString().PadLeft(10, '0');
            return Nom1.CompareTo(Nom2);
        }

        public Boolean Equals(eComposant Comp)
        {
            String Nom1 = _SwComposant.GetPathName() + _Configuration.Nom + "_" + NoOccurence.ToString().PadLeft(10, '0');
            String Nom2 = Comp.SwComposant.GetPathName() + Comp.Configuration.Nom + "_" + Comp.NoOccurence.ToString().PadLeft(10, '0');
            return Nom1.Equals(Nom2);
        }

#endregion
    }
}
