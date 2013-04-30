using System;
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
    public interface IExtComposant
    {
        Component2 SwComposant { get; }
        ExtModele Modele { get; }
        ExtConfiguration Configuration { get; }
        Boolean EstExclu { get; set; }
        Boolean EstSupprime { get; set; }
        int Nb { get; }
        ExtRecherche NouvelleRecherche { get; }
        Repere Repere { get; }
        ArrayList ComposantsEnfants(String NomComposant = "", Boolean PrendreEnCompteSupprime = false);
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("C46318AE-5820-11E2-A863-A3046188709B")]
    [ProgId("Frameworks.ExtComposant")]
    public class ExtComposant : IExtComposant, IComparable<ExtComposant>, IComparer<ExtComposant>, IEquatable<ExtComposant>
    {
        #region "Variables locales"

        private Boolean _EstInitialise = false;

        private Component2 _SwComposant;
        private ExtModele _Modele;
        private ExtConfiguration _Configuration;
        private int _Nb = 0;

        #endregion

        #region "Constructeur\Destructeur"

        public ExtComposant() { }

        #endregion

        #region "Propriétés"

        /// <summary>
        /// Retourne l'objet Component2 associé.
        /// </summary>
        public Component2 SwComposant { get { Debug.Info(MethodBase.GetCurrentMethod()); return _SwComposant; } }

        /// <summary>
        /// Retourne le modele ExtModele associé.
        /// </summary>
        public ExtModele Modele { get { Debug.Info(MethodBase.GetCurrentMethod()); return _Modele; } }

        /// <summary>
        /// Retourne la configuration ExtConfiguration associée.
        /// </summary>
        public ExtConfiguration Configuration { get { Debug.Info(MethodBase.GetCurrentMethod()); return _Configuration; } }

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
        /// Renvoi un nouvel objet Recherche
        /// </summary>
        public ExtRecherche NouvelleRecherche
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());

                ExtRecherche pNouvelleRecherche = new ExtRecherche();

                if (pNouvelleRecherche.Init(this))
                    return pNouvelleRecherche;

                return null;
            }
        }

        public Repere Repere
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());

                Double[] pMatrice = _SwComposant.Transform2.ArrayData;
                Repere pRepere = new Repere();
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
        internal Boolean Init(Component2 SwComposant, ExtModele Modele)
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

                _Configuration = new ExtConfiguration();

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
        /// Méthode interne.
        /// Renvoi la liste des composants enfants filtrée par les arguments.
        /// </summary>
        /// <param name="NomComposant"></param>
        /// <param name="PrendreEnCompteSupprime"></param>
        /// <returns></returns>
        internal List<ExtComposant> ListComposantsEnfants(String NomComposant = "", Boolean PrendreEnCompteSupprime = false)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            List<ExtComposant> pListe = new List<ExtComposant>();

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
                    ExtModele pModele = _Modele.SW.Modele(pSwComposant.GetPathName());
                    // Ensuite, on créer un nouveau Composant avec la ref du SwComposant et du modele
                    ExtComposant pComposant = new ExtComposant();
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

            List<ExtComposant> pListeComps = ListComposantsEnfants(NomComposant, PrendreEnCompteSupprime);
            ArrayList pArrayComps = new ArrayList();

            if (pListeComps.Count > 0)
                pArrayComps = new ArrayList(pListeComps);

            return pArrayComps;
        }

        #endregion

        #region "Interfaces génériques"

        int IComparable<ExtComposant>.CompareTo(ExtComposant Comp)
        {
            String Nom1 = _SwComposant.GetPathName() + _Configuration.Nom;
            String Nom2 = Comp.SwComposant.GetPathName() + Comp.Configuration.Nom;
            return Nom1.CompareTo(Nom2);
        }

        int IComparer<ExtComposant>.Compare(ExtComposant Comp1, ExtComposant Comp2)
        {
            String Nom1 = Comp1.SwComposant.GetPathName() + Comp1.Configuration.Nom;
            String Nom2 = Comp2.SwComposant.GetPathName() + Comp2.Configuration.Nom;
            return Nom1.CompareTo(Nom2);
        }

        bool IEquatable<ExtComposant>.Equals(ExtComposant Comp)
        {
            String Nom1 = _SwComposant.GetPathName() + _Configuration.Nom;
            String Nom2 = Comp.SwComposant.GetPathName() + Comp.Configuration.Nom;
            return Nom1.Equals(Nom2);
        }

        #endregion
    }
}