using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Globalization;

namespace Framework_SW2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("0AE3E176-4DF1-40F2-92CD-9A7C83C8373A")]
    public interface IExtFonction
    {
        Feature SwFonction { get; }
        eModele Modele { get; }
        String Nom { get; set; }
        String TypeDeLaFonction { get; }
        EtatFonction_e Etat { get; }
        ExtFonction FonctionParent { get; }
        String DateDeCreation { get; }
        String DateDeModification { get; }
        void Activer();
        void Desactiver();
        void EnregistrerEtat();
        void RestaurerEtat();
        void Supprimer(swDeleteSelectionOptions_e Options);
        ArrayList ListeDesCorps();
        ArrayList ListeDesFonctionsParent(String NomARechercher = "");
        ArrayList ListeDesSousFonctions(String NomARechercher = "");
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("9009C0B9-61F1-42A1-AB7C-67DDF6AFB037")]
    [ProgId("Frameworks.ExtFonction")]
    public class ExtFonction : IExtFonction, IComparable<ExtFonction>, IComparer<ExtFonction>, IEquatable<ExtFonction>
    {
        #region "Variables locales"
        
        private Boolean _EstInitialise = false;

        private EtatFonction_e _EtatEnregistre;
        private eModele _Modele;
        private Feature _SwFonction;

        #endregion

        #region "Constructeur\Destructeur"

        public ExtFonction() { }

        #endregion

        #region "Propriétés"

        /// <summary>
        /// Retourne l'objet Feature associé.
        /// </summary>
        public Feature SwFonction { get { Debug.Info(MethodBase.GetCurrentMethod());  return _SwFonction; } }

        /// <summary>
        /// Retourne le parent ExtModele.
        /// </summary>
        public eModele Modele { get { Debug.Info(MethodBase.GetCurrentMethod());  return _Modele; } }

        /// <summary>
        /// Retourne ou défini le nom de la fonction.
        /// </summary>
        public String Nom
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                return SwFonction.Name;
            }
            set
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                FeatureManager SwGestFonc = Modele.SwModele.FeatureManager;
                String pNom = value;
                int Indice = 1;
                while (SwGestFonc.IsNameUsed((int)swNameType_e.swFeatureName, pNom))
                {
                    pNom += "_" + Indice;
                    Indice++;
                }

                SwFonction.Name = pNom;
            }
        }

        /// <summary>
        /// Retourne le type de la fonction.
        /// </summary>
        public String TypeDeLaFonction { get { Debug.Info(MethodBase.GetCurrentMethod());  return SwFonction.GetTypeName2(); } }

        /// <summary>
        /// Renvoi l'etat "Supprimer" ou "Actif" de la fonction
        /// A tester, je ne suis pas sûr du fonctionnement avec les objets
        /// </summary>
        public EtatFonction_e Etat
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                String NomConfig = _Modele.GestDeConfigurations.ConfigurationActive.Nom;
                Object[] pArrayConfig = { NomConfig };
                Object[] pArrayResult;

                pArrayResult = SwFonction.IsSuppressed2((int)swInConfigurationOpts_e.swThisConfiguration, pArrayConfig);

                if ( Convert.ToBoolean(pArrayResult[0]) == false)
                    return EtatFonction_e.cActivee;
                
                return EtatFonction_e.cDesactivee;
            }
        }

        /// <summary>
        /// Retourne le parent directe de la fonction
        /// </summary>
        public ExtFonction FonctionParent
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                ExtFonction pFonctionParent = new ExtFonction();
                if (pFonctionParent.Init(SwFonction.GetOwnerFeature(), Modele))
                    return pFonctionParent;

                return null;
            }
        }

        public String DateDeCreation
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                return _SwFonction.DateCreated;
            }
        }
        public String DateDeModification
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                return _SwFonction.DateModified;
            }
        }

        /// <summary>
        /// Fonction interne.
        /// Test l'initialisation de l'objet ExtModele.
        /// </summary>
        internal Boolean EstInitialise { get { Debug.Info(MethodBase.GetCurrentMethod());  return _EstInitialise; } }

        #endregion

        #region "Méthodes"

        /// <summary>
        /// Méthode interne.
        /// Initialiser l'objet ExtFonction.
        /// </summary>
        /// <param name="SwFonction"></param>
        /// <param name="Modele"></param>
        /// <returns></returns>
        internal Boolean Init(Feature SwFonction, eModele Modele)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            if ((SwFonction != null) && (Modele != null) && Modele.EstInitialise)
            {
                _Modele = Modele;
                _SwFonction = SwFonction;
                Debug.Info(this.Nom);
                _EstInitialise = true;
            }
            else
            {
                Debug.Info("!!!!! Erreur d'initialisation");
            }
            return _EstInitialise;
        }

        /// <summary>
        /// Activer la fonction.
        /// </summary>
        public void Activer()
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            String TypeFonction;
            String NomFonctionPourSelection = SwFonction.GetNameForSelection(out TypeFonction);
            ModelDoc2 pSwModele = _Modele.SwModele;

            pSwModele.Extension.SelectByID2(NomFonctionPourSelection, TypeFonction, 0, 0, 0, false, -1, null, 0);
            pSwModele.EditUnsuppress2();
            pSwModele.EditUnsuppressDependent2();
        }

        /// <summary>
        /// Desactiver la fonction.
        /// "Supprimer" la fonction en language Solidworks.
        /// </summary>
        public void Desactiver()
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            String TypeFonction;
            String NomFonctionPourSelection = SwFonction.GetNameForSelection(out TypeFonction);
            ModelDoc2 pSwModele = _Modele.SwModele;

            pSwModele.Extension.SelectByID2(NomFonctionPourSelection, TypeFonction, 0, 0, 0, false, -1, null, 0);
            pSwModele.EditSuppress2();
        }

        /// <summary>
        /// Enregistre l'état de la fonction.
        /// </summary>
        public void EnregistrerEtat()
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            _EtatEnregistre = Etat;
        }

        /// <summary>
        /// Restaure l'état enregistré de la fonction.
        /// </summary>
        public void RestaurerEtat()
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            if (_EtatEnregistre == EtatFonction_e.cActivee)
                Activer();
            else
                Desactiver();
        }

        /// <summary>
        /// Supprimer la fonction
        /// </summary>
        /// <param name="Options"></param>
        public void Supprimer(swDeleteSelectionOptions_e Options)
        {
            _SwFonction.Select2(false, 0);
            _Modele.SwModele.Extension.DeleteSelection2((int)Options);
        }

        /// <summary>
        /// Méthode interne.
        /// Renvoi la liste des corps associés à la fonction.
        /// </summary>
        /// <returns></returns>
        internal List<eCorps> ListListeDesCorps()
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            List<eCorps> pListeCorps = new List<eCorps>();

            if (_SwFonction.GetFaceCount() == 0)
                return pListeCorps;

            foreach (Face2 Face in _SwFonction.GetFaces())
            {
                eCorps Corps = new eCorps();
                if (Corps.Init(Face.GetBody(), _Modele.Piece) && (pListeCorps.Contains(Corps) == false))
                    pListeCorps.Add(Corps);
            }

            return pListeCorps;
        }

        /// <summary>
        /// Renvoi la liste des corps associés à la fonction.
        /// </summary>
        /// <returns></returns>
        public ArrayList ListeDesCorps()
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            List<eCorps> pListeCorps = ListListeDesCorps();
            ArrayList pArrayCorps = new ArrayList();

            if (pListeCorps.Count > 0)
                pArrayCorps = new ArrayList(pListeCorps);

            return pArrayCorps;
        }

        /// <summary>
        /// Méthode interne.
        /// Renvoi la liste des fonctions parent
        /// </summary>
        /// <param name="NomARechercher"></param>
        /// <returns></returns>
        internal List<ExtFonction> ListListeDesFonctionsParent(string NomARechercher = "")
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            List<ExtFonction> pListeFonctions = new List<ExtFonction>();

            Array pFonctionsParent = SwFonction.GetParents();

            foreach(Feature pSwFonctionParent in pFonctionsParent)
            {
                ExtFonction pFonction = new ExtFonction();
                if ((Regex.IsMatch(pSwFonctionParent.Name, NomARechercher)) && pFonction.Init(pSwFonctionParent, _Modele))
                    pListeFonctions.Add(pFonction);
            }

            return pListeFonctions;

        }

        /// <summary>
        /// Renvoi la liste des fonction parent
        /// </summary>
        /// <param name="NomARechercher"></param>
        /// <returns></returns>
        public ArrayList ListeDesFonctionsParent(string NomARechercher = "")
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            List<ExtFonction> pListeFonctions = ListListeDesFonctionsParent(NomARechercher);
            ArrayList pArrayFonctions = new ArrayList();

            if (pListeFonctions.Count > 0)
                pArrayFonctions = new ArrayList(pListeFonctions);

            return pArrayFonctions;
        }

        /// <summary>
        /// Méthode interne.
        /// Renvoi la liste des sous-fonctions
        /// </summary>
        /// <param name="NomARechercher"></param>
        /// <returns></returns>
        internal List<ExtFonction> ListListeDesSousFonctions(string NomARechercher = "")
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            List<ExtFonction> pListeFonctions = new List<ExtFonction>();

            Feature pSwSousFonction = SwFonction.GetFirstSubFeature();

            while (pSwSousFonction != null)
            {
                ExtFonction pFonction = new ExtFonction();

                if ((Regex.IsMatch(pSwSousFonction.Name, NomARechercher))
                    && pFonction.Init(pSwSousFonction, _Modele)
                    && !(pListeFonctions.Contains(pFonction)))
                    pListeFonctions.Add(pFonction);

                pSwSousFonction = pSwSousFonction.GetNextSubFeature();
            }


            return pListeFonctions.Distinct().ToList();

        }

        /// <summary>
        /// Renvoi la liste des sous-fonctions
        /// </summary>
        /// <param name="NomARechercher"></param>
        /// <returns></returns>
        public ArrayList ListeDesSousFonctions(string NomARechercher = "")
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            List<ExtFonction> pListeFonctions = ListListeDesSousFonctions(NomARechercher);
            ArrayList pArrayFonctions = new ArrayList();

            if (pListeFonctions.Count > 0)
                pArrayFonctions = new ArrayList(pListeFonctions);

            return pArrayFonctions;
        }

        #endregion

        #region "Interfaces génériques"

        int IComparable<ExtFonction>.CompareTo(ExtFonction Fonction)
        {
            return  (_Modele.SwModele.GetPathName() + _SwFonction.Name).CompareTo(Fonction.Modele.SwModele.GetPathName() +  Fonction.SwFonction.Name);
        }

        int IComparer<ExtFonction>.Compare(ExtFonction Fonction1, ExtFonction Fonction2)
        {
            return (Fonction1.Modele.SwModele.GetPathName() + Fonction1._SwFonction.Name).CompareTo(Fonction2.Modele.SwModele.GetPathName() + Fonction2._SwFonction.Name);
        }

        bool IEquatable<ExtFonction>.Equals(ExtFonction Fonction)
        {
            return (Fonction.Modele.SwModele.GetPathName() + Fonction.SwFonction.Name).Equals(_Modele.SwModele.GetPathName() + _SwFonction.Name);
        }

        #endregion
    }
}
