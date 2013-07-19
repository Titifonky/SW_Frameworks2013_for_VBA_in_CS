using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace Framework
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("0AE3E176-4DF1-40F2-92CD-9A7C83C8373A")]
    public interface IeFonction
    {
        Feature SwFonction { get; }
        Feature SwFonctionOriginale { get; }
        eModele Modele { get; }
        String Nom { get; set; }
        String TypeDeLaFonction { get; }
        TypeFonction_e TypeFonction { get; }
        EtatFonction_e EtatCourant { get; }
        eFonction FonctionParent { get; }
        String DateDeCreation { get; }
        String DateDeModification { get; }
        EtatFonction_e Etat(eConfiguration Config);
        void Activer(eConfiguration Config = null, Boolean ActiverLesFonctionsDependantes = false);
        void Desactiver(eConfiguration Config = null);
        void Supprimer(swDeleteSelectionOptions_e Options);
        void Selectionner(Boolean Ajouter = true);
        void DeSelectionner();
        ArrayList ListeDesCorps();
        ArrayList ListeDesFonctionsParent(String NomARechercher = "");
        ArrayList ListeDesSousFonctions(String NomARechercher = "");
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("9009C0B9-61F1-42A1-AB7C-67DDF6AFB037")]
    [ProgId("Frameworks.eFonction")]
    public class eFonction : IeFonction, IComparable<eFonction>, IComparer<eFonction>, IEquatable<eFonction>
    {
        #region "Variables locales"

        private Boolean _EstInitialise = false;

        private eModele _Modele = null;
        private Feature _SwFonctionModele = null;
        private Feature _SwFonctionOriginale = null;
        private Object _PID = null;

        #endregion

        #region "Constructeur\Destructeur"

        public eFonction() { }

        #endregion

        #region "Propriétés"

        /// <summary>
        /// Retourne l'objet Feature associé suivant le contexte.
        /// </summary>
        public Feature SwFonction
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());

                if (_PID != null)
                {
                    int pErreur = 0;
                    Feature pSwFonction = Modele.SwModele.Extension.GetObjectByPersistReference3(_PID, out pErreur);
                    Debug.Print("PID Erreur : " + pErreur);
                    if ((pErreur == (int)swPersistReferencedObjectStates_e.swPersistReferencedObject_Ok)
                        || (pErreur == (int)swPersistReferencedObjectStates_e.swPersistReferencedObject_Suppressed))
                        _SwFonctionModele = pSwFonction;
                }
                else
                {
                    Debug.Print("Pas de PID");
                    MajPID();
                }

                ModelDoc2 pDocActif = Modele.SW.SwSW.ActiveDoc;

                if (Modele.SwModele.GetPathName().Equals(pDocActif.GetPathName()) || Modele.Composant.SwComposant.IsRoot())
                {
                    Debug.Print("Fonction du modele");
                }
                else
                {
                    Debug.Print("Fonction du composant");
                    if (_SwFonctionModele != null)
                    {
                        _SwFonctionModele = Modele.Composant.SwComposant.FeatureByName(_SwFonctionModele.Name);
                    }
                }

                return _SwFonctionModele;
            }
        }

        /// <summary>
        /// Retourne le parent ExtModele.
        /// </summary>
        public Feature SwFonctionOriginale { get { Debug.Print(MethodBase.GetCurrentMethod()); return _SwFonctionOriginale; } }

        /// <summary>
        /// Retourne le parent ExtModele.
        /// </summary>
        public eModele Modele { get { Debug.Print(MethodBase.GetCurrentMethod()); return _Modele; } }

        /// <summary>
        /// Retourne ou défini le nom de la fonction.
        /// </summary>
        public String Nom
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                return SwFonction.Name;
            }
            set
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                FeatureManager SwGestFonc = Modele.SwModele.FeatureManager;
                String pNomBase = value;
                String pNom = pNomBase;
                int Indice = 1;
                while (SwGestFonc.IsNameUsed((int)swNameType_e.swFeatureName, pNom))
                {
                    pNom = pNomBase + "_" + Indice;
                    Indice++;
                }

                SwFonction.Name = pNom;
            }
        }

        /// <summary>
        /// Retourne le type de la fonction SW.
        /// </summary>
        public String TypeDeLaFonction { get { Debug.Print(MethodBase.GetCurrentMethod()); return SwFonction.GetTypeName2(); } }

        public TypeFonction_e TypeFonction
        {
            get
            {
                return (TypeFonction_e)Enum.Parse(typeof(TypeFonction_e), "c" + TypeDeLaFonction, true);
            }
        }

        /// <summary>
        /// Renvoi l'etat "Supprimer" ou "Actif" de la fonction
        /// </summary>
        public EtatFonction_e EtatCourant
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                Boolean[] pArrayResult;

                pArrayResult = SwFonction.IsSuppressed2((int)swInConfigurationOpts_e.swThisConfiguration, null);

                if ((pArrayResult != null) && (Convert.ToBoolean(pArrayResult[0]) == false))
                    return EtatFonction_e.cActivee;

                return EtatFonction_e.cDesactivee;
            }
        }

        /// <summary>
        /// Retourne le parent directe de la fonction
        /// </summary>
        public eFonction FonctionParent
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                eFonction pFonctionParent = new eFonction();
                if (pFonctionParent.Init(SwFonction.GetOwnerFeature(), Modele))
                    return pFonctionParent;

                return null;
            }
        }

        /// <summary>
        /// Date de creation de la fonction
        /// </summary>
        public String DateDeCreation
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                return SwFonction.DateCreated;
            }
        }

        /// <summary>
        /// Date de modification de la fonction.
        /// </summary>
        public String DateDeModification
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                return SwFonction.DateModified;
            }
        }

        /// <summary>
        /// Fonction interne.
        /// Test l'initialisation de l'objet ExtModele.
        /// </summary>
        internal Boolean EstInitialise { get { Debug.Print(MethodBase.GetCurrentMethod()); return _EstInitialise; } }

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
            Debug.Print(MethodBase.GetCurrentMethod());

            if ((SwFonction != null) && (Modele != null) && Modele.EstInitialise)
            {
                _Modele = Modele;
                _SwFonctionModele = SwFonction;
                _SwFonctionOriginale = SwFonction;

                // On met à jour le PID
                MajPID();

                // On met à jour la fonction pour récupérer celle du modèle si ce n'est pas le cas
                if (_PID != null)
                {
                    int pErreur = 0;
                    Feature pSwFonction = Modele.SwModele.Extension.GetObjectByPersistReference3(_PID, out pErreur);
                    Debug.Print("PID Erreur : " + pErreur);
                    if ((pErreur == (int)swPersistReferencedObjectStates_e.swPersistReferencedObject_Ok)
                        || (pErreur == (int)swPersistReferencedObjectStates_e.swPersistReferencedObject_Suppressed))
                        _SwFonctionModele = pSwFonction;
                }

                if (_SwFonctionModele != null)
                {
                    Debug.Print(_SwFonctionModele.Name);
                    _EstInitialise = true;
                }
                else
                {
                    Debug.Print("!!!!! Erreur d'initialisation");
                }
            }

            return _EstInitialise;
        }

        private void MajPID()
        {
            Debug.Print(MethodBase.GetCurrentMethod());

            if (_SwFonctionModele == null)
                return;
            // Si la fonction est de type "MaterialFolder", la méthode GetPersistReference3 plante lamentablement
            swFeatureType_e TypeFonction = new swFeatureType_e();
            if (_SwFonctionModele.GetTypeName2() != TypeFonction.swTnMaterialFolder)
                _PID = _Modele.SwModele.Extension.GetPersistReference3(_SwFonctionModele);
        }

        /// <summary>
        /// Renvoi l'etat "Supprimer" ou "Actif" de la fonction suivant la configuration donnée
        /// </summary>
        public EtatFonction_e Etat(eConfiguration Config)
        {
            Debug.Print(MethodBase.GetCurrentMethod());
            String[] pArrayConfig = { Config.Nom };
            Boolean[] pArrayResult;

            Feature pSwFonction = SwFonction;

            pArrayResult = pSwFonction.IsSuppressed2((int)swInConfigurationOpts_e.swSpecifyConfiguration, pArrayConfig);

            if ((pArrayResult != null) && (Convert.ToBoolean(pArrayResult[0]) == false))
                return EtatFonction_e.cActivee;

            return EtatFonction_e.cDesactivee;
        }

        /// <summary>
        /// Activer la fonction.
        /// </summary>
        public void Activer(eConfiguration Config = null, Boolean ActiverLesFonctionsDependantes = false)
        {
            Debug.Print(MethodBase.GetCurrentMethod());

            String pNomConfig;

            if (Config != null)
                pNomConfig = Config.Nom;
            else
                pNomConfig = _Modele.SwModele.ConfigurationManager.ActiveConfiguration.Name;

            Feature pSwFonction = SwFonction;

            if (TypeDeLaFonction == "SketchBlockInst")
                pSwFonction = FonctionParent.SwFonction;

            String[] pTabConfig = { pNomConfig };

            pSwFonction.SetSuppression2((int)swFeatureSuppressionAction_e.swUnSuppressFeature, (int)swInConfigurationOpts_e.swSpecifyConfiguration, pTabConfig);
            
            if (ActiverLesFonctionsDependantes)
                pSwFonction.SetSuppression2((int)swFeatureSuppressionAction_e.swUnSuppressDependent, (int)swInConfigurationOpts_e.swSpecifyConfiguration, pTabConfig);

        }

        /// <summary>
        /// Desactiver la fonction.
        /// "Supprimer" la fonction en language Solidworks.
        /// </summary>
        public void Desactiver(eConfiguration Config = null)
        {
            Debug.Print(MethodBase.GetCurrentMethod());

            String pNomConfig;

            if (Config != null)
                pNomConfig = Config.Nom;
            else
                pNomConfig = _Modele.SwModele.ConfigurationManager.ActiveConfiguration.Name;

            Feature pSwFonction = SwFonction;

            if (TypeDeLaFonction == "SketchBlockInst")
                pSwFonction = FonctionParent.SwFonction;

            String[] pTabConfig = { pNomConfig };

            pSwFonction.SetSuppression2((int)swFeatureSuppressionAction_e.swSuppressFeature, (int)swInConfigurationOpts_e.swSpecifyConfiguration, pTabConfig);
        }

        /// <summary>
        /// Supprimer la fonction
        /// </summary>
        /// <param name="Options"></param>
        public void Supprimer(swDeleteSelectionOptions_e Options)
        {
            Debug.Print(MethodBase.GetCurrentMethod());

            Selectionner(false);
            Modele.SW.Modele().SwModele.Extension.DeleteSelection2((int)Options);
        }

        /// <summary>
        /// Selectionner la fonction.
        /// </summary>
        /// <param name="Ajouter"></param>
        public void Selectionner(Boolean Ajouter = true)
        {
            Debug.Print(MethodBase.GetCurrentMethod());
            Feature pSwFonction = SwFonction;
            if (pSwFonction != null)
            {
                pSwFonction.Select2(Ajouter, -1);

                //String T;
                //String NomSel = pSwFonction.GetNameForSelection(out T);
                //Debug.Print("===========================> Selectionner : " + NomSel);
                //Modele.SwModele.Extension.SelectByID2(NomSel, T, 0, 0, 0, Ajouter, -1, null, 0);
            }
            else
                Debug.Print("===========================> Selectionner : Erreur");
        }

        /// <summary>
        /// Deselectionner la fonction.
        /// </summary>
        public void DeSelectionner()
        {
            Debug.Print(MethodBase.GetCurrentMethod());
            SwFonction.DeSelect();
        }

        /// <summary>
        /// Méthode interne.
        /// Renvoi la liste des corps associés à la fonction.
        /// </summary>
        /// <returns></returns>
        internal List<eCorps> ListListeDesCorps()
        {
            Debug.Print(MethodBase.GetCurrentMethod());

            List<eCorps> pListeCorps = new List<eCorps>();

            if (SwFonction.GetFaceCount() == 0)
                return pListeCorps;

            foreach (Face2 Face in _SwFonctionModele.GetFaces())
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
            Debug.Print(MethodBase.GetCurrentMethod());

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
        internal List<eFonction> ListListeDesFonctionsParent(string NomARechercher = "")
        {
            Debug.Print(MethodBase.GetCurrentMethod());

            List<eFonction> pListeFonctions = new List<eFonction>();

            Array pFonctionsParent = SwFonction.GetParents();

            foreach (Feature pSwFonctionParent in pFonctionsParent)
            {
                eFonction pFonction = new eFonction();
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
            Debug.Print(MethodBase.GetCurrentMethod());

            List<eFonction> pListeFonctions = ListListeDesFonctionsParent(NomARechercher);
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
        internal List<eFonction> ListListeDesSousFonctions(string NomARechercher = "")
        {
            Debug.Print(MethodBase.GetCurrentMethod());

            List<eFonction> pListeFonctions = new List<eFonction>();

            Feature pSwSousFonction = SwFonction.GetFirstSubFeature();

            while (pSwSousFonction != null)
            {
                eFonction pFonction = new eFonction();

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
            Debug.Print(MethodBase.GetCurrentMethod());

            List<eFonction> pListeFonctions = ListListeDesSousFonctions(NomARechercher);
            ArrayList pArrayFonctions = new ArrayList();

            if (pListeFonctions.Count > 0)
                pArrayFonctions = new ArrayList(pListeFonctions);

            return pArrayFonctions;
        }

        #endregion

        #region "Interfaces génériques"

        public int CompareTo(eFonction Fonction)
        {
            return (_Modele.SwModele.GetPathName() + _SwFonctionModele.Name).CompareTo(Fonction.Modele.SwModele.GetPathName() + Fonction.SwFonction.Name);
        }

        public int Compare(eFonction Fonction1, eFonction Fonction2)
        {
            return (Fonction1.Modele.SwModele.GetPathName() + Fonction1._SwFonctionModele.Name).CompareTo(Fonction2.Modele.SwModele.GetPathName() + Fonction2._SwFonctionModele.Name);
        }

        public Boolean Equals(eFonction Fonction)
        {
            return (Fonction.Modele.SwModele.GetPathName() + Fonction.SwFonction.Name).Equals(_Modele.SwModele.GetPathName() + _SwFonctionModele.Name);
        }

        #endregion
    }
}
