using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Reflection;

namespace Framework_SW2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("0AE3E176-4DF1-40F2-92CD-9A7C83C8373A")]
    public interface IExtFonction
    {
        Feature SwFonction { get; }
        ExtModele Modele { get; }
        String Nom { get; set; }
        String TypeDeLaFonction { get; }
        EtatFonction_e Etat { get; }
        void Activer();
        void Desactiver();
        void EnregistrerEtat();
        void RestaurerEtat();
        ArrayList ListeDesCorps();
        ArrayList ListeDesSousFonctions(String NomARechercher = "");
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("9009C0B9-61F1-42A1-AB7C-67DDF6AFB037")]
    [ProgId("Frameworks.ExtFonction")]
    public class ExtFonction : IExtFonction
    {
        #region "Variables locales"
        
        private Boolean _EstInitialise = false;

        private EtatFonction_e _EtatEnregistre;
        private ExtModele _Modele;
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
        public ExtModele Modele { get { Debug.Info(MethodBase.GetCurrentMethod());  return _Modele; } }

        /// <summary>
        /// Retourne ou défini le nom de la fonction.
        /// </summary>
        public String Nom { get { Debug.Info(MethodBase.GetCurrentMethod());  return SwFonction.Name; } set { Debug.Info(MethodBase.GetCurrentMethod());  SwFonction.Name = value; } }

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
        internal Boolean Init(Feature SwFonction, ExtModele Modele)
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
        /// Méthode interne.
        /// Renvoi la liste des corps associés à la fonction.
        /// </summary>
        /// <returns></returns>
        internal List<ExtCorps> ListListeDesCorps()
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            List<ExtCorps> pListeCorps = new List<ExtCorps>();

            if (_SwFonction.GetFaceCount() == 0)
                return pListeCorps;

            foreach (Face2 Face in _SwFonction.GetFaces())
            {
                ExtCorps Corps = new ExtCorps();
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

            List<ExtCorps> pListeCorps = ListListeDesCorps();
            ArrayList pArrayCorps = new ArrayList();

            if (pListeCorps.Count > 0)
                pArrayCorps = new ArrayList(pListeCorps);

            return pArrayCorps;
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

                if ((Regex.IsMatch(pSwSousFonction.Name, NomARechercher)) && pFonction.Init(pSwSousFonction, _Modele))
                    pListeFonctions.Add(pFonction);

                pSwSousFonction = pSwSousFonction.GetNextSubFeature();
            }


            return pListeFonctions;

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
    }
}
