using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

/////////////////////////// Implementation terminée ///////////////////////////

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
        private Debug _Debug = Debug.Instance;
        private Boolean _EstInitialise = false;

        private EtatFonction_e _EtatEnregistre;
        private ExtModele _Modele;
        private Feature _SwFonction;

        #endregion

        #region "Constructeur\Destructeur"

        public ExtFonction() { }

        #endregion

        #region "Propriétés"

        public Feature SwFonction { get { return _SwFonction; } }

        public ExtModele Modele { get { return _Modele; } }

        public String Nom { get { return SwFonction.Name; } set { SwFonction.Name = value; } }

        public String TypeDeLaFonction { get { return SwFonction.GetTypeName2(); } }

        /// <summary>
        /// Renvoi l'etat "Supprimer" ou "Actif" de la fonction
        /// A tester, je ne suis pas sûr du fonctionnement avec les objets
        /// </summary>
        public EtatFonction_e Etat
        {
            get
            {
                String NomConfig = _Modele.GestDeConfigurations.ConfigurationActive.Nom;
                Object[] pArrayConfig = { NomConfig };
                Object[] pArrayResult;

                pArrayResult = SwFonction.IsSuppressed2((int)swInConfigurationOpts_e.swThisConfiguration, pArrayConfig);

                if ( Convert.ToBoolean(pArrayResult[0]) == false)
                    return EtatFonction_e.cActivee;
                
                return EtatFonction_e.cDesactivee;
            }
        }

        internal Boolean EstInitialise { get { return _EstInitialise; } }

        #endregion

        #region "Méthodes"

        internal Boolean Init(Feature SwFonction, ExtModele Modele)
        {
            _MethodBase Methode = System.Reflection.MethodBase.GetCurrentMethod();

            if ((SwFonction != null) && (Modele != null) && Modele.EstInitialise)
            {
                _Modele = Modele;
                _SwFonction = SwFonction;

                _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name + " : " + this.Nom);
                _EstInitialise = true;
            }
            else
            {
                _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name + " : Erreur d'initialisation");
            }
            return _EstInitialise;
        }

        public void Activer()
        {
            String TypeFonction;
            String NomFonctionPourSelection = SwFonction.GetNameForSelection(out TypeFonction);
            ModelDoc2 pSwModele = _Modele.SwModele;

            pSwModele.Extension.SelectByID2(NomFonctionPourSelection, TypeFonction, 0, 0, 0, false, -1, null, 0);
            pSwModele.EditUnsuppress2();
            pSwModele.EditUnsuppressDependent2();
        }

        public void Desactiver()
        {
            String TypeFonction;
            String NomFonctionPourSelection = SwFonction.GetNameForSelection(out TypeFonction);
            ModelDoc2 pSwModele = _Modele.SwModele;

            pSwModele.Extension.SelectByID2(NomFonctionPourSelection, TypeFonction, 0, 0, 0, false, -1, null, 0);
            pSwModele.EditSuppress2();
        }

        public void EnregistrerEtat()
        {
            _EtatEnregistre = Etat;
        }

        public void RestaurerEtat()
        {
            if (_EtatEnregistre == EtatFonction_e.cActivee)
                Activer();
            else
                Desactiver();
        }

        internal List<ExtCorps> ListListeDesCorps()
        {
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

        internal List<ExtFonction> ListListeDesSousFonctions(string NomARechercher = "")
        {
            List<ExtFonction> pListeFonctions = new List<ExtFonction>();

            Feature pSwSousFonction = SwFonction.GetFirstSubFeature();

            while (pSwSousFonction != null)
            {
                ExtFonction pFonction = new ExtFonction();

                if ((Regex.IsMatch(pSwSousFonction.Name, NomARechercher)) && pFonction.Init(pSwSousFonction,_Modele))
                    pListeFonctions.Add(pFonction);

                pSwSousFonction = pSwSousFonction.GetNextSubFeature();
            }


            return pListeFonctions;

        }

        public ArrayList ListeDesCorps()
        {
            _MethodBase Methode = System.Reflection.MethodBase.GetCurrentMethod();
            _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name);

            List<ExtCorps> pListeCorps = ListListeDesCorps();
            ArrayList pArrayCorps = new ArrayList();

            if (pListeCorps.Count > 0)
                pArrayCorps = new ArrayList(pListeCorps);

            return pArrayCorps;
        }

        public ArrayList ListeDesSousFonctions(string NomARechercher = "")
        {
            _MethodBase Methode = System.Reflection.MethodBase.GetCurrentMethod();
            _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name);

            List<ExtFonction> pListeFonctions = ListListeDesSousFonctions(NomARechercher);
            ArrayList pArrayFonctions = new ArrayList();

            if (pListeFonctions.Count > 0)
                pArrayFonctions = new ArrayList(pListeFonctions);

            return pArrayFonctions;
        }

        #endregion
    }
}
