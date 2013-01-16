using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.IO;
using System.Text.RegularExpressions;

/////////////////////////// Implementation terminée ///////////////////////////

namespace Framework_SW2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("0AE3E176-4DF1-40F2-92CD-9A7C83C8373A")]
    public interface IExtFonction
    {
        Feature SwFonction { get; }
        ExtPiece Piece { get; }
        String Nom { get; set; }
        String TypeDeLaFonction { get; }
        EtatFonction_e Etat { get; }
        void Activer();
        void Supprimer();
        void EnregistrerEtat();
        void RestaurerEtat();
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
        private ExtPiece _Piece;
        private Feature _SwFonction;

        #endregion

        #region "Constructeur\Destructeur"

        public ExtFonction() { }

        #endregion

        #region "Propriétés"

        public Feature SwFonction { get { return _SwFonction; } }

        public ExtPiece Piece { get { return _Piece; } }

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
                String NomConfig = _Piece.Modele.GestDeConfigurations.ConfigurationActive.Nom;
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

        internal Boolean Init(Feature SwFonction, ExtPiece Piece)
        {
            _MethodBase Methode = System.Reflection.MethodBase.GetCurrentMethod();

            if ((SwFonction != null) && (Piece != null) && Piece.EstInitialise)
            {

                _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name);

                _Piece = Piece;
                _SwFonction = SwFonction;
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
            ModelDoc2 pSwModele = _Piece.Modele.SwModele;

            pSwModele.Extension.SelectByID2(NomFonctionPourSelection, TypeFonction, 0, 0, 0, false, -1, null, 0);
            pSwModele.EditUnsuppress2();
            pSwModele.EditUnsuppressDependent2();
        }

        public void Supprimer()
        {
            String TypeFonction;
            String NomFonctionPourSelection = SwFonction.GetNameForSelection(out TypeFonction);
            ModelDoc2 pSwModele = _Piece.Modele.SwModele;

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
                Supprimer();
        }

        internal List<ExtFonction> ListListeDesSousFonctions(string NomARechercher = "")
        {
            List<ExtFonction> pListeFonctions = new List<ExtFonction>();

            Feature pSwSousFonction = SwFonction.GetFirstSubFeature();

            while (pSwSousFonction != null)
            {
                ExtFonction pFonction = new ExtFonction();

                if ((Regex.IsMatch(pSwSousFonction.Name, NomARechercher)) && pFonction.Init(pSwSousFonction,_Piece))
                    pListeFonctions.Add(pFonction);

                pSwSousFonction = pSwSousFonction.GetNextSubFeature();
            }


            return pListeFonctions;

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
