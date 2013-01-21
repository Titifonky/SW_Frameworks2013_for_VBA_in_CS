using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using System.Text.RegularExpressions;

namespace Framework_SW2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("795533EC-5820-11E2-875F-83046188709B")]
    public interface IExtDessin
    {
        DrawingDoc SwDessin { get; }
        ExtModele Modele { get; }
        ExtFeuille FeuilleActive { get; }
        ExtFeuille Feuille(String Nom);
        Boolean FeuilleExiste(String Nom);
        ArrayList ListeDesFeuilles(String NomARechercher = "");
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("7E7EFEE8-5820-11E2-B8E9-84046188709B")]
    [ProgId("Frameworks.ExtDessin")]
    public class ExtDessin : IExtDessin
    {
        #region "Variables locales"
        private Debug _Debug = Debug.Instance;
        private Boolean _EstInitialise = false;

        private ExtModele _Modele;
        private DrawingDoc _SwDessin;
        #endregion

        #region "Constructeur\Destructeur"

        public ExtDessin() { }

        #endregion

        #region "Propriétés"

        public DrawingDoc SwDessin { get { return _SwDessin; } }

        public ExtModele Modele { get { return _Modele; } }

        public ExtFeuille FeuilleActive
        {
            get
            {
                ExtFeuille Feuille = new ExtFeuille();
                if (Feuille.Init(_SwDessin.GetCurrentSheet(), this))
                    return Feuille;
                
                return null;
            }
        }

        internal Boolean EstInitialise { get { return _EstInitialise; } }

        #endregion

        #region "Méthodes"

        internal Boolean Init(ExtModele Modele)
        {
            _MethodBase Methode = System.Reflection.MethodBase.GetCurrentMethod();

            if ((Modele != null) && Modele.EstInitialise && (Modele.TypeDuModele == TypeFichier_e.cDessin))
            {
                _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name + " : " + Modele.Chemin);

                _Modele = Modele;
                _SwDessin = Modele.SwModele as DrawingDoc;
                _EstInitialise = true;
            }
            else
            {
                _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name + " : Erreur d'initialisation");
            }

            return _EstInitialise;
        }

        public ExtFeuille Feuille(String Nom)
        {
            ExtFeuille Feuille = new ExtFeuille();
            if (Feuille.Init(_SwDessin.get_Sheet(Nom), this))
                return Feuille;

            return null;
        }

        public Boolean FeuilleExiste(String Nom)
        {
            if (_SwDessin.GetSheetCount() == 0)
                return false;

            foreach (String pNomFeuille in _SwDessin.GetSheetNames())
            {
                if (pNomFeuille == Nom)
                    return true;
            }

            return false;
        }

        internal List<ExtFeuille> ListListeDesFeuilles(String NomARechercher = "")
        {
            List<ExtFeuille> pListeFeuilles = new List<ExtFeuille>();

            if (_SwDessin.GetSheetCount() == 0)
                return pListeFeuilles;

            foreach (String NomFeuille in _SwDessin.GetSheetNames())
            {
                ExtFeuille pFeuille = new ExtFeuille();

                if (Regex.IsMatch(NomFeuille, NomARechercher) && pFeuille.Init(_SwDessin.get_Sheet(NomFeuille),this))
                {
                    pListeFeuilles.Add(pFeuille);
                }
            }

            return pListeFeuilles;

        }

        public ArrayList ListeDesFeuilles(String NomARechercher = "")
        {
            _MethodBase Methode = System.Reflection.MethodBase.GetCurrentMethod();
            _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name);

            List<ExtFeuille> pListeFeuilles = ListListeDesFeuilles(NomARechercher);
            ArrayList pArrayFeuilles = new ArrayList();

            if (pListeFeuilles.Count > 0)
                pArrayFeuilles = new ArrayList(pListeFeuilles);

            return pArrayFeuilles;
        }

        #endregion

    }
}
