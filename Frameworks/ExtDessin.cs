using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using System.Text.RegularExpressions;
using System.Reflection;

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
        private Boolean _EstInitialise = false;

        private ExtModele _Modele;
        private DrawingDoc _SwDessin;
        #endregion

        #region "Constructeur\Destructeur"

        public ExtDessin() { }

        #endregion

        #region "Propriétés"

        public DrawingDoc SwDessin { get { Debug.Info(MethodBase.GetCurrentMethod());  return _SwDessin; } }

        public ExtModele Modele { get { Debug.Info(MethodBase.GetCurrentMethod());  return _Modele; } }

        public ExtFeuille FeuilleActive
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());

                ExtFeuille pFeuille = new ExtFeuille();
                Sheet pSwFeuille = _SwDessin.GetCurrentSheet();
                if (pFeuille.Init(pSwFeuille, this))
                    return pFeuille;

                return null;
            }
        }

        internal Boolean EstInitialise { get { Debug.Info(MethodBase.GetCurrentMethod());  return _EstInitialise; } }

        #endregion

        #region "Méthodes"

        internal Boolean Init(ExtModele Modele)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            if ((Modele != null) && Modele.EstInitialise && (Modele.TypeDuModele == TypeFichier_e.cDessin))
            {
                Debug.Info(Modele.Chemin);

                _Modele = Modele;
                _SwDessin = Modele.SwModele as DrawingDoc;
                _EstInitialise = true;
            }
            else
            {
                Debug.Info("\t !!!!! Erreur d'initialisation");
            }

            return _EstInitialise;
        }

        public ExtFeuille Feuille(String Nom)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            ExtFeuille pFeuille = new ExtFeuille();
            Sheet pSwFeuille = _SwDessin.get_Sheet(Nom);

            if (pFeuille.Init(pSwFeuille, this))
                return pFeuille;

            return null;
        }

        public Boolean FeuilleExiste(String Nom)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

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
            Debug.Info(MethodBase.GetCurrentMethod());

            List<ExtFeuille> pListeFeuilles = new List<ExtFeuille>();

            if (_SwDessin.GetSheetCount() == 0)
                return pListeFeuilles;

            foreach (String NomFeuille in _SwDessin.GetSheetNames())
            {
                ExtFeuille pFeuille = new ExtFeuille();
                Sheet pSwFeuille = _SwDessin.get_Sheet(NomFeuille);

                if (Regex.IsMatch(NomFeuille, NomARechercher) && pFeuille.Init(pSwFeuille, this))
                {
                    pListeFeuilles.Add(pFeuille);
                }
            }

            return pListeFeuilles;

        }

        public ArrayList ListeDesFeuilles(String NomARechercher = "")
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            List<ExtFeuille> pListeFeuilles = ListListeDesFeuilles(NomARechercher);
            ArrayList pArrayFeuilles = new ArrayList();

            if (pListeFeuilles.Count > 0)
                pArrayFeuilles = new ArrayList(pListeFeuilles);

            return pArrayFeuilles;
        }

        #endregion

    }
}
