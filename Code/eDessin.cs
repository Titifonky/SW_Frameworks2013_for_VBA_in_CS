using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using System.Text.RegularExpressions;
using System.Reflection;
using SolidWorks.Interop.swconst;

namespace Framework
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("795533EC-5820-11E2-875F-83046188709B")]
    public interface IeDessin
    {
        DrawingDoc SwDessin { get; }
        eModele Modele { get; }
        eFeuille FeuilleActive { get; }
        eFeuille Feuille(String Nom);
        Boolean FeuilleExiste(String Nom);
        ArrayList ListeDesFeuilles(String NomARechercher = "");
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("7E7EFEE8-5820-11E2-B8E9-84046188709B")]
    [ProgId("Frameworks.eDessin")]
    public class eDessin : IeDessin
    {
#region "Variables locales"
        private Boolean _EstInitialise = false;

        private eModele _Modele = null;
        private DrawingDoc _SwDessin = null;
#endregion

#region "Constructeur\Destructeur"

        public eDessin() { }

#endregion

#region "Propriétés"

        /// <summary>
        /// Retourne l'objet DrawingDoc associé.
        /// </summary>
        public DrawingDoc SwDessin { get { Debug.Print(MethodBase.GetCurrentMethod());  return _SwDessin; } }

        /// <summary>
        /// Retourne le parent ExtModele.
        /// </summary>
        public eModele Modele { get { Debug.Print(MethodBase.GetCurrentMethod());  return _Modele; } }

        /// <summary>
        /// Retourne la feuille active.
        /// </summary>
        public eFeuille FeuilleActive
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());

                eFeuille pFeuille = new eFeuille();
                Sheet pSwFeuille = _SwDessin.GetCurrentSheet();
                if (pFeuille.Init(pSwFeuille, this))
                    return pFeuille;

                return null;
            }
        }

        /// <summary>
        /// Fonction interne.
        /// Test l'initialisation de l'objet ExtDessin.
        /// </summary>
        internal Boolean EstInitialise { get { Debug.Print(MethodBase.GetCurrentMethod());  return _EstInitialise; } }

#endregion

#region "Méthodes"

        /// <summary>
        /// Méthode interne.
        /// Initialiser l'objet ExtDessin.
        /// </summary>
        /// <param name="Modele"></param>
        /// <returns></returns>
        internal Boolean Init(eModele Modele)
        {
            Debug.Print(MethodBase.GetCurrentMethod());

            if ((Modele != null) && Modele.EstInitialise && (Modele.TypeDuModele == TypeFichier_e.cDessin))
            {
                Debug.Print(Modele.FichierSw.Chemin);

                _Modele = Modele;
                _SwDessin = Modele.SwModele as DrawingDoc;
                _EstInitialise = true;
                _Modele.SwModele.Extension.UsePageSetup = (int)swPageSetupInUse_e.swPageSetupInUse_DrawingSheet;
            }
            else
            {
                Debug.Print("\t !!!!! Erreur d'initialisation");
            }

            return _EstInitialise;
        }

        /// <summary>
        /// Renvoi la feuille à partir du nom.
        /// </summary>
        /// <param name="Nom"></param>
        /// <returns></returns>
        public eFeuille Feuille(String Nom)
        {
            Debug.Print(MethodBase.GetCurrentMethod());

            eFeuille pFeuille = new eFeuille();
            Sheet pSwFeuille = _SwDessin.get_Sheet(Nom);

            if (pFeuille.Init(pSwFeuille, this))
                return pFeuille;

            return null;
        }

        /// <summary>
        /// Test si la feuille existe.
        /// </summary>
        /// <param name="Nom"></param>
        /// <returns></returns>
        public Boolean FeuilleExiste(String Nom)
        {
            Debug.Print(MethodBase.GetCurrentMethod());

            if (_SwDessin.GetSheetCount() == 0)
                return false;

            foreach (String pNomFeuille in _SwDessin.GetSheetNames())
            {
                if (pNomFeuille == Nom)
                    return true;
            }

            return false;
        }

        /// <summary>
        /// Méthode interne.
        /// Renvoi la liste des feuilles filtrée par les arguments.
        /// </summary>
        /// <param name="NomARechercher"></param>
        /// <returns></returns>
        internal List<eFeuille> ListListeDesFeuilles(String NomARechercher = "")
        {
            Debug.Print(MethodBase.GetCurrentMethod());

            List<eFeuille> pListeFeuilles = new List<eFeuille>();

            if (_SwDessin.GetSheetCount() == 0)
                return pListeFeuilles;

            foreach (String NomFeuille in _SwDessin.GetSheetNames())
            {
                eFeuille pFeuille = new eFeuille();
                Sheet pSwFeuille = _SwDessin.get_Sheet(NomFeuille);

                if (Regex.IsMatch(NomFeuille, NomARechercher) && pFeuille.Init(pSwFeuille, this))
                {
                    pListeFeuilles.Add(pFeuille);
                }
            }

            return pListeFeuilles;

        }

        /// <summary>
        /// Renvoi la liste des feuilles filtrée par les arguments.
        /// </summary>
        /// <param name="NomARechercher"></param>
        /// <returns></returns>
        public ArrayList ListeDesFeuilles(String NomARechercher = "")
        {
            Debug.Print(MethodBase.GetCurrentMethod());

            List<eFeuille> pListeFeuilles = ListListeDesFeuilles(NomARechercher);
            ArrayList pArrayFeuilles = new ArrayList();

            if (pListeFeuilles.Count > 0)
                pArrayFeuilles = new ArrayList(pListeFeuilles);

            return pArrayFeuilles;
        }

#endregion

    }
}
