using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

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
        Boolean AfficherNotesDePliage { get; set; }
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("7E7EFEE8-5820-11E2-B8E9-84046188709B")]
    [ProgId("Frameworks.eDessin")]
    public class eDessin : IeDessin
    {
        #region "Variables locales"

        private static readonly String cNOMCLASSE = typeof(eDessin).Name;

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
        public DrawingDoc SwDessin { get { Log.Methode(cNOMCLASSE); return _SwDessin; } }

        /// <summary>
        /// Retourne le parent eModele.
        /// </summary>
        public eModele Modele { get { Log.Methode(cNOMCLASSE); return _Modele; } }

        /// <summary>
        /// Retourne la feuille active.
        /// </summary>
        public eFeuille FeuilleActive
        {
            get
            {
                Log.Methode(cNOMCLASSE);

                eFeuille pFeuille = new eFeuille();
                Sheet pSwFeuille = _SwDessin.GetCurrentSheet();
                if (pFeuille.Init(pSwFeuille, this))
                    return pFeuille;

                return null;
            }
        }

        /// <summary>
        /// Fonction interne.
        /// Test l'initialisation de l'objet eDessin.
        /// </summary>
        internal Boolean EstInitialise { get { Log.Methode(cNOMCLASSE); return _EstInitialise; } }

        public Boolean AfficherNotesDePliage
        {
            get
            {
                Log.Methode(cNOMCLASSE);
                return _Modele.SwModele.Extension.GetUserPreferenceToggle(
                                                                            (int)swUserPreferenceToggle_e.swShowSheetMetalBendNotes,
                                                                            (int)swUserPreferenceOption_e.swDetailingNoOptionSpecified
                                                                            );
            }
            set
            {
                Log.Methode(cNOMCLASSE);
                _Modele.SwModele.Extension.SetUserPreferenceToggle(
                                                                    (int)swUserPreferenceToggle_e.swShowSheetMetalBendNotes,
                                                                    (int)swUserPreferenceOption_e.swDetailingNoOptionSpecified,
                                                                    value
                                                                    );
            }
        }

        #endregion

        #region "Méthodes"

        /// <summary>
        /// Méthode interne.
        /// Initialiser l'objet eDessin.
        /// </summary>
        /// <param name="Modele"></param>
        /// <returns></returns>
        internal Boolean Init(eModele Modele)
        {
            Log.Methode(cNOMCLASSE);

            if ((Modele != null) && Modele.EstInitialise && (Modele.TypeDuModele == TypeFichier_e.cDessin))
            {
                Log.Message(Modele.FichierSw.Chemin);

                _Modele = Modele;
                _SwDessin = Modele.SwModele as DrawingDoc;
                _EstInitialise = true;
                _Modele.SwModele.Extension.UsePageSetup = (int)swPageSetupInUse_e.swPageSetupInUse_DrawingSheet;
            }
            else
            {
                Log.Message("\t !!!!! Erreur d'initialisation");
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
            Log.Methode(cNOMCLASSE);

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
            Log.Methode(cNOMCLASSE);

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
        public ArrayList ListeDesFeuilles(String NomARechercher = "")
        {
            Log.Methode(cNOMCLASSE);

            ArrayList pListeFeuilles = new ArrayList();

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

        #endregion

    }
}
