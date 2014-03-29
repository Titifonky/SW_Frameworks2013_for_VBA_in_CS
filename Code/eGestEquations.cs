using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections;
using System.Runtime.InteropServices;

namespace Framework
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("C5BA0E43-E538-4BFD-87A2-9DD7032FDEC3")]
    public interface IeGestEquations
    {
        EquationMgr SwGestEquations { get; }
        eModele Modele { get; }
        eEquation AjouterEquation(String Nom, String Expression, swInConfigurationOpts_e QuelConfigs, ArrayList ListeConfigs = null, Boolean EcraserExistante = false, int Index = -1);
        eEquation RecupererEquation(String Nom);
        eEquation RecupererEquation(int Index);
        Boolean SupprimerEquation(String Nom);
        Boolean SupprimerEquation(int Index);
        Boolean EquationExiste(String Nom);
        ArrayList ListeDesEquations(String NomARechercher = "");
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("25D7C4FA-40FB-46E0-9A5B-75022C9D8E53")]
    [ProgId("Frameworks.eGestEquations")]
    public class eGestEquations : IeGestEquations
    {
        #region "Variables locales"

        private static readonly String cNOMCLASSE = typeof(eGestEquations).Name;

        private Boolean _EstInitialise = false;

        private eModele _Modele = null;
        private EquationMgr _SwGestEquations = null;

        #endregion

        #region "Constructeur\Destructeur"

        public eGestEquations() { }

        #endregion

        #region "Propriétés"

        /// <summary>
        /// Retourne le gestionnaire EquationMgr associé
        /// </summary>
        public EquationMgr SwGestEquations { get { Log.Methode(cNOMCLASSE); return _SwGestEquations; } }

        /// <summary>
        /// Retourne le parent eModele 
        /// </summary>
        public eModele Modele { get { Log.Methode(cNOMCLASSE); return _Modele; } }

        /// <summary>
        /// Fonction interne
        /// Test l'initialisation de l'objet eModele
        /// </summary>
        internal Boolean EstInitialise { get { Log.Methode(cNOMCLASSE); return _EstInitialise; } }

        #endregion

        #region "Méthodes"

        /// <summary>
        /// Méthode interne
        /// Initialise l'objet GestEquations
        /// </summary>
        /// <param name="SwGestionnaire"></param>
        /// <param name="Modele"></param>
        /// <returns></returns>
        internal Boolean Init(EquationMgr SwGestionnaire, eModele Modele)
        {
            Log.Methode(cNOMCLASSE);

            if ((SwGestionnaire != null) && (Modele != null) && Modele.EstInitialise)
            {
                _SwGestEquations = SwGestionnaire;
                _Modele = Modele;
                _EstInitialise = true;
            }
            else
            {
                Log.Message("!!!!! Erreur d'initialisation");
            }

            return _EstInitialise;
        }

        /// <summary>
        /// Ajouter une equation
        /// </summary>
        /// <param name="Nom"></param>
        /// <param name="TypePropriete"></param>
        /// <param name="Expression"></param>
        /// <param name="EcraserExistante"></param>
        /// <returns></returns>
        public eEquation AjouterEquation(String Nom, String Expression, swInConfigurationOpts_e QuelConfigs, ArrayList ListeConfigs = null, Boolean EcraserExistante = false, int Index = -1)
        {
            Log.Methode(cNOMCLASSE);

            // Si on écrase, on supprime l'equation
            if (EcraserExistante)
                SupprimerEquation(Nom);

            // On la récupère
            eEquation pEqu = RecupererEquation(Nom);

            Object[] pNomsConfigs = null;

            // Si elle n'existe pas on la crée et on lui assigne l'expression
            if (pEqu == null)
            {
                if (ListeConfigs.Count > 0)
                {
                    pNomsConfigs = new Object[ListeConfigs.Count];
                    for (int i = 0; i < ListeConfigs.Count; i++)
                    {
                        Log.Message("===================================>>>>" + ((eConfiguration)ListeConfigs[i]).Nom);
                        pNomsConfigs[i] = ((eConfiguration)ListeConfigs[i]).Nom;
                    }
                }

                if (_SwGestEquations.Add3(Index, '"' + Nom + "\" = \"" + Expression + '"', true, (int)QuelConfigs, pNomsConfigs) != -1)
                {
                    pEqu = new eEquation();
                    pEqu.Init(this, Index);
                }
            }

            // Si tout est ok, on la renvoi
            if (pEqu.EstInitialise)
                return pEqu;

            return null;
        }

        /// <summary>
        /// Récupérer une equation
        /// </summary>
        /// <param name="Nom"></param>
        /// <returns></returns>
        public eEquation RecupererEquation(String Nom)
        {
            Log.Methode(cNOMCLASSE);

            int? Index = IndexEquationAvecLeNom(Nom);

            if (Index != null)
            {
                eEquation pEqu = new eEquation();

                if (pEqu.Init(this, (int)Index))
                    return pEqu;
            }

            return null;
        }

        public eEquation RecupererEquation(int Index)
        {
            Log.Methode(cNOMCLASSE);

            eEquation pEqu = new eEquation();

            if (pEqu.Init(this, Index))
                return pEqu;

            return null;
        }

        /// <summary>
        /// Supprime une equation
        /// </summary>
        /// <param name="Nom"></param>
        /// <returns></returns>
        public Boolean SupprimerEquation(String Nom)
        {
            Log.Methode(cNOMCLASSE);

            int? Index = IndexEquationAvecLeNom(Nom);

            if (Index != null)
            {
                if (_SwGestEquations.Delete((int)Index) == 1)
                    return true;
            }

            return false;
        }

        /// <summary>
        /// Supprime une equation
        /// </summary>
        /// <param name="Nom"></param>
        /// <returns></returns>
        public Boolean SupprimerEquation(int Index)
        {
            Log.Methode(cNOMCLASSE);

            if (_SwGestEquations.Delete(Index) == 0)
                return true;

            return false;
        }

        /// <summary>
        /// Teste l'existence d'une equation
        /// </summary>
        /// <param name="Nom"></param>
        /// <returns></returns>
        public Boolean EquationExiste(String Nom)
        {
            if (IndexEquationAvecLeNom(Nom) != null)
                return true;

            return false;
        }

        internal int? IndexEquationAvecLeNom(String Nom)
        {
            if (_SwGestEquations.GetCount() > 0)
            {
                for (int i = 0; i < _SwGestEquations.GetCount(); i++)
                {
                    if (NomEquationAvexIndex(i) == Nom)
                        return i;
                }
            }

            return null;
        }

        internal String NomEquationAvexIndex(int Index)
        {
            if (Index < _SwGestEquations.GetCount())
            {
                String pEqu = _SwGestEquations.get_Equation(Index);
                return pEqu.Split('=')[0].Replace('"', ' ').Trim();
            }

            return null;
        }

        internal String ExpressionEquationAvecIndex(int Index)
        {
            if (Index < _SwGestEquations.GetCount())
            {
                String pEqu = _SwGestEquations.get_Equation(Index);
                return pEqu.Split('=')[1].Trim();
            }

            return null;
        }

        /// <summary>
        /// Méthode interne
        /// Renvoi la liste des equation filtrée par les arguments
        /// Si NomARechercher est vide, toutes les equations sont renvoyées
        /// </summary>
        /// <param name="NomARechercher"></param>
        /// <returns></returns>
        public ArrayList ListeDesEquations(String NomARechercher = "")
        {
            Log.Methode(cNOMCLASSE);

            ArrayList pListeEquations = new ArrayList();

            if (_SwGestEquations.GetCount() > 0)
            {
                for (int i = 0; i < _SwGestEquations.GetCount(); i++)
                {
                    eEquation pEqu = new eEquation();
                    if (pEqu.Init(this, i) && exRegex.IsMatch(NomEquationAvexIndex(i), NomARechercher))
                        pListeEquations.Add(pEqu);
                }
            }

            return pListeEquations;
        }

        #endregion
    }
}
