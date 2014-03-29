using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace Framework
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("E1ACC0AF-CE44-4BE7-B9CB-E732F9B587FF")]
    public interface IeEquation
    {
        eGestEquations GestEquations { get; }
        int Index { get; }
        String Nom { get; }
        String Expression { get; set; }
        String Valeur { get; }
        Boolean Renommer(String NvNom);
        Boolean Supprimer();
        eEquation CopierVers(eGestEquations GestEquations, Boolean Ecraser = false);
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("D8AAAA4F-D1D1-4B15-8A64-BD313C12168F")]
    [ProgId("Frameworks.eEquation")]
    public class eEquation : IeEquation, IComparable<ePropriete>, IComparer<ePropriete>, IEquatable<ePropriete>
    {
        #region "Variables locales"

        private static readonly String cNOMCLASSE = typeof(eEquation).Name;

        private Boolean _EstInitialise = false;

        private eGestEquations _GestEquations = null;
        private String _Nom = "";
        #endregion

        #region "Constructeur\Destructeur"

        public eEquation() { }

        #endregion

        #region "Propriétés"

        /// <summary>
        /// Retourne le parent GestEquations.
        /// </summary>
        public eGestEquations GestEquations { get { Log.Methode(cNOMCLASSE); return _GestEquations; } }

        /// <summary>
        /// Retourne l'index de l'equation.
        /// </summary>
        public int Index { get { Log.Methode(cNOMCLASSE); return (int)_GestEquations.IndexEquationAvecLeNom(_Nom); } }

        /// <summary>
        /// Retourne le nom de l'equation.
        /// </summary>
        public String Nom { get { Log.Methode(cNOMCLASSE); return _Nom; } }

        /// <summary>
        /// Retourne ou défini l'expression a calculer.
        /// </summary>
        public String Expression
        {
            get
            {
                Log.Methode(cNOMCLASSE);

                return _GestEquations.ExpressionEquationAvecIndex((int)_GestEquations.IndexEquationAvecLeNom(_Nom));
            }
            set
            {
                //if (_GestEquations.SwGestEquations.GetConfigurationOption((int)_GestEquations.IndexEquationAvecLeNom(_Nom)) == (int)swInConfigurationOpts_e.swSpecifyConfiguration)

                Log.Methode(cNOMCLASSE);
            }
        }

        /// <summary>
        /// Retourne le résultat calculée de l'expression.
        /// </summary>
        public String Valeur
        {
            get
            {
                Log.Methode(cNOMCLASSE);

                return "";
            }
        }

        /// <summary>
        /// Fonction interne.
        /// Test l'initialisation de l'objet eEquation.
        /// </summary>
        internal Boolean EstInitialise { get { Log.Methode(cNOMCLASSE); return _EstInitialise; } }

        #endregion

        #region "Méthodes"

        /// <summary>
        /// Méthode interne.
        /// Initialiser l'objet eEquation.
        /// </summary>
        /// <param name="Gestionnaire"></param>
        /// <param name="Index"></param>
        /// <returns></returns>
        internal Boolean Init(eGestEquations Gestionnaire, int Index)
        {
            Log.Methode(cNOMCLASSE);

            if ((Gestionnaire != null) && Gestionnaire.EstInitialise)
            {
                _GestEquations = Gestionnaire;
                _Nom = Gestionnaire.NomEquationAvexIndex(Index);

                _EstInitialise = true;

            }
            else
            {
                Log.Message("!!!!! Erreur d'initialisation");
            }

            return _EstInitialise;
        }

        /// <summary>
        /// Renommer l'equation.
        /// </summary>
        /// <param name="NvNom"></param>
        /// <returns></returns>
        public Boolean Renommer(String NvNom)
        {
            Log.Methode(cNOMCLASSE);

            return false;
        }

        /// <summary>
        /// Supprimer l'equation.
        /// </summary>
        /// <returns></returns>
        public Boolean Supprimer()
        {
            Log.Methode(cNOMCLASSE);

            if (_GestEquations.SupprimerEquation(_Nom))
            {
                _Nom = "";
                _EstInitialise = false;
                return true;
            }
            else
                return false;
        }

        /// <summary>
        /// Copier une equation d'un modèle à un autre
        /// </summary>
        /// <param name="GestEquations"></param>
        /// <param name="Ecraser"></param>
        public eEquation CopierVers(eGestEquations GestEquations, Boolean Ecraser = false)
        {
            if ((GestEquations != null) && (GestEquations.EstInitialise))
            {
                return null;
            }

            return null;
        }

        #endregion

        #region "Interfaces génériques"

        public int CompareTo(ePropriete Prop)
        {
            String Nom1 = _GestEquations.Modele.SwModele.GetPathName() + _Nom;
            String Nom2 = Prop.GestDeProprietes.Modele.SwModele.GetPathName() + Prop.Nom;
            return Nom1.CompareTo(Nom2);
        }

        public int Compare(ePropriete Prop1, ePropriete Prop2)
        {
            String Nom1 = Prop1.GestDeProprietes.Modele.SwModele.GetPathName() + Prop1.Nom;
            String Nom2 = Prop2.GestDeProprietes.Modele.SwModele.GetPathName() + Prop2.Nom;
            return Nom1.CompareTo(Nom2);
        }

        public Boolean Equals(ePropriete Prop)
        {
            String Nom1 = _GestEquations.Modele.SwModele.GetPathName() + _Nom;
            String Nom2 = Prop.GestDeProprietes.Modele.SwModele.GetPathName() + Prop.Nom;
            return Nom1.Equals(Nom2);
        }

        #endregion
    }
}

