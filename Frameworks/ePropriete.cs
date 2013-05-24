using System;
using System.Runtime.InteropServices;
using SolidWorks.Interop.swconst;
using System.Collections.Generic;
using System.Reflection;

namespace Framework_SW2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("5E7B77AC-630B-11E2-9FF2-8FF06088709B")]
    public interface IePropriete
    {
        eGestDeProprietes GestDeProprietes { get; }
        String Nom { get; }
        swCustomInfoType_e TypeDeLaPropriete { get; }
        String Expression { get; set; }
        String Valeur { get; }
        Boolean Renommer(String NvNom);
        Boolean Supprimer();
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("6502CAB2-630B-11E2-B156-90F06088709B")]
    [ProgId("Frameworks.ePropriete")]
    public class ePropriete : IePropriete, IComparable<ePropriete>, IComparer<ePropriete>, IEquatable<ePropriete>
    {
        #region "Variables locales"

        private Boolean _EstInitialise = false;

        private eGestDeProprietes _GestDeProprietes = null;
        private String _Nom = "";
        #endregion

        #region "Constructeur\Destructeur"

        public ePropriete() { }

        #endregion

        #region "Propriétés"

        /// <summary>
        /// Retourne le parent GestDeProprietes.
        /// </summary>
        public eGestDeProprietes GestDeProprietes { get { Debug.Info(MethodBase.GetCurrentMethod()); return _GestDeProprietes; } }

        /// <summary>
        /// Retourne le nom de la propriété.
        /// </summary>
        public String Nom { get { Debug.Info(MethodBase.GetCurrentMethod()); return _Nom; } }

        /// <summary>
        /// Retourne le type de la propriété.
        /// </summary>
        public swCustomInfoType_e TypeDeLaPropriete
        {
            get { Debug.Info(MethodBase.GetCurrentMethod()); return (swCustomInfoType_e)_GestDeProprietes.SwGestDeProprietes.GetType2(_Nom); }
        }

        /// <summary>
        /// Retourne ou défini l'expression a calculer.
        /// </summary>
        public String Expression
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                String Expression;
                String Valeur;

                Get(out Expression, out Valeur);

                return Expression;
            }
            set { Debug.Info(MethodBase.GetCurrentMethod()); _GestDeProprietes.SwGestDeProprietes.Set(_Nom, value); }
        }

        /// <summary>
        /// Retourne le résultat calculée de l'expression.
        /// </summary>
        public String Valeur
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                String Expression;
                String Valeur;

                Get(out Expression, out Valeur);

                return Valeur;
            }
        }

        /// <summary>
        /// Fonction interne.
        /// Test l'initialisation de l'objet ExtPropriete.
        /// </summary>
        internal Boolean EstInitialise { get { Debug.Info(MethodBase.GetCurrentMethod()); return _EstInitialise; } }

        #endregion

        #region "Méthodes"

        /// <summary>
        /// Méthode interne.
        /// Initialiser l'objet ExtPropriete.
        /// </summary>
        /// <param name="Gestionnaire"></param>
        /// <param name="Nom"></param>
        /// <returns></returns>
        internal Boolean Init(eGestDeProprietes Gestionnaire, String Nom)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            if ((Gestionnaire != null) && Gestionnaire.EstInitialise && !String.IsNullOrEmpty(Nom))
            {
                List<String> pListeNom = new List<string>(Gestionnaire.SwGestDeProprietes.GetNames());

                if (pListeNom.Contains(Nom))
                {
                    Debug.Info(this.Nom);

                    _GestDeProprietes = Gestionnaire;
                    _Nom = Nom;

                    _EstInitialise = true;
                }
                else
                {
                    Debug.Info("!!!!! Erreur d'initialisation");
                }
            }

            return _EstInitialise;
        }

        /// <summary>
        /// Méthode privée.
        /// Retourne Expression et Valeur.
        /// Permet une compatiblitée descendante en testant la version SW utilisée.
        /// </summary>
        /// <param name="Expression"></param>
        /// <param name="Valeur"></param>
        private void Get(out String Expression, out String Valeur)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            if (String.IsNullOrEmpty(this.Nom))
            {
                _EstInitialise = false;
                Expression = null;
                Valeur = null;
                return;
            }

            // Pour la compatibilité
#if SW2012
            _GestDeProprietes.SwGestDeProprietes.Get4(_Nom, true, out Expression, out Valeur);
#else
                _GestDeProprietes.SwGestDeProprietes.Get2(_Nom, out Expression, out Valeur);
#endif

        }

        /// <summary>
        /// Renommer la propriété.
        /// </summary>
        /// <param name="NvNom"></param>
        /// <returns></returns>
        public Boolean Renommer(String NvNom)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            swCustomInfoType_e pTypeDeLaPropriete = TypeDeLaPropriete;
            String pExpression = Expression;
            if (_GestDeProprietes.SwGestDeProprietes.Delete(_Nom) == 1)
            {
                _GestDeProprietes.SwGestDeProprietes.Add2(NvNom, (int)pTypeDeLaPropriete, pExpression);
                _Nom = NvNom;
                return true;
            }
            return false;
        }

        /// <summary>
        /// Supprimer la propriété.
        /// </summary>
        /// <returns></returns>
        public Boolean Supprimer()
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            if (_GestDeProprietes.SwGestDeProprietes.Delete(Nom) == 1)
            {
                _Nom = "";
                _EstInitialise = false;
                return true;
            }
            else
                return false;
        }

        #endregion

        #region "Interfaces génériques"

        int IComparable<ePropriete>.CompareTo(ePropriete Prop)
        {
            String Nom1 = _GestDeProprietes.Modele.SwModele.GetPathName() + _Nom;
            String Nom2 = Prop.GestDeProprietes.Modele.SwModele.GetPathName() + Prop.Nom;
            return Nom1.CompareTo(Nom2);
        }

        int IComparer<ePropriete>.Compare(ePropriete Prop1, ePropriete Prop2)
        {
            String Nom1 = Prop1.GestDeProprietes.Modele.SwModele.GetPathName() + Prop1.Nom;
            String Nom2 = Prop2.GestDeProprietes.Modele.SwModele.GetPathName() + Prop2.Nom;
            return Nom1.CompareTo(Nom2);
        }

        bool IEquatable<ePropriete>.Equals(ePropriete Prop)
        {
            String Nom1 = _GestDeProprietes.Modele.SwModele.GetPathName() + _Nom;
            String Nom2 = Prop.GestDeProprietes.Modele.SwModele.GetPathName() + Prop.Nom;
            return Nom1.Equals(Nom2);
        }

        #endregion
    }
}

