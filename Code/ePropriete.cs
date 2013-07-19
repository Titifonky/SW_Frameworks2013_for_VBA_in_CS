using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using SolidWorks.Interop.swconst;

namespace Framework
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
        ePropriete CopierVers(eGestDeProprietes GestDeProprietes, Boolean Ecraser = false);
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
        public eGestDeProprietes GestDeProprietes { get { Debug.Print(MethodBase.GetCurrentMethod()); return _GestDeProprietes; } }

        /// <summary>
        /// Retourne le nom de la propriété.
        /// </summary>
        public String Nom { get { Debug.Print(MethodBase.GetCurrentMethod()); return _Nom; } }

        /// <summary>
        /// Retourne le type de la propriété.
        /// </summary>
        public swCustomInfoType_e TypeDeLaPropriete
        {
            get { Debug.Print(MethodBase.GetCurrentMethod()); return (swCustomInfoType_e)_GestDeProprietes.SwGestDeProprietes.GetType2(_Nom); }
        }

        /// <summary>
        /// Retourne ou défini l'expression a calculer.
        /// </summary>
        public String Expression
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());
                String Expression;
                String Valeur;

                Get(out Expression, out Valeur);

                return Expression;
            }
            set { Debug.Print(MethodBase.GetCurrentMethod()); _GestDeProprietes.SwGestDeProprietes.Set(_Nom, value); }
        }

        /// <summary>
        /// Retourne le résultat calculée de l'expression.
        /// </summary>
        public String Valeur
        {
            get
            {
                Debug.Print(MethodBase.GetCurrentMethod());
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
        internal Boolean EstInitialise { get { Debug.Print(MethodBase.GetCurrentMethod()); return _EstInitialise; } }

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
            Debug.Print(MethodBase.GetCurrentMethod());

            if ((Gestionnaire != null) && Gestionnaire.EstInitialise && !String.IsNullOrEmpty(Nom))
            {
                List<String> pListeNom = null;

                if (Gestionnaire.SwGestDeProprietes.Count > 0)
                    pListeNom = new List<string>(Gestionnaire.SwGestDeProprietes.GetNames());

                if ((pListeNom != null) && (pListeNom.Contains(Nom)))
                {
                    Debug.Print(this.Nom);

                    _GestDeProprietes = Gestionnaire;
                    _Nom = Nom;

                    _EstInitialise = true;
                }
                else
                {
                    Debug.Print("!!!!! Erreur d'initialisation");
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
            Debug.Print(MethodBase.GetCurrentMethod());

            if (String.IsNullOrEmpty(this.Nom))
            {
                _EstInitialise = false;
                Expression = null;
                Valeur = null;
                return;
            }

            // Pour la compatibilité
#if SW2013
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
            Debug.Print(MethodBase.GetCurrentMethod());

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
            Debug.Print(MethodBase.GetCurrentMethod());

            if (_GestDeProprietes.SwGestDeProprietes.Delete(Nom) == 1)
            {
                _Nom = "";
                _EstInitialise = false;
                return true;
            }
            else
                return false;
        }

        /// <summary>
        /// Copier une propriété d'un modèle à un autre
        /// </summary>
        /// <param name="GestDeProprietes"></param>
        /// <param name="Ecraser"></param>
        public ePropriete CopierVers(eGestDeProprietes GestDeProprietes, Boolean Ecraser = false)
        {
            if ((GestDeProprietes != null) && (GestDeProprietes.EstInitialise))
            {
                return GestDeProprietes.AjouterPropriete(Nom, TypeDeLaPropriete, Expression, Ecraser);
            }

            return null;
        }

#endregion

#region "Interfaces génériques"

        public int CompareTo(ePropriete Prop)
        {
            String Nom1 = _GestDeProprietes.Modele.SwModele.GetPathName() + _Nom;
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
            String Nom1 = _GestDeProprietes.Modele.SwModele.GetPathName() + _Nom;
            String Nom2 = Prop.GestDeProprietes.Modele.SwModele.GetPathName() + Prop.Nom;
            return Nom1.Equals(Nom2);
        }

#endregion
    }
}

