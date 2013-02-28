using System;
using System.Runtime.InteropServices;
using SolidWorks.Interop.swconst;
using System.Collections.Generic;
using System.Reflection;

namespace Framework_SW2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("5E7B77AC-630B-11E2-9FF2-8FF06088709B")]
    public interface IExtPropriete
    {
        GestDeProprietes GestDeProprietes { get; }
        String Nom { get; }
        swCustomInfoType_e TypeDeLaPropriete { get; }
        String Expression { get; set; }
        String Valeur { get;}
        Boolean Renommer(String NvNom);
        Boolean Supprimer();
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("6502CAB2-630B-11E2-B156-90F06088709B")]
    [ProgId("Frameworks.GestDeConfigurations")]
    public class ExtPropriete : IExtPropriete, IComparable<ExtPropriete>, IComparer<ExtPropriete>, IEquatable<ExtPropriete>
    {
        #region "Variables locales"
        
        private Boolean _EstInitialise = false;

        private GestDeProprietes _GestDeProprietes;
        private String _Nom;
        #endregion

        #region "Constructeur\Destructeur"

        public ExtPropriete() { }

        #endregion

        #region "Propriétés"

        public GestDeProprietes GestDeProprietes { get { Debug.Info(MethodBase.GetCurrentMethod());  return _GestDeProprietes; } }

        public String Nom
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                if (_GestDeProprietes.SwGestDeProprietes.Count > 0)
                {
                    foreach (String pNom in _GestDeProprietes.SwGestDeProprietes.GetNames())
                    {
                        if (pNom == _Nom)
                            return _Nom;
                    }
                }
                return null;
            }
        }

        public swCustomInfoType_e TypeDeLaPropriete
        {
            get
            {
                Debug.Info(MethodBase.GetCurrentMethod());
                if (!String.IsNullOrEmpty(this.Nom))
                {
                    return (swCustomInfoType_e)_GestDeProprietes.SwGestDeProprietes.GetType2(_Nom);
                }
                return swCustomInfoType_e.swCustomInfoUnknown;
            }
        }

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
            set { Debug.Info(MethodBase.GetCurrentMethod());  _GestDeProprietes.SwGestDeProprietes.Set(_Nom, value); }
        }

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

        internal Boolean EstInitialise { get { Debug.Info(MethodBase.GetCurrentMethod());  return _EstInitialise; } }

        #endregion

        #region "Méthodes"

        internal Boolean Init(GestDeProprietes Gestionnaire, String Nom)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            if ((Gestionnaire != null) && Gestionnaire.EstInitialise && !String.IsNullOrEmpty(Nom))
            {

                _GestDeProprietes = Gestionnaire;
                _Nom = Nom;

                if (!String.IsNullOrEmpty(this.Nom))
                {
                    Debug.Info(this.Nom);
                    _EstInitialise = true;
                }
                else
                {
                    _Nom = null;
                    Debug.Info("\t !!!!! Erreur d'initialisation");
                }
            }
            
            return _EstInitialise;
        }

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
            if (_GestDeProprietes.Modele.SW.VersionDeBase == "sw2013")
                _GestDeProprietes.SwGestDeProprietes.Get4(_Nom, true, out Expression, out Valeur);
            else
                _GestDeProprietes.SwGestDeProprietes.Get2(_Nom, out Expression, out Valeur);

        }

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

        int IComparable<ExtPropriete>.CompareTo(ExtPropriete Prop)
        {
            String Nom1 = _GestDeProprietes.Modele.SwModele.GetPathName() + _Nom;
            String Nom2 = Prop.GestDeProprietes.Modele.SwModele.GetPathName() + Prop.Nom;
            return Nom1.CompareTo(Nom2);
        }

        int IComparer<ExtPropriete>.Compare(ExtPropriete Prop1, ExtPropriete Prop2)
        {
            String Nom1 = Prop1.GestDeProprietes.Modele.SwModele.GetPathName() + Prop1.Nom;
            String Nom2 = Prop2.GestDeProprietes.Modele.SwModele.GetPathName() + Prop2.Nom;
            return Nom1.CompareTo(Nom2);
        }

        bool IEquatable<ExtPropriete>.Equals(ExtPropriete Prop)
        {
            String Nom1 = _GestDeProprietes.Modele.SwModele.GetPathName() + _Nom;
            String Nom2 = Prop.GestDeProprietes.Modele.SwModele.GetPathName() + Prop.Nom;
            return Nom1.Equals(Nom2);
        }

        #endregion
    }
}

