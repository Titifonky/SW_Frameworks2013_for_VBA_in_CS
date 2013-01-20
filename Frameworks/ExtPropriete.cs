using System;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

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
    public class ExtPropriete : IExtPropriete
    {
        #region "Variables locales"
        private Debug _Debug = Debug.Instance;
        private Boolean _EstInitialise = false;

        private GestDeProprietes _GestDeProprietes;
        private String _Nom;
        #endregion

        #region "Constructeur\Destructeur"

        public ExtPropriete() { }

        #endregion

        #region "Propriétés"

        public GestDeProprietes GestDeProprietes { get { return _GestDeProprietes; } }

        public String Nom
        {
            get
            {
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
                String Expression;
                String Valeur;

                Get(out Expression, out Valeur);

                return Expression;
            }
            set { _GestDeProprietes.SwGestDeProprietes.Set(_Nom, value); }
        }

        public String Valeur
        {
            get
            {
                String Expression;
                String Valeur;

                Get(out Expression, out Valeur);

                return Valeur;
            }
        }

        internal Boolean EstInitialise { get { return _EstInitialise; } }

        #endregion

        #region "Méthodes"

        internal Boolean Init(GestDeProprietes Gestionnaire, String Nom)
        {
            _MethodBase Methode = System.Reflection.MethodBase.GetCurrentMethod();

            if ((Gestionnaire != null) && Gestionnaire.EstInitialise && !String.IsNullOrEmpty(Nom))
            {
                

                _GestDeProprietes = Gestionnaire;
                _Nom = Nom;

                if (!String.IsNullOrEmpty(this.Nom))
                {
                    _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name);
                    _EstInitialise = true;
                }
                else
                {
                    _Nom = null;
                    _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name + " : Erreur d'initialisation");
                }
            }
            
            return _EstInitialise;
        }

        private void Get(out String Expression, out String Valeur)
        {
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
    }
}

