using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using System.Text.RegularExpressions;
using SolidWorks.Interop.swconst;

namespace Framework_SW2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("EA0CB46B-80A3-4EC7-92A4-50935002CAA9")]
    public interface IGestDeProprietes
    {
        CustomPropertyManager SwGestDeProprietes { get; }
        ExtModele Modele { get; }
        Boolean AjouterPropriete(String NomPropriete, swCustomInfoType_e TypePropriete, String ValeurPropriete);
        Boolean SupprimerPropriete(String NomPropriete);
        String RecupererPropriete(String NomPropriete);
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("F0F44B4A-BD1C-46E9-BB27-53A6D4342F6E")]
    [ProgId("Frameworks.GestDeConfigurations")]
    public class GestDeProprietes : IGestDeProprietes
    {
        #region "Variables locales"
        private Debug _Debug = Debug.Instance;
        private Boolean _EstInitialise = false;

        private ExtModele _Modele;
        private CustomPropertyManager _swGestDeProprietes;

        #endregion

        #region "Constructeur\Destructeur"

        public GestDeProprietes() { }

        #endregion

        #region "Propriétés"

        public CustomPropertyManager SwGestDeProprietes { get { return _swGestDeProprietes; } }

        public ExtModele Modele { get { return _Modele; } }

        internal Boolean EstInitialise { get { return _EstInitialise; } }

        #endregion

        #region "Méthodes"

        internal Boolean Init(CustomPropertyManager SwGestionnaire, ExtModele Modele)
        {
            _MethodBase Methode = System.Reflection.MethodBase.GetCurrentMethod();

            if ((SwGestionnaire != null) && (Modele != null) && Modele.EstInitialise)
            {
                _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name);

                _swGestDeProprietes = SwGestionnaire;
                _Modele = Modele;
                _EstInitialise = true;
            }
            else
            {
                _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name + " : Erreur d'initialisation");
            }

            return _EstInitialise;
        }

        Boolean AjouterPropriete(String Nom, swCustomInfoType_e TypeProp, String Valeur)
        {
            SupprimerPropriete(Nom);
            if (_swGestDeProprietes.Add2(Nom, (int)TypeProp, Valeur) == 0)
                return false;
            else
                return true;
        }

        Boolean SupprimerPropriete(String Nom)
        {
            if (_swGestDeProprietes.Delete(Nom) == 1)
                return false;
            else
                return true;
        }

        String RecupererPropriete(String Nom)
        {
            String Val;
            String ResVal;

            // Pour la compatibilité
            if (_Modele.SW.VersionDeBase == "sw2013")
                _swGestDeProprietes.Get4(Nom, true, out Val, out ResVal);
            else
                _swGestDeProprietes.Get2(Nom, out Val, out ResVal);

            return ResVal;
        }

        #endregion
    }
}
