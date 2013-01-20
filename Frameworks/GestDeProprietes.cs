using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace Framework_SW2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("EA0CB46B-80A3-4EC7-92A4-50935002CAA9")]
    public interface IGestDeProprietes
    {
        CustomPropertyManager SwGestDeProprietes { get; }
        ExtModele Modele { get; }
        ExtPropriete AjouterPropriete(String Nom, swCustomInfoType_e TypePropriete, String Expression, Boolean EcraserExistante = false);
        ExtPropriete RecupererPropriete(String Nom);
        ArrayList ListeDesProprietes(String NomARechercher = "");
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
        private CustomPropertyManager _SwGestDeProprietes;

        #endregion

        #region "Constructeur\Destructeur"

        public GestDeProprietes() { }

        #endregion

        #region "Propriétés"

        public CustomPropertyManager SwGestDeProprietes { get { return _SwGestDeProprietes; } }

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

                _SwGestDeProprietes = SwGestionnaire;
                _Modele = Modele;
                _EstInitialise = true;
            }
            else
            {
                _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name + " : Erreur d'initialisation");
            }

            return _EstInitialise;
        }

        /// <summary>
        /// Ajoute une propriété à la piece.
        /// </summary>
        /// <param name="Nom"></param>
        /// <param name="TypePropriete"></param>
        /// <param name="Expression"></param>
        /// <returns></returns>
        public ExtPropriete AjouterPropriete(String Nom, swCustomInfoType_e TypePropriete, String Expression, Boolean EcraserExistante = false)
        {
            if (EcraserExistante)
            {
                _SwGestDeProprietes.Delete(Nom);
            }

            ExtPropriete Propriete = new ExtPropriete();

            // On initialise la propriete. Si elle n'existe pas on la créer.
            
            if ((Propriete.Init(this, Nom) == false) && (_SwGestDeProprietes.Add2(Nom, (int)TypePropriete, Expression) == 1))
            {
                // Si la propriete a été crée, on lui passe la bonne Expression
                if (Propriete.Init(this, Nom))
                    Propriete.Expression = Expression;
            }

            // Si tout est ok, on la renvoi
            if (Propriete.EstInitialise)
                return Propriete;

            return null;
        }

        public ExtPropriete RecupererPropriete(String Nom)
        {
            ExtPropriete Propriete = new ExtPropriete();

            if (Propriete.Init(this, Nom))
                return Propriete;

            return null;
        }

        internal List<ExtPropriete> ListListeDesProprietes(String NomARechercher = "")
        {
            List<ExtPropriete> pListeProps = new List<ExtPropriete>();

            if (_SwGestDeProprietes.Count > 0)
            {
                foreach (String pNom in _SwGestDeProprietes.GetNames())
                {
                    ExtPropriete Prop = new ExtPropriete();
                    if (Prop.Init(this, pNom) && Regex.IsMatch(pNom, NomARechercher))
                        pListeProps.Add(Prop);
                }
            }

            return pListeProps;
        }

        public ArrayList ListeDesProprietes(String NomARechercher = "")
        {
            _MethodBase Methode = System.Reflection.MethodBase.GetCurrentMethod();
            _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name);

            List<ExtPropriete> pListeProprietes = ListListeDesProprietes(NomARechercher);
            ArrayList pArrayProprietes = new ArrayList();

            if (pListeProprietes.Count > 0)
                pArrayProprietes = new ArrayList(pListeProprietes);

            return pArrayProprietes;
        }

        #endregion
    }
}
