using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Reflection;

namespace Framework_SW2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("EA0CB46B-80A3-4EC7-92A4-50935002CAA9")]
    public interface IGestDeProprietes
    {
        CustomPropertyManager SwGestDeProprietes { get; }
        ExtModele Modele { get; }
        ExtPropriete AjouterPropriete(String Nom, swCustomInfoType_e TypePropriete, String Expression, Boolean EcraserExistante = false);
        Boolean SupprimerPropriete(String Nom);
        ExtPropriete RecupererPropriete(String Nom);
        ArrayList ListeDesProprietes(String NomARechercher = "");
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("F0F44B4A-BD1C-46E9-BB27-53A6D4342F6E")]
    [ProgId("Frameworks.GestDeConfigurations")]
    public class GestDeProprietes : IGestDeProprietes
    {
        #region "Variables locales"
        
        private Boolean _EstInitialise = false;

        private ExtModele _Modele;
        private CustomPropertyManager _SwGestDeProprietes;

        #endregion

        #region "Constructeur\Destructeur"

        public GestDeProprietes() { }

        #endregion

        #region "Propriétés"

        public CustomPropertyManager SwGestDeProprietes { get { Debug.Info(MethodBase.GetCurrentMethod());  return _SwGestDeProprietes; } }

        public ExtModele Modele { get { Debug.Info(MethodBase.GetCurrentMethod());  return _Modele; } }

        internal Boolean EstInitialise { get { Debug.Info(MethodBase.GetCurrentMethod());  return _EstInitialise; } }

        #endregion

        #region "Méthodes"

        internal Boolean Init(CustomPropertyManager SwGestionnaire, ExtModele Modele)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            if ((SwGestionnaire != null) && (Modele != null) && Modele.EstInitialise)
            {
                _SwGestDeProprietes = SwGestionnaire;
                _Modele = Modele;
                _EstInitialise = true;
            }
            else
            {
                Debug.Info("!!!!! Erreur d'initialisation");
            }

            return _EstInitialise;
        }

        public ExtPropriete AjouterPropriete(String Nom, swCustomInfoType_e TypePropriete, String Expression, Boolean EcraserExistante = false)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            // Si on écrase, on supprime la propriété
            if (EcraserExistante)
                _SwGestDeProprietes.Delete(Nom);

            // On la récupère
            ExtPropriete Propriete = RecupererPropriete(Nom);

            // Si elle n'existe pas on la créer et on lui assigne l'expression
            if (Propriete == null)
            {
                _SwGestDeProprietes.Add2(Nom, (int)TypePropriete, Expression);
                Propriete = new ExtPropriete();
                Propriete.Init(this, Nom);
            }

            // Si tout est ok, on la renvoi
            if (Propriete.EstInitialise)
                return Propriete;

            return null;
        }

        public ExtPropriete RecupererPropriete(String Nom)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            ExtPropriete Propriete = new ExtPropriete();

            if (Propriete.Init(this, Nom))
                return Propriete;

            return null;
        }

        public Boolean SupprimerPropriete(String Nom)
        {
            Debug.Info(MethodBase.GetCurrentMethod());

            if (_SwGestDeProprietes.Delete(Nom) == 1)
                return true;
            
            return false;
        }

        internal List<ExtPropriete> ListListeDesProprietes(String NomARechercher = "")
        {
            Debug.Info(MethodBase.GetCurrentMethod());

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
            Debug.Info(MethodBase.GetCurrentMethod());

            List<ExtPropriete> pListeProprietes = ListListeDesProprietes(NomARechercher);
            ArrayList pArrayProprietes = new ArrayList();

            if (pListeProprietes.Count > 0)
                pArrayProprietes = new ArrayList(pListeProprietes);

            return pArrayProprietes;
        }

        #endregion
    }
}
