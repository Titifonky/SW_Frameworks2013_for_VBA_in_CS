using System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace Framework
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("EA0CB46B-80A3-4EC7-92A4-50935002CAA9")]
    public interface IeGestDeProprietes
    {
        CustomPropertyManager SwGestDeProprietes { get; }
        eModele Modele { get; }
        ePropriete AjouterPropriete(String Nom, swCustomInfoType_e TypePropriete, String Expression, Boolean EcraserExistante = false);
        ePropriete RecupererPropriete(String Nom);
        Boolean SupprimerPropriete(String Nom);
        Boolean ProprieteExiste(String Nom);
        ArrayList ListeDesProprietes(String NomARechercher = "");
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("F0F44B4A-BD1C-46E9-BB27-53A6D4342F6E")]
    [ProgId("Frameworks.eGestDeProprietes")]
    public class eGestDeProprietes : IeGestDeProprietes
    {
#region "Variables locales"
        
        private Boolean _EstInitialise = false;

        private eModele _Modele = null;
        private CustomPropertyManager _SwGestDeProprietes = null;

#endregion

#region "Constructeur\Destructeur"

        public eGestDeProprietes() { }

#endregion

#region "Propriétés"

        /// <summary>
        /// Retourne le gestionnaire CustomPropertyManager associé
        /// </summary>
        public CustomPropertyManager SwGestDeProprietes { get { Debug.Print(MethodBase.GetCurrentMethod());  return _SwGestDeProprietes; } }

        /// <summary>
        /// Retourne le parent ExtModele 
        /// </summary>
        public eModele Modele { get { Debug.Print(MethodBase.GetCurrentMethod());  return _Modele; } }

        /// <summary>
        /// Fonction interne
        /// Test l'initialisation de l'objet ExtModele
        /// </summary>
        internal Boolean EstInitialise { get { Debug.Print(MethodBase.GetCurrentMethod());  return _EstInitialise; } }

#endregion

#region "Méthodes"

        /// <summary>
        /// Méthode interne
        /// Initialise l'objet GestDeProprietes
        /// </summary>
        /// <param name="SwGestionnaire"></param>
        /// <param name="Modele"></param>
        /// <returns></returns>
        internal Boolean Init(CustomPropertyManager SwGestionnaire, eModele Modele)
        {
            Debug.Print(MethodBase.GetCurrentMethod());

            if ((SwGestionnaire != null) && (Modele != null) && Modele.EstInitialise)
            {
                _SwGestDeProprietes = SwGestionnaire;
                _Modele = Modele;
                _EstInitialise = true;
            }
            else
            {
                Debug.Print("!!!!! Erreur d'initialisation");
            }

            return _EstInitialise;
        }

        /// <summary>
        /// Ajouter une propriété
        /// </summary>
        /// <param name="Nom"></param>
        /// <param name="TypePropriete"></param>
        /// <param name="Expression"></param>
        /// <param name="EcraserExistante"></param>
        /// <returns></returns>
        public ePropriete AjouterPropriete(String Nom, swCustomInfoType_e TypePropriete, String Expression, Boolean EcraserExistante = false)
        {
            Debug.Print(MethodBase.GetCurrentMethod());

            // Si on écrase, on supprime la propriété
            if (EcraserExistante)
                _SwGestDeProprietes.Delete(Nom);

            // On la récupère
            ePropriete Propriete = RecupererPropriete(Nom);

            // Si elle n'existe pas on la crée et on lui assigne l'expression
            if (Propriete == null)
            {
                _SwGestDeProprietes.Add2(Nom, (int)TypePropriete, Expression);
                Propriete = new ePropriete();
                Propriete.Init(this, Nom);
            }

            // Si tout est ok, on la renvoi
            if (Propriete.EstInitialise)
                return Propriete;

            return null;
        }

        /// <summary>
        /// Récupérer une propriété
        /// </summary>
        /// <param name="Nom"></param>
        /// <returns></returns>
        public ePropriete RecupererPropriete(String Nom)
        {
            Debug.Print(MethodBase.GetCurrentMethod());

            ePropriete Propriete = new ePropriete();

            if (Propriete.Init(this, Nom))
                return Propriete;

            return null;
        }

        /// <summary>
        /// Supprime une propriété
        /// </summary>
        /// <param name="Nom"></param>
        /// <returns></returns>
        public Boolean SupprimerPropriete(String Nom)
        {
            Debug.Print(MethodBase.GetCurrentMethod());

            if (_SwGestDeProprietes.Delete(Nom) == 1)
                return true;
            
            return false;
        }

        /// <summary>
        /// Teste l'existence d'une propriété
        /// </summary>
        /// <param name="Nom"></param>
        /// <returns></returns>
        public Boolean ProprieteExiste(String Nom)
        {
            ePropriete pProp = new ePropriete();

            return pProp.Init(this, Nom);
        }

        /// <summary>
        /// Méthode interne
        /// Renvoi la liste des propriétés filtrée par les arguments
        /// Si NomARechercher est vide, toutes les propriétés sont renvoyées
        /// </summary>
        /// <param name="NomARechercher"></param>
        /// <returns></returns>
        internal List<ePropriete> ListListeDesProprietes(String NomARechercher = "")
        {
            Debug.Print(MethodBase.GetCurrentMethod());

            List<ePropriete> pListeProps = new List<ePropriete>();

            if (_SwGestDeProprietes.Count > 0)
            {
                foreach (String pNom in _SwGestDeProprietes.GetNames())
                {
                    ePropriete Prop = new ePropriete();
                    if (Prop.Init(this, pNom) && Regex.IsMatch(pNom, NomARechercher))
                        pListeProps.Add(Prop);
                }
            }

            return pListeProps;
        }

        /// <summary>
        /// Renvoi la liste des propriétés filtrée par les arguments
        /// Si NomARechercher est vide, toutes les propriétés sont renvoyées
        /// </summary>
        /// <param name="NomARechercher"></param>
        /// <returns></returns>
        public ArrayList ListeDesProprietes(String NomARechercher = "")
        {
            Debug.Print(MethodBase.GetCurrentMethod());

            List<ePropriete> pListeProprietes = ListListeDesProprietes(NomARechercher);
            ArrayList pArrayProprietes = new ArrayList();

            if (pListeProprietes.Count > 0)
                pArrayProprietes = new ArrayList(pListeProprietes);

            return pArrayProprietes;
        }

#endregion
    }
}
