using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace Framework2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("A8C91882-5820-11E2-A1E0-98046188709B")]
    public interface IExtRecherche
    {
        ExtComposant Composant { get;}
        Boolean PrendreEnCompteConfig { get; set; }
        Boolean PrendreEnCompteExclus { get; set; }
        Boolean PrendreEnCompteSupprime { get; set; }
        Boolean Init(ExtComposant Composant);
        String NomCle(ExtComposant Composant);
        ArrayList Lancer(TypeFichier_e TypeComposant, String NomComposant = "");
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("AD74AE82-5820-11E2-9F7B-99046188709B")]
    [ProgId("Frameworks.ExtRecherche")]
    public class ExtRecherche : IExtRecherche
    {
        #region "Variables locales"
        private Debug _Debug = Debug.Instance;

        private ExtComposant _Composant;
        private Boolean _PrendreEnCompteConfig = true;
        private Boolean _PrendreEnCompteExclus = false;
        private Boolean _PrendreEnCompteSupprime = false;
        private List<ExtComposant> _ListeComposants;

        #endregion

        #region "Constructeur\Destructeur"

        public ExtRecherche()
        {
        }

        #endregion

        #region "Propriétés"

        public ExtComposant Composant
        {
            get { return _Composant; }
        }

        public Boolean PrendreEnCompteConfig
        {
            get { return _PrendreEnCompteConfig; }
            set { _PrendreEnCompteConfig = value; }
        }

        public Boolean PrendreEnCompteExclus
        {
            get { return _PrendreEnCompteExclus; }
            set { _PrendreEnCompteExclus = value; }
        }

        public Boolean PrendreEnCompteSupprime
        {
            get { return _PrendreEnCompteSupprime; }
            set { _PrendreEnCompteSupprime = value; }
        }

        #endregion

        #region "Méthodes"

        public Boolean Init(ExtComposant Composant)
        {
            _MethodBase Methode = System.Reflection.MethodBase.GetCurrentMethod();

            if (!(Composant == null))
            {
                _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name);

                _Composant = Composant;
                return true;
            }

            _Debug.DebugAjouterLigne(this.GetType().Name + "." + Methode.Name + " : Erreur d'initialisation");
            return false;
        }

        public String NomCle(ExtComposant Composant)
        {
            String pNomCle = Composant.Modele.Chemin;
            if (_PrendreEnCompteConfig)
                pNomCle = pNomCle + " " + Composant.Configuration.Nom;

            return pNomCle;
        }

        private void RecListComposants(ExtComposant ComposantRacine, TypeFichier_e TypeComposant, String NomComposant = "")
        {
            foreach (ExtComposant Comp in ComposantRacine.ComposantsEnfants(_PrendreEnCompteSupprime))
            {
                if (!Comp.EstExclu | _PrendreEnCompteExclus)
                {
                    if ((Comp.Modele.TypeDuModele == TypeComposant) && (Path.GetFileName(Comp.Modele.Chemin) == NomComposant))
                    {
                        if (_ListeComposants.Exists(C => NomCle(C) == NomCle(Comp)))
                        {
                            ExtComposant Composant = _ListeComposants.Find
                        }

                    }
                }
            }

            //For Each ComposantListe In ComposantRacine.ListedesComposantsEnfants(PrendreEnCompteSupprime)
        
            //    If (ComposantListe.EstExclu Imp pPrendreEnCompteExclus) Then
            
            //        If ComposantListe.Modele.Est(TypeComposant) And (ComposantListe.Modele.Fichier.NomDuFichier Like NomComposant) Then
                
            //            Cle = NomCle(ComposantListe)
                
            //            If CleExiste(pCollPieces, Cle) Then
            //                Set Composant = pCollPieces.Item(Cle)
            //                Composant.Nb = Composant.Nb + 1
            //            Else
            //                Set Composant = New ExtComposant
            //                Set Composant = ComposantListe
            //                Composant.Nb = 1
            //                pCollPieces.Add Composant, Cle
            //            End If
                    
            //        End If
            
            //        If ComposantListe.Modele.Est(cAssemblage) And Not (ComposantListe.EstSupprime) Then
            //            ListerLesComposants ComposantListe, TypeComposant, NomComposant
            //        End If
            
            //    End If
        
            //Next ComposantListe
        }

        internal List<ExtComposant> ListComposants(TypeFichier_e TypeComposant, String NomComposant = "")
        {
            _ListeComposants = new List<ExtComposant>();

            switch (_Composant.Modele.TypeDuModele)
            {
                case TypeFichier_e.cAssemblage :
                    if(_Composant.swComposant.IGetChildrenCount() == 0)
                    {
                        RecListComposants(_Composant, TypeComposant, NomComposant + "*");
                    }
                    break;
                default:
                    _ListeComposants.Add(_Composant);
                    break;
            }

            _ListeComposants.Sort();

            return _ListeComposants;
        }

        public ArrayList Lancer(TypeFichier_e TypeComposant, String NomComposant = "")
        {
            List<ExtComposant> pListeComps = ListComposants(TypeComposant, NomComposant);
            ArrayList pArrayComps = new ArrayList();

            if (pListeComps.Count > 0)
                pArrayComps = new ArrayList(pListeComps);

            return pArrayComps;
        }

        #endregion
    }
}
