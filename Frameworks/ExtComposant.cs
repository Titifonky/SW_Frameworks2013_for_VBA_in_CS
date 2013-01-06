using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace Framework2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("BF2ED17A-5820-11E2-8160-9F046188709B")]
    public interface IExtComposant
    {
        Component2 swComposant { get; }
        ExtModele Modele { get; }
        ExtConfiguration Configuration { get; }
        int Nb { get; set; }
        Boolean EstExclu { get; set; }
        Boolean EstSupprime { get; }
        Boolean Init(Component2 Composant, ExtModele Modele);
        ArrayList ComposantsEnfants(Boolean PrendreEnCompteSupprime = false);
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("C46318AE-5820-11E2-A863-A3046188709B")]
    [ProgId("Frameworks.ExtComposant")]
    public class ExtComposant : IExtComposant, IComparable<ExtComposant>, IComparer<ExtComposant>
    {
        #region "Variables locales"

        private Component2 _swComposant;
        private ExtModele _Modele;
        private ExtDebug _Debug;
        private ExtConfiguration _Configuration;
        private int _Nb = 0;

        #endregion

        #region "Constructeur\Destructeur"

        public ExtComposant()
        {
        }

        #endregion

        #region "Propriétés"

        public Component2 swComposant
        {
            get { return _swComposant; }
        }

        public ExtModele Modele
        {
            get { return _Modele; }
        }

        public ExtConfiguration Configuration
        {
            get { return _Configuration; }
        }

        public int Nb
        {
            get { return _Nb; }
            set { _Nb = value; }
        }

        public Boolean EstExclu
        {
            get
            {
                if (!(_swComposant.ExcludeFromBOM == false))
                    return true;
                return false;
            }
            set { _swComposant.ExcludeFromBOM = value; }
        }

        public Boolean EstSupprime
        {
            get
            {
                if (!(_swComposant.IsSuppressed() == false))
                    return true;
                return false;
            }
        }

        #endregion

        #region "Méthodes"

        public Boolean Init(Component2 Composant, ExtModele Modele)
        {
            _MethodBase Methode = System.Reflection.MethodBase.GetCurrentMethod();

            if (!((Composant.Equals(null)) && (Modele.Equals(null))))
            {
                _Debug = Modele.SW.Debug;
                _Debug.ErreurAjouterLigne(this.GetType().Name + "." + Methode.Name);

                _swComposant = Composant;
                _Modele = Modele;
                _Nb = 1;

                Configuration Config;
                if (String.IsNullOrEmpty(_swComposant.ReferencedConfiguration))
                    Config = _Modele.swModele.GetActiveConfiguration();
                else
                    Config = _Modele.swModele.GetConfigurationByName(_swComposant.ReferencedConfiguration);

                _Configuration = new ExtConfiguration();
                _Configuration.Init(Config, _Modele);
                return true;
            }

            return false;
        }

        internal List<ExtComposant> ListComposantsEnfants(Boolean PrendreEnCompteSupprime = false)
        {
            List<ExtComposant> Liste = new List<ExtComposant>();
            
            if (_swComposant.IGetChildrenCount() == 0)
                return Liste;

            foreach (Component2 Composant in _swComposant.GetChildren())
            {
                /// Si le composant est supprimé mais qu'on a decidé de le prendre en compte, c'est bon
                if ((Composant.IsSuppressed() == false) | PrendreEnCompteSupprime)
                {
                    ExtModele ModeleExt = _Modele.SW.Modele(Composant.GetPathName());
                    ExtComposant CompExt = new ExtComposant();
                    CompExt.Init(Composant, ModeleExt);
                    ModeleExt.Composant = CompExt;
                    _Modele.SW.Debug.ErreurAjouterLigne(ModeleExt.Chemin);
                    _Modele.SW.Debug.ExecutionAjouterLigne(ModeleExt.Chemin);
                    Liste.Add(CompExt);
                }
            }

            Liste.Sort();
            return Liste;

        }

        public ArrayList ComposantsEnfants(Boolean PrendreEnCompteSupprime = false)
        {
            List<ExtComposant> pListeComps = ListComposantsEnfants(PrendreEnCompteSupprime);
            ArrayList pArrayComps = new ArrayList();

            if (pListeComps.Count > 0)
                pArrayComps = new ArrayList(pListeComps);

            return pArrayComps;
        }

        #endregion

        int IComparable<ExtComposant>.CompareTo(ExtComposant Comp)
        {
            return _swComposant.GetPathName().CompareTo(Comp._swComposant.GetPathName());
        }

        int IComparer<ExtComposant>.Compare(ExtComposant x, ExtComposant y)
        {
            return x._Modele.swModele.GetPathName().CompareTo(y._Modele.swModele.GetPathName());
        }
    }
}