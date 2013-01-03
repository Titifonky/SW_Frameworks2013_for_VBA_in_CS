using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace Frameworks2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("098E5D1A-5585-11E2-9230-89E06188709B")]
    public interface IExtComposant
    {
        Component2 swComposant { get; }
        ExtModele Modele { get; }
        ExtConfiguration Configuration { get; }
        int Nb { get; set; }
        Boolean EstExclu { get; set; }
        Boolean EstSupprime { get; }
        Boolean Init(Component2 Composant, ExtModele Modele);
        Array ComposantsEnfants(Boolean PrendreEnCompteSupprime = false);
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("10679F3E-5585-11E2-B541-8FE06188709B")]
    [ProgId("Frameworks.ExtComposant")]
    public class ExtComposant : IExtComposant, IComparable<ExtComposant>, IComparer<ExtComposant>
    {
        #region "Variables locales"

        private Component2 _swComposant;
        private ExtModele _Modele;
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
            if (!((Composant.Equals(null)) && (Modele.Equals(null))))
            {
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
            
            foreach (Component2 Composant in _swComposant.GetChildren())
            {
                if ((Composant.IsSuppressed() == false) | PrendreEnCompteSupprime)
                {
                    ExtModele ModeleExt = _Modele.SW.Modele(Composant.GetPathName());
                    ExtComposant CompExt = new ExtComposant();
                    CompExt.Init(Composant, ModeleExt);
                    ModeleExt.Composant = CompExt;
                    Liste.Add(CompExt);
                }
            }

            Liste.Sort();
            return Liste;

        }

        public Array ComposantsEnfants(Boolean PrendreEnCompteSupprime = false)
        {
            //Array Arr = ListComposantsEnfants(PrendreEnCompteSupprime).ToArray();
            return ListComposantsEnfants(PrendreEnCompteSupprime).ToArray();
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