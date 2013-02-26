using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;

namespace Framework_SW2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("CACBFAD0-5820-11E2-B60C-A9046188709B")]
    public interface IExtAssemblage
    {
        AssemblyDoc SwAssemblage { get; }
        ExtModele Modele { get; }
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("CF8CC568-5820-11E2-B525-AA046188709B")]
    [ProgId("Frameworks.ExtAssemblage")]
    public class ExtAssemblage : IExtAssemblage
    {
        #region "Variables locales"
        private Debug _Debug = Debug.Instance;
        private Boolean _EstInitialise = false;

        private ExtModele _Modele;
        private AssemblyDoc _SwAssemblage;
        #endregion

        #region "Constructeur\Destructeur"

        public ExtAssemblage() { }

        #endregion

        #region "Propriétés"

        public AssemblyDoc SwAssemblage { get { return _SwAssemblage; } }

        public ExtModele Modele { get { return _Modele; } }

        internal Boolean EstInitialise { get { return _EstInitialise; } }

        #endregion

        #region "Méthodes"

        internal Boolean Init(ExtModele Modele)
        {
            _Debug.DebugAjouterLigne(this.GetType().Name + "." + System.Reflection.MethodBase.GetCurrentMethod().Name);

            if ((Modele != null) && Modele.EstInitialise && (Modele.TypeDuModele == TypeFichier_e.cAssemblage))
            {
                _Debug.DebugAjouterLigne("\t -> " + Modele.Chemin);

                _Modele = Modele;
                _SwAssemblage = Modele.SwModele as AssemblyDoc;
                _EstInitialise = true;
            }
            else
            {
                _Debug.DebugAjouterLigne("\t !!!!! Erreur d'initialisation");
            }

            return _EstInitialise;
        }

        #endregion
    }
}
