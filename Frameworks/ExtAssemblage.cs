using System;
using System.Runtime.InteropServices;

namespace Frameworks2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("CACBFAD0-5820-11E2-B60C-A9046188709B")]
    public interface IExtAssemblage
    {
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("CF8CC568-5820-11E2-B525-AA046188709B")]
    [ProgId("Frameworks.ExtAssemblage")]
    public class ExtAssemblage : IExtAssemblage
    {
        #region "Variables locales"
        #endregion

        #region "Constructeur\Destructeur"

        public ExtAssemblage()
        {
        }

        #endregion

        #region "Propriétés"
        #endregion

        #region "Méthodes"
        #endregion
    }
}
