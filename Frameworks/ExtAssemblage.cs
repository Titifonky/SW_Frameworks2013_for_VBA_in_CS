using System;
using System.Runtime.InteropServices;

namespace Frameworks2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("532A9140-513F-11E2-9642-A0DC6088709B")]
    public interface IExtAssemblage
    {
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("CDF7E002-5136-11E2-9D4E-44CD6088709B")]
    [ProgId("Frameworks.ExtAssemblage")]
    public class ExtAssemblage : ExtModele , IExtAssemblage
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
