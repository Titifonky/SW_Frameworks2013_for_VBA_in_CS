using System;
using System.Runtime.InteropServices;

namespace Frameworks2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("4947C5B2-513F-11E2-8ED3-9FDC6088709B")]
    public interface IExtPiece
    {
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("430AF246-513F-11E2-8108-7ADC6088709B")]
    [ProgId("Frameworks.ExtAssemblage")]
    class ExtPiece : ExtModele, IExtPiece
    {
        #region "Variables locales"
        #endregion

        #region "Constructeur\Destructeur"

        public ExtPiece()
        {
        }

        #endregion

        #region "Propriétés"
        #endregion

        #region "Méthodes"
        #endregion
    }
}
