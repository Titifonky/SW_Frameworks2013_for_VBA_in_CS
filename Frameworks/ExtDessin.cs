using System;
using System.Runtime.InteropServices;

namespace Frameworks2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("ED380DD2-551B-11E2-9642-634E6188709B")]
    public interface IExtDessin
    {
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("F4965E62-551B-11E2-92B7-6B4E6188709B")]
    [ProgId("Frameworks.ExtDessin")]
    public class ExtDessin : ExtModele, IExtDessin
    {
        #region "Variables locales"
        #endregion

        #region "Constructeur\Destructeur"

        public ExtDessin()
        {
        }

        #endregion

        #region "Propriétés"
        #endregion

        #region "Méthodes"
        #endregion
    }
}
