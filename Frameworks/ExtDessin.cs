using System;
using System.Runtime.InteropServices;

namespace Framework2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("795533EC-5820-11E2-875F-83046188709B")]
    public interface IExtDessin
    {
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("7E7EFEE8-5820-11E2-B8E9-84046188709B")]
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
