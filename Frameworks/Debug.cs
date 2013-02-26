using System;
using System.IO;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace Framework_SW2013
{
    internal sealed class Debug
    {
        private static Debug _instance = null;
        private static readonly object _Lock = new object();

        private ExtSldWorks _SW;
        private String _CheminFichierDebug;
        private Int32 _NbTab = 0;
        private String _Tab = "    ";

        #region "Propriétés"

        public static Debug Instance
        {
            get
            {
                lock (_Lock)
                {
                    if (_instance == null)
                    {
                        _instance = new Debug();
                    }
                }
                return _instance;
            }
        }

        #endregion

        #region "Constructeur\Destructeur"

        public Debug()
        {
        }

        #endregion

        #region "Méthodes"

        public Boolean Init(ExtSldWorks SW)
        {
            if (SW != null)
            {
                _SW = SW;
                String pDossierMacros = _SW.SwSW.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swFileLocationsMacros);

                _CheminFichierDebug = Path.Combine(pDossierMacros, "Debug.txt");
                StreamWriter pFichierDebug = new StreamWriter(_CheminFichierDebug, false, System.Text.Encoding.Unicode);
                pFichierDebug.Close();

                return true;
            }

            return false;
        }

        public void DebugAjouterLigne(String Ligne)
        {
            StreamWriter pFichierDebug = new StreamWriter(_CheminFichierDebug, true, System.Text.Encoding.Unicode);
            
            pFichierDebug.WriteLine(_Tab.Repeter(_NbTab) + Ligne);
            pFichierDebug.Close();
        }

        public void DebugFin()
        {
        }

        #endregion
    }
}
