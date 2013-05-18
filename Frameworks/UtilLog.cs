using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using System.Collections.Generic;

namespace Framework_SW2013
{
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    [Guid("8B05EE8A-688C-4FC6-8614-B29C69D52323")]
    public interface ILog
    {
        Boolean Actif { get; set; }
        Boolean Init(eModele Modele, String NomMacro = "");
        void Info(String Message, int Tab = 0);
    }

    [ClassInterface(ClassInterfaceType.None)]
    [Guid("AF0E5AF8-0E3C-47D0-A85F-BB1BCCE26DD8")]
    [ProgId("Frameworks.Log")]
    public class Log : ILog
    {
        #region "Variables locales"

        private Boolean _ACTIF = true;
        private Boolean _INIT = false;

        private String _Dossier;
        private String _Fichier;
        private String _Chemin_Fichier;

        #endregion

        #region "Constructeur\Destructeur"

        public Log() { }

        #endregion

        #region "Propriétés"

        public Boolean Actif { get { return _ACTIF; } set { _ACTIF = value; } }

        #endregion

        #region "Méthodes"

        public Boolean Init(eModele Modele, String NomMacro = "")
        {

            if ((Modele != null) && (Modele.EstInitialise) && _ACTIF)
            {
                _Dossier = Modele.FichierSw.NomDuDossier;

                _Fichier = "Log" + " _Macro " + NomMacro + " _Fichier " + Modele.FichierSw.NomDuFichierSansExt + ".txt";

                _Chemin_Fichier = Path.Combine(_Dossier, _Fichier);

                File.Delete(_Chemin_Fichier);

                StreamWriter pFichierLog = new StreamWriter(_Chemin_Fichier, false, System.Text.Encoding.Unicode);

                pFichierLog.WriteLine(Date() + " -> Start Log");

                pFichierLog.Close();

                _INIT = true;
            }

            return _INIT;

        }

        public void Info(String Message, int Tab = 0)
        {
            if (!(_ACTIF && _INIT))
                return;

            try
            {
                StreamWriter pFichierDebug = new StreamWriter(_Chemin_Fichier, true, System.Text.Encoding.Unicode);
                pFichierDebug.WriteLine(Date() + "  " + "\t".Repeter(Tab + 1) + Message);
                pFichierDebug.Close();
            }
            catch
            {
                StreamWriter pFichierDebug = new StreamWriter(_Chemin_Fichier, true, System.Text.Encoding.Unicode);
                pFichierDebug.WriteLine(Date() + " -> Erreur");
                pFichierDebug.Close();
            }

        }

        private String Date()
        {
            return String.Format("{0:0000}/{1:00}/{2:00} {3:00}:{4:00}:{5:00}.{6:000}",
                System.DateTime.Now.Year, System.DateTime.Now.Month, System.DateTime.Now.Day,
                System.DateTime.Now.Hour, System.DateTime.Now.Minute, System.DateTime.Now.Second, System.DateTime.Now.Millisecond);
        }

        #endregion
    }
}

