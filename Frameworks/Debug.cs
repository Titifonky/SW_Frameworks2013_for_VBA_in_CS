using System;
using System.IO;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Diagnostics;

namespace Framework_SW2013
{
    internal static class Debug
    {
        #region "Propriétés"

        private static String DEBUG_FICHIER = "";
        private static Boolean DEBUG_ACTIF = true;
        private static String _InfoMethode = "";

        public static Boolean Actif { get { return DEBUG_ACTIF; } set { DEBUG_ACTIF = value; } }

        #endregion

        #region "Méthodes"

        public static Boolean Init(SldWorks SW)
        {
            if (SW != null)
            {
                String pDossierMacros = SW.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swFileLocationsMacros);

                String pCheminFichierDebug = Path.Combine(pDossierMacros, System.Reflection.Assembly.GetCallingAssembly().FullName.Split(',')[0] + "_Debug.txt");
                File.Delete(pCheminFichierDebug);
                StreamWriter pFichierDebug = new StreamWriter(pCheminFichierDebug, false, System.Text.Encoding.Unicode);
                pFichierDebug.WriteLine("Start Debug");
                pFichierDebug.Close();

                DEBUG_FICHIER = pCheminFichierDebug;

                return true;
            }

            return false;
        }

        /// <summary>
        /// Si Message est vide, écrit le nom de l'objet et de la méthode.
        /// Sinon, écrit Message.
        /// </summary>
        /// <param name="Message"></param>
        public static void Info(String Message)
        {
            try
            {
                if (DEBUG_ACTIF == false)
                    return;

                StackTrace pPile = new StackTrace();

                StreamWriter pFichierDebug = new StreamWriter(DEBUG_FICHIER, true, System.Text.Encoding.Unicode);
                pFichierDebug.WriteLine("\t".Repeter(pPile.FrameCount - 1) + _InfoMethode + " : " + Message);
                pFichierDebug.Close();
            }
            catch
            {
                StreamWriter pFichierDebug = new StreamWriter(DEBUG_FICHIER, true, System.Text.Encoding.Unicode);
                pFichierDebug.WriteLine("Erreur");
                pFichierDebug.Close();
            }

        }

        public static void Info(_MethodBase Methode)
        {
            try
            {
                if (DEBUG_ACTIF == false)
                    return;

                StackTrace pPile = new StackTrace();
                _InfoMethode = Methode.DeclaringType.ToString().Split('.')[1] + "." + Methode.Name;

                StreamWriter pFichierDebug = new StreamWriter(DEBUG_FICHIER, true, System.Text.Encoding.Unicode);
                pFichierDebug.WriteLine("\t".Repeter(pPile.FrameCount - 1) + _InfoMethode);
                pFichierDebug.Close();
            }
            catch
            {
                StreamWriter pFichierDebug = new StreamWriter(DEBUG_FICHIER, true, System.Text.Encoding.Unicode);
                pFichierDebug.WriteLine("Erreur");
                pFichierDebug.Close();
            }

        }

        public static void Info(String Message, _MethodBase Methode)
        {
            try
            {
                if (DEBUG_ACTIF == false)
                    return;

                StackTrace pPile = new StackTrace();

                StreamWriter pFichierDebug = new StreamWriter(DEBUG_FICHIER, true, System.Text.Encoding.Unicode);
                pFichierDebug.WriteLine("\t".Repeter(pPile.FrameCount - 1) + Methode.DeclaringType.ToString().Split('.')[1] + "." + Methode.Name + Message);
                pFichierDebug.Close();
            }
            catch
            {
                StreamWriter pFichierDebug = new StreamWriter(DEBUG_FICHIER, true, System.Text.Encoding.Unicode);
                pFichierDebug.WriteLine("Erreur");
                pFichierDebug.Close();
            }

        }

        #endregion
    }
}
