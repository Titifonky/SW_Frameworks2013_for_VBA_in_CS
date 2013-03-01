using System;
using System.IO;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Diagnostics;
using System.Reflection;

namespace Framework_SW2013
{
    internal static class Debug
    {
        #region "Propriétés"

        private static String DEBUG_FICHIER = Path.Combine(Directory.GetParent(Assembly.GetExecutingAssembly().Location).FullName, Assembly.GetExecutingAssembly().FullName.Split(',')[0] + "_Debug.txt");
        private static Boolean DEBUG_ACTIF = true;

        public static Boolean Actif { get { return DEBUG_ACTIF; } set { DEBUG_ACTIF = value; } }

        #endregion

        #region "Méthodes"

        public static void Init()
        {
            File.Delete(DEBUG_FICHIER);
            StreamWriter pFichierDebug = new StreamWriter(DEBUG_FICHIER, false, System.Text.Encoding.Unicode);
            pFichierDebug.WriteLine("Start Debug");
            pFichierDebug.Close();
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
                pFichierDebug.WriteLine("\t".Repeter(pPile.FrameCount - 1) + "-> " + Message);
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
                String pInfoMethode = Methode.DeclaringType.ToString().Split('.')[1] + "." + Methode.Name;

                StreamWriter pFichierDebug = new StreamWriter(DEBUG_FICHIER, true, System.Text.Encoding.Unicode);
                pFichierDebug.WriteLine("\t".Repeter(pPile.FrameCount - 1) + pInfoMethode);
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
                String pInfoMethode = Methode.DeclaringType.ToString().Split('.')[1] + "." + Methode.Name;

                StreamWriter pFichierDebug = new StreamWriter(DEBUG_FICHIER, true, System.Text.Encoding.Unicode);
                pFichierDebug.WriteLine("\t".Repeter(pPile.FrameCount - 1) + pInfoMethode + "-> " + Message);
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
