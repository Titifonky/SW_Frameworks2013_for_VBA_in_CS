using log4net;
using log4net.Appender;
using log4net.Config;
using log4net.Repository;
using System;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;

namespace Framework
{
    [Flags]
    internal enum LogLevelL4N
    {
        DEBUG = 1,
        ERROR = 2,
        FATAL = 4,
        INFO = 8,
        WARN = 16
    }

    internal static class Log
    {
        private static readonly ILog _Logger = LogManager.GetLogger("FrameworksSW");

        private static Boolean _EstInitialise = false;

        private static Boolean _Actif = true;

        static Log()
        {
            String Dossier = Path.GetDirectoryName(Assembly.GetAssembly(typeof(Log)).Location);
            String Chemin = Dossier + @"\" + "log4net.config";
            XmlConfigurator.Configure(_Logger.Logger.Repository, new FileInfo(Chemin));

            IAppender[] appenders = _Logger.Logger.Repository.GetAppenders();
            foreach (IAppender appender in appenders)
            {
                FileAppender fileAppender = appender as FileAppender;

                String CheminFichier = Path.Combine(Dossier, Path.GetFileName(fileAppender.File));
                if (File.Exists(CheminFichier))
                    File.Delete(CheminFichier);

                fileAppender.File = Path.Combine(Dossier, Path.GetFileName(fileAppender.File));
                fileAppender.ActivateOptions();
            }
        }

        internal static void Stopper()
        {
            IAppender[] appenders = _Logger.Logger.Repository.GetAppenders();
            foreach (IAppender appender in appenders)
            {
                appender.Close();
            }
            _Logger.Logger.Repository.Shutdown();
        }

        internal static void Entete()
        {
            if (_EstInitialise)
                return;

            Write("\n ");
            Write("====================================================================================================");
            Write("|                                                                                                  |");
            Write("|                                       SOLIDWORKS DEBUG                                           |");
            Write("|                                                                                                  |");
            Write("====================================================================================================");
            Write("\n ");

            _EstInitialise = true;
        }

        internal static Boolean Activer
        {
            get
            {
                return _Actif;
            }
            set
            {
                _Actif = value;

                log4net.Core.Level pLevel = log4net.Core.Level.Debug;
                if (value)
                    pLevel = log4net.Core.Level.All;

                ILoggerRepository repository = _Logger.Logger.Repository;
                repository.Threshold = pLevel;

                ((log4net.Repository.Hierarchy.Logger)_Logger.Logger).Level = pLevel;

                log4net.Repository.Hierarchy.Hierarchy h = (log4net.Repository.Hierarchy.Hierarchy)repository;
                log4net.Repository.Hierarchy.Logger rootLogger = h.Root;
                rootLogger.Level = pLevel;

            }
        }

        internal static void Write(String Message, LogLevelL4N Level = LogLevelL4N.DEBUG)
        {
            if (Level.Equals(LogLevelL4N.DEBUG))
                _Logger.Debug(Message);
            else if (Level.Equals(LogLevelL4N.ERROR))
                _Logger.Error(Message);
            else if (Level.Equals(LogLevelL4N.FATAL))
                _Logger.Fatal(Message);
            else if (Level.Equals(LogLevelL4N.INFO))
                _Logger.Info(Message);
            else if (Level.Equals(LogLevelL4N.WARN))
                _Logger.Warn(Message);
        }

        internal static void Message(String Message)
        {
            if (!_Actif)
                return;

            Write("\t\t\t\t-> " + Message);
        }

        internal static void Methode(String NomClasse, [CallerMemberName] String Methode = "")
        {
            if (!_Actif)
                return;

            Write("\t\t\t" + NomClasse + "." + Methode);
        }

        internal static void Methode(String NomClasse, String Message, [CallerMemberName] String Methode = "")
        {
            if (!_Actif)
                return;

            Write("\t\t\t" + NomClasse + "." + Methode);
            Write("\t\t\t\t-> " + Message);
        }
    }
}

