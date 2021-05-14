using log4net;
using log4net.Appender;
using log4net.Core;
using log4net.Layout;
using log4net.Repository.Hierarchy;
using System;
using System.Diagnostics;

namespace SBO.Hub.Helpers
{
    public class LogHelper
    {
        #region Propriedades Publicas
        /// <summary>
        /// types de Log - Grau de Severidade
        /// </summary>
        public enum LogType
        {
            INFO,
            WARNING,
            ERROR,
            FATAL
        }
        #endregion

        #region Métodos Publicos
        /// <summary>
        /// Escreve uma mensagem no log
        /// </summary>
        /// <param name="mensagem">Mensagem a ser escrita no log</param>
        /// <param name="logType">Grau de severidade da mensagem</param>
        /// <param name="obj">obj que chamou este método: this</param>
        public static void WriteMessageLog(string mensagem, LogType logType, object obj)
        {
            WriteLog(mensagem, logType, obj);
        }

        /// <summary>
        /// Este método escreve a exceção no Log
        /// </summary>
        /// <param name="ex">Exceção a ser escrita no Log</param>
        /// <param name="type">Grau de severidade da exceção</param>
        /// <param name="obj">obj que chamou este método: this</param>
        public static void WriteExceptionLog(Exception ex, LogType type, object obj)
        {
            string mensagem = ex.Message;
            WriteLog(mensagem, type, obj);

            if (ex.InnerException != null)
            {
                WriteExceptionLog(ex.InnerException, type, obj);
            }
        }
        #endregion

        #region Métodos Privados
        /// <summary>
        /// Este método escreve uma linha nova no log, com a mensagem passada
        /// </summary>
        /// <param name="mensagem">mensagem a ser escrita no log</param>
        /// <param name="obj">obj que chamou este método: this</param>
        private static void WriteLog(string mensagem, LogType type, object obj)
        {
            // Pegamos o Log do namespace passado
            ILog log = LogManager.GetLogger(obj.GetType());

            switch (type)
            {
                case LogType.INFO:
                    log.InfoFormat("{0}", mensagem);
                    break;
                case LogType.WARNING:
                    log.WarnFormat("{0}", mensagem);
                    break;
                case LogType.ERROR:
                    log.ErrorFormat("{0}", mensagem);
                    break;
                case LogType.FATAL:
                    log.FatalFormat("{0}", mensagem);
                    break;
            }
        }
        #endregion

        #region Métodos Internos
        /// <summary>
        /// Configura o Log4Net, especificando o diretório para ser gravado os arquivos e o formato
        /// </summary>
        public static void ConfigureLog()
        {
            Hierarchy hierarchy = (Hierarchy)LogManager.GetRepository();

            PatternLayout patternLayout = new PatternLayout();
            patternLayout.ConversionPattern = "%date [%thread] %-5level %logger - %message%newline";
            patternLayout.ActivateOptions();

            RollingFileAppender roller = new RollingFileAppender();
            roller.AppendToFile = true;
            roller.File = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + @"\SAP\SAP Business One\Log\Addon\SBOHub_Log_" + System.Windows.Forms.Application.ProductName + ".log";
            roller.Layout = patternLayout;
            roller.MaxSizeRollBackups = 5;
            roller.MaximumFileSize = "10MB";
            roller.RollingStyle = RollingFileAppender.RollingMode.Size;
            roller.StaticLogFileName = true;
            roller.ActivateOptions();
            hierarchy.Root.AddAppender(roller);

            MemoryAppender memory = new MemoryAppender();
            memory.ActivateOptions();
            hierarchy.Root.AddAppender(memory);

            hierarchy.Root.Level = Level.All;
            hierarchy.Configured = true;


            //Configura o modo de depuração (DEBUG-ONLY)
            Debug.Listeners.Add(new TextWriterTraceListener(Console.Out));
            Debug.AutoFlush = true;
            Debug.Indent();
        }
        #endregion
    }
}
