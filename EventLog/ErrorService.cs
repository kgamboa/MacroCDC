using System;
using System.Collections.Generic;

using System.Linq;
using System.Text;
using System.Diagnostics;
using System.IO;
using System.Configuration;

public enum LogMsgType
{
    Debug,
    Information,
    Alert,
    Error,
    Fatal
}
public enum ErrorType
{
    ERROR_AdminPAQ,
    ERROR_Facturacion,
    ERROR_Captura,
    HORARIO_PRODUCTIVO,
    ERROR_Licencia,
    ERROR_Poliza,
    ERROR_Corte,
    ERROR_REP,
    ERROR_EnvioCorreos,
    ERROR_Actualizacion,
    ERROR_EstadoCuenta,
    ERROR_Sabana,
    ERROR_Duplicados,
    ERROR_Bancos,
    ERROR_MultiOtorgante,
    ERROR_SolicitudWEB,
    ERROR_ComercialSDK,
    ERROR
}

public class ErrorServices
{


    #region IErrorService Members


    private static ErrorServices _singleton = null;
    public static ErrorServices singleton
    {
        get
        {
            if (_singleton == null)
            {
                _singleton = new ErrorServices();
            }
            return _singleton;
        }
    }

    private string _path = null;
    /// <summary>
    /// Prefijo del log con ruta.
    /// </summary>
    public string Path
    {
        get
        {
            return _path;
        }
        set
        {
            _path = value;

            if (!Directory.Exists(_path))
                Directory.CreateDirectory(_path);
        }
    }

    /// <summary>
    /// Nombre del archivo actual.
    /// </summary>
    private string _LogPath;
    public string LogPath
    {
        get
        {
            if (Path == null)
                return string.Empty;
            var currentFile = Path + "\\" + Proceso + "_";
            if (_LogPath == null || !_LogPath.Contains(currentFile))
                _LogPath = currentFile + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".log";
            return _LogPath;
        }
    }

    private string _errorLogPath;
    public string ErrorLogPath
    {
        get
        {
            if (_errorLogPath == null)
                if (Proceso != null || Proceso != string.Empty)
                    _errorLogPath = Path + "\\Error_" + Proceso + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".log";
                else
                    _errorLogPath = Path + "\\Error_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".log";
            return _errorLogPath;
        }
    }


    private string _pedidoLogPath;
    public string PedidoLogPath
    {
        get
        {
            if (_pedidoLogPath == null)
                if (Proceso != null || Proceso != string.Empty)
                    _pedidoLogPath = Path + "\\Pedido_" + Proceso + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".log";
                else
                    _pedidoLogPath = Path + "\\Pedido" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".log";
            return _pedidoLogPath;
        }
    }
    private string _defaultEventSource = "Facturacion";
    public string defaultEventSource
    {
        get
        {
            return _defaultEventSource;
        }
        set
        {
            _defaultEventSource = value;
        }
    }

    private bool failToCreateDefaultLog = false;
    private EventLog _eventLog = null;
    public EventLog eventLog
    {
        get
        {
            if (_eventLog == null)
                createDefaultEventLog();
            return _eventLog;
        }
        set
        {
            _eventLog = value;
        }
    }

    private int _logLevel;
    public int logLevel
    {
        get
        {
            return _logLevel;
        }
        set
        {
            _logLevel = value;
        }
    }

    private void createDefaultEventLog()
    {
        lock (_defaultEventSource)
        {

            if (failToCreateDefaultLog)
                return;
            try
            {

                ///Crear el origen de eventos si no existe 
                if (!EventLog.SourceExists(defaultEventSource))
                {
                    EventSourceCreationData data = new EventSourceCreationData(defaultEventSource, "Facturacion");
                    EventLog.CreateEventSource(data);
                }

                ///Crear el objeto para EventLog
                _eventLog = new EventLog("Facturacion", defaultEventSource);
            }
            catch (Exception)
            {
                failToCreateDefaultLog = true;
            }
        }
    }

    public void LogError(String message, ErrorType errorType, LogMsgType logMsgType)
    {
        /// Si el service log es menor al indicado, no procesa nada.
        if (!compareLevel(logMsgType, this.logLevel))
            return;
        if (Path != null)
            LogMessageToHd(message, 1);
        if (eventLog != null)
            LogMessageToEventLog(message, logMsgType);
        //BiometriaLogger.WriteLog(ELogLevel.DEBUG, message);
    }
    internal void LogMessage(String message)
    {
        LogMessageToHd(message, 0);
    }
    internal string LogMessage(params object[] param)
    {
        StringBuilder sbTemp = new StringBuilder();
        if (param.Length > 0)
            sbTemp.Append(string.Join(" ", param));

        LogMessageToHd(sbTemp.ToString(), 0);
        return sbTemp.ToString();
    }
    /// <summary>
    /// Regresa true si log mnesaje > log configurado
    /// </summary>
    /// <param name="logMsgType">log del mensaje</param>
    /// <param name="logLevel">log configurado</param>
    /// <returns></returns>
    private bool compareLevel(LogMsgType logMsgType, int logLevel)
    {
        switch (logMsgType)
        {
            case LogMsgType.Debug:
                if (logLevel <= 1) return true;
                break;
            case LogMsgType.Information:
                if (logLevel <= 2) return true;
                break;
            case LogMsgType.Alert:
                if (logLevel <= 3) return true;
                break;
            case LogMsgType.Error:
                if (logLevel <= 4) return true;
                break;
            case LogMsgType.Fatal:
                if (logLevel <= 5) return true;
                break;
        }
        return false;
    }

    private void LogMessageToEventLog(String message, LogMsgType logMsgType)
    {
        EventLogEntryType eventLogType;
        switch (logMsgType)
        {
            case LogMsgType.Debug:
                eventLogType = EventLogEntryType.Information;
                break;
            case LogMsgType.Information:
                eventLogType = EventLogEntryType.Information;
                break;
            case LogMsgType.Alert:
                eventLogType = EventLogEntryType.Warning;
                break;
            case LogMsgType.Fatal:
            case LogMsgType.Error:
                eventLogType = EventLogEntryType.Error;
                break;
            default:
                eventLogType = EventLogEntryType.Information;
                break;
        }
        try
        {
            eventLog.WriteEntry(message, eventLogType);
        }
        catch { }
    }

    private void LogMessageToHd(String message, int type)
    {
        String writep = Environment.NewLine + DateTime.Now.ToShortTimeString() + " -> ";
        writep += message;// + Environment.NewLine;

        try
        {

            if (type == 1)
                File.AppendAllText(ErrorLogPath, writep, Encoding.Unicode);
            else if (type == 0)
                File.AppendAllText(LogPath, writep, Encoding.Unicode);
            else if (type == 2)
                File.AppendAllText(PedidoLogPath, writep, Encoding.Unicode);
        }
        catch { }
    }




    #endregion

    internal string LogError(ErrorType errorType, Exception ex, params object[] param)
    {
        StringBuilder sbTemp = new StringBuilder(errorType.ToString());
        if (ex != null)
        {
            string Mensaje = ex.Message;
            if (ex.InnerException != null)
            {
                if (ex.InnerException.Message != null)
                    Mensaje += ex.InnerException.Message;

                if (ex.InnerException.StackTrace != null)
                    Mensaje += ex.InnerException.StackTrace;
            }
            if (ex.StackTrace != null)
                Mensaje += ex.StackTrace;

            sbTemp.Append(":" + Mensaje);
        }
        if (param.Length > 0)
            sbTemp.Append(" " + string.Join(" - ", param));

        LogError(sbTemp.ToString(), ErrorType.ERROR_Facturacion, LogMsgType.Fatal);
        return sbTemp.ToString();
    }

    private string proceso;
    public string Proceso
    {
        get { return proceso; }
        set
        {
            proceso = value;
            _LogPath = null;
        }
    }
}
