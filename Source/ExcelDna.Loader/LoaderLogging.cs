//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Excel-DNA is licensed under the zlib license. See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;

namespace ExcelDna.Loader.Logging
{
    // NOTE: This enum appears here and in TraceLogging in ExcelDna.Integration
    enum IntegrationTraceEventId
    {
        Initialization = 1,
        DnaCompilation = 2,
        Registration = 3,
        ComAddIn = 4,
    }

    class TraceLogger
    {
        static internal TraceSource IntegrationTraceSource; // Set after Integration is initialized
    }

    // NOTE: There is a similar RegistrationLogger class in ExcelDna.Integration.
    // It's easier to maintain two copies for now.
    class Logger
    {
        int _eventId;

        Logger(IntegrationTraceEventId traceEventId)
        {
            _eventId = (int)traceEventId;
        }

        void Log(TraceEventType eventType, string message, params object[] args)
        {
            if (TraceLogger.IntegrationTraceSource == null)
            {
                // We are in the pre-initialization stage. Just log to Debug.
                // NOTE: Without the explcit check and short-circuit here the loading would (sometimes!) fail.
                //       This is despite the catch-all exception handler that should be suppressing all errors.
                //       Somehow the null-reference exception being thrown and caught interferes with the AppDomain_AssemblyResolve!?
                //       The problem only happened when ExcelDna.Integration had to be resolved from resources (in which case the calls here would come from AssemblyResolve.
                //       Under the debugger, the problem was inconsistent (sometimes the load worked fine), but outside the debugger it always failed.
                string formattedMessage = args != null && args.Length > 0
                    ? string.Format(CultureInfo.InvariantCulture, message, args)
                    : message;
                Debug.Print(eventType.ToString() + " : " +
                            string.Format(CultureInfo.InvariantCulture, "{0:yyyy-MM-dd HH:mm:ss.fff} {1}", DateTime.Now, formattedMessage));
                return;
            }
            try
            {
                if (args != null && args.Length > 0)
                {
                    message = string.Format(CultureInfo.InvariantCulture, message, args);
                }
                string timestampedMessage = string.Format(CultureInfo.InvariantCulture,
                    "{0:yyyy-MM-dd HH:mm:ss.fff} {1}", DateTime.Now, message);
                TraceLogger.IntegrationTraceSource.TraceEvent(eventType, _eventId, timestampedMessage);
            }
            catch (Exception e)
            {
                // We certainly want to suppress errors here, though they indicate Excel-DNA bugs.
                Debug.Print("ExcelDna.Loader - Logger.Log error: " + e.Message);
            }
        }

        public void Verbose(string message, params object[] args)
        {
            Log(TraceEventType.Verbose, message, args);
        }

        public void Info(string message, params object[] args)
        {
            Log(TraceEventType.Information, message, args);
        }

        public void Warn(string message, params object[] args)
        {
            Log(TraceEventType.Warning, message, args);
        }

        public void Error(string message, params object[] args)
        {
            Log(TraceEventType.Error, message, args);
        }

        public void Error(Exception ex, string message, params object[] args)
        {
            if (args != null)
            {
                try
                {
                    message = string.Format(CultureInfo.InvariantCulture, message, args);
                }
                catch (Exception fex)
                {
                    Debug.Print("Logger.Error formatting exception " + fex.Message);
                }
            }
            Log(TraceEventType.Error, "{0} : {1} - {2}", message, ex.GetType().Name, ex.Message);
        }

        static Logger _initializationLogger = new Logger(IntegrationTraceEventId.Initialization);
        static internal Logger Initialization { get { return _initializationLogger; } }
        static Logger _registrationLogger = new Logger(IntegrationTraceEventId.Registration);
        static internal Logger Registration { get { return _registrationLogger; } }
        static Logger _dnaCompilationLogger = new Logger(IntegrationTraceEventId.DnaCompilation);
        static internal Logger DnaCompilation { get { return _dnaCompilationLogger; } }
        static Logger _comAddInLogger = new Logger(IntegrationTraceEventId.ComAddIn);
        static internal Logger ComAddIn { get { return _comAddInLogger; } }

    }
}
