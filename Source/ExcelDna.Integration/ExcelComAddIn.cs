//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Excel-DNA is licensed under the zlib license. See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using ExcelDna.ComInterop.ComRegistration;
using ExcelDna.Integration.CustomUI;
using ExcelDna.Integration.Extensibility;
using ExcelDna.Logging;

namespace ExcelDna.Integration
{
    // NOTE: A COM add-in might be able to know phow Excel started by looking in the Custom() values passed to OnConnection
    // See: https://msdn.microsoft.com/en-us/library/aa189748%28office.10%29.aspx
    // Custom()	Variant	An array of Variant type values that provides additional data. 
    // The numeric value of the first element in this array indicates how the host application was started: 
    // from the user interface (1), 
    // by embedding a document created in the host application in another application (2), 
    // or through Automation (3).

    [ComVisible(true)]
#pragma warning disable CS0618 // Type or member is obsolete (but probably not forever)
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
#pragma warning restore CS0618 // Type or member is obsolete
    public class ExcelComAddIn : IDTExtensibility2
    {
        internal DnaLibrary DnaLibrary { get; set; }
        private string _progId;
        private object _addInInst;
        protected string ProgId
        {
            get { return _progId; }
        }

        internal void SetProgId(string progId)
        {
            _progId = progId;
        }

        public string FriendlyName { get; protected set; }
        public string Description { get; protected set; }

        #region IDTExtensibility2 interface
        public virtual void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            _addInInst = AddInInst;
            Logger.ComAddIn.Verbose("ExcelComAddIn.OnConnection");
        }

        public virtual void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            ExcelComAddInHelper.OnUnloadComAddIn(this, _addInInst);
            Logger.ComAddIn.Verbose("ExcelComAddIn.OnDisconnection");
        }

        public virtual void OnAddInsUpdate(ref Array custom)
        {
            Logger.ComAddIn.Verbose("ExcelComAddIn.OnAddInsUpdate");
        }

        public virtual void OnStartupComplete(ref Array custom)
        {
            Logger.ComAddIn.Verbose("ExcelComAddIn.OnStartupComplete");
        }

        public virtual void OnBeginShutdown(ref Array custom)
        {
            Logger.ComAddIn.Verbose("ExcelComAddIn.OnBeginShutdown");
        }
        #endregion
    }

    [ComVisible(true)]
#pragma warning disable CS0618 // Type or member is obsolete (but probably not forever)
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
#pragma warning restore CS0618 // Type or member is obsolete
    internal class DummyComAddIn : IDTExtensibility2
    {
        #region IDTExtensibility2 interface
        public virtual void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            Logger.ComAddIn.Verbose("DummyComAddIn.OnConnection");
        }

        public virtual void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            Logger.ComAddIn.Verbose("DummyComAddIn.OnDisconnection");
        }

        public virtual void OnAddInsUpdate(ref Array custom)
        {
            Logger.ComAddIn.Verbose("DummyComAddIn.OnAddInsUpdate");
        }

        public virtual void OnStartupComplete(ref Array custom)
        {
            Logger.ComAddIn.Verbose("DummyComAddIn.OnStartupComplete");
        }

        public virtual void OnBeginShutdown(ref Array custom)
        {
            Logger.ComAddIn.Verbose("DummyComAddIn.OnBeginShutdown");
        }
        #endregion
    }
    
    public static class ExcelComAddInHelper
    {
        // Com Add-ins loaded for Ribbons.
        static List<object> loadedComAddIns = new List<object>();
        static readonly object _wpsRetryLock = new object();
        static readonly Dictionary<string, int> _wpsRetryCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        const int WpsComAddInRetryMax = 5;
        const int WpsComAddInRetryDelayMs = 1500;

        public static void OnUnloadComAddIn(ExcelComAddIn addIn, object addInInst)
        {
            loadedComAddIns.Remove(addInInst);
        }

        public static void LoadComAddIn(ExcelComAddIn addIn)
        {
            // If we are called without the addIn's DnaLibrary being set, default to the current library
            if (addIn.DnaLibrary == null)
                addIn.DnaLibrary = DnaLibrary.CurrentLibrary;

            Guid clsId;
            string progId;

            // If we have both an explicit Guid and an explicit ProgId, then use those, else create synthetic ProgId.
            object[] progIdAttrs = addIn.GetType().GetCustomAttributes(typeof(ProgIdAttribute), false);
            object[] guidAttrs = addIn.GetType().GetCustomAttributes(typeof(GuidAttribute), false);
            if (progIdAttrs.Length >= 1 && guidAttrs.Length >= 1)
            {
                // Use the attributes
                ProgIdAttribute progIdAtt = (ProgIdAttribute)progIdAttrs[0];
                progId = progIdAtt.Value;

                GuidAttribute guidAtt = (GuidAttribute)guidAttrs[0];
                clsId = new Guid(guidAtt.Value);
            }
            else
            {
                // Use a stable Guid derived from the Xll Path (since Excel stores load-times and other info for every COM add-in loaded in the registry)
                clsId = ExcelDnaUtil.XllGuid;
                // and make the ProgId from this Guid - max 39 chars....
                // Change from Dna.xxx.n to Dna_xxx_n to avoid McAfee bug that blocks registry writes with a "." anywhere
                progId = "Dna_" + clsId.ToString("N") + "_" + loadedComAddIns.Count;
            }
            addIn.SetProgId(progId);

            LoadComAddInInternal(addIn, clsId, progId);
        }

        public static void LoadComAddIn(ExcelComAddIn addIn, Guid clsId, string progId)
        {
            // If we are called without the addIn's DnaLibrary being set, default to the current library
            if (addIn.DnaLibrary == null)
                addIn.DnaLibrary = DnaLibrary.CurrentLibrary;

            if (string.IsNullOrEmpty(progId))
                throw new ArgumentException("progId must be specified.", nameof(progId));

            addIn.SetProgId(progId);

            LoadComAddInInternal(addIn, clsId, progId);
        }

        private static void LoadComAddInInternal(ExcelComAddIn addIn, Guid clsId, string progId)
        {
            // Put together some nicer descriptions for the Add-ins dialog.
            string friendlyName;
            if (addIn.FriendlyName != null)
                friendlyName = addIn.FriendlyName;
            else if (addIn is ExcelRibbon)
                friendlyName = addIn.DnaLibrary.Name; // + " (Ribbon Helper)"; (No more - it is displayed in the Ribbon tooltip!)
            else if (addIn is ExcelCustomTaskPaneAddIn)
                friendlyName = addIn.DnaLibrary.Name + " (Custom Task Pane Helper)";
            else
                friendlyName = addIn.DnaLibrary.Name + " (COM Add-in Helper)";
            string description = addIn.Description ?? string.Format("Dynamically created COM Add-in to load custom UI for the Excel Add-in {0}, located at {1}.", addIn.DnaLibrary.Name, DnaLibrary.XllPath);

            bool isWpsHost = IsWpsHost();
            if (isWpsHost && IsWpsComAddInLoaded(progId))
            {
                Logger.ComAddIn.Verbose("WPS: COM Add-In already loaded. Skipping reload for {0}", progId);
                return;
            }

            Logger.ComAddIn.Verbose("Getting Application object");
            object app = GetApplicationForComAddIn(isWpsHost);
            if (app == null)
            {
                if (isWpsHost)
                {
                    Logger.ComAddIn.Info("Application object unavailable in WPS; scheduling COM Add-In retry.");
                    ScheduleWpsComAddInRetry(addIn, clsId, progId);
                }
                else
                {
                    Logger.ComAddIn.Error("Application object unavailable; skipping Ribbon/COM Add-In load.");
                }
                return;
            }

            Type appType = app.GetType();
            Logger.ComAddIn.Verbose("Got Application object: " + appType.ToString());

            CultureInfo ci = new CultureInfo(1033);
            object excelComAddIns;
            object comAddIn;

            try
            {
                Logger.ComAddIn.Verbose("Loading Ribbon/COM Add-In {0} ({1}) {2} / {3}",
                    addIn.GetType().FullName, friendlyName, progId, clsId);

                using (new ProgIdRegistration(progId, clsId))
                using (new ClsIdRegistration(clsId, progId))
                using (new ComAddInRegistration(progId, friendlyName, description))
                using (new AutomationSecurityOverride(app))
                {
                    // NOTE: A bug was introduced in Excel around version 2310 (16.0.16921.20000) that broke some of the COM add-in load scenarios for ribbons and CTP
                    // This point in the code will (at least under these versions) load the add-in via COM DllGetClassObject
                    // But the add-in is not 'fully' loaded, e.g. the Ribbon is not loaded
                    // However, events on the add-in object will fire, specifically the OnDisconnect event when we do Connect = false later.
                    // So to avoid inconvenience to our 'real' add-in, we give Excel a dummy add-in for now
                    var dummyAddIn = new DummyComAddIn();
                    using (new SingletonClassFactoryRegistration(dummyAddIn, clsId))
                    {
                        excelComAddIns = appType.InvokeMember("COMAddIns", BindingFlags.GetProperty, null, app, null, ci);

                        try
                        {
                            appType.InvokeMember("Update", BindingFlags.InvokeMethod, null, excelComAddIns, null, ci);
                        }
                        catch (Exception ex)
                        {
                            Logger.ComAddIn.Info("COMAddIns.Update failed - continuing. {0}", ex);
                        }

                        try
                        {
                            comAddIn = excelComAddIns.GetType().InvokeMember("Item", BindingFlags.InvokeMethod, null, excelComAddIns, new object[] { progId }, ci);
                        }
                        catch (Exception ex)
                        {
                            Logger.ComAddIn.Info("COMAddIns.Item failed for {0}. {1}", progId, ex);
                            if (!isWpsHost)
                                throw;

                            comAddIn = TryAddComAddIn(excelComAddIns, progId, ci);
                        }

                        // At this point Excel knows how to load our add-in by CLSID, so we could clean up the
                        // registry aggressively, before the actual (dangerous?) loading starts.
                        // But this seems to lead to some distress - Excel has some assertion checked when
                        // it updates the LoadBehavior after a successful load....

                        object connectState = comAddIn.GetType().InvokeMember("Connect", BindingFlags.GetProperty, null, comAddIn, null, ci);
                        TrySetComAddInConnect(comAddIn, false, ci);
                    }
                    // Swap out the dummy add-in for the real one
                    using (new SingletonClassFactoryRegistration(addIn, clsId))
                    {
                        TrySetComAddInConnect(comAddIn, true, ci);
                    }
                    //                            Debug.Print("COMAddin is loaded.");
                    loadedComAddIns.Add(comAddIn);
                    if (isWpsHost)
                    {
                        MarkWpsComAddInLoaded(progId);
                    }

                    Logger.ComAddIn.Verbose("Completed Loading Ribbon/COM Add-In");

                }
            }
            catch (UnauthorizedAccessException secex)
            {
                Logger.ComAddIn.Error(secex, "The Ribbon/COM Add-In helper required by add-in {0} could not be registered.\r\nThis may be due to restricted permissions on the HKCU\\Software\\Classes key", DnaLibrary.CurrentLibrary.Name);
            }
            catch (Exception ex)
            {
                // If Excel is running with the /K switch it seems to indicate we're running 
                // in a COM-unfriendly mode where (sometimes) the COM add-in for the ribbon won't load. 
                // We skip the log display in this case.
                // CONSIDER: How would an add-in know that its COM AddIn load failed in this case?
                if (Environment.CommandLine.Contains(" /K"))
                {
                    Logger.ComAddIn.Info("Load Ribbon/COM Add-In exception with /K in CommandLine \r\n{0}", ex.ToString());
                }
                else
                {
                    Logger.ComAddIn.Error(ex, "The Ribbon/COM add-in helper required by add-in {0} could not be registered.\r\nThis may be due to the helper add-in being disabled by Excel.\r\nTo repair, open Disabled items in the Options->Add-Ins page and re-enable target add-in, then restart Excel.\r\n\r\nError details: {1}", DnaLibrary.CurrentLibrary.Name, ex.ToString());
                }
            }
        }

        internal static void UnloadComAddIns()
        {
            CultureInfo ci = new CultureInfo(1033);
            foreach (object comAddIn in loadedComAddIns.ToArray()) // Disconnecting an add-in removes it from loadedComAddIns, so we operate on a copy of the collection.
            {
                comAddIn.GetType().InvokeMember("Connect", System.Reflection.BindingFlags.SetProperty, null, comAddIn, new object[] { false }, ci);
                Logger.ComAddIn.Info("Ribbon/COM Add-In Unloaded.");
            }
        }

        private static bool IsWpsHost()
        {
            try
            {
                var fileName = Process.GetCurrentProcess().MainModule?.FileName;
                if (string.IsNullOrEmpty(fileName))
                    return false;

                var processName = Path.GetFileNameWithoutExtension(fileName);
                return string.Equals(processName, "et", StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
            }
        }

        private static object GetApplicationForComAddIn(bool isWpsHost)
        {
            object app = null;
            try
            {
                if (!isWpsHost)
                {
                    app = ExcelDnaUtil.GetApplicationNotProtectedNoThrow();
                    if (app != null)
                    {
                        return app;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.ComAddIn.Info("GetApplicationNotProtectedNoThrow failed. {0}", ex);
            }

            const int maxRetries = 10;
            const int delayMs = 100;
            for (int i = 0; i < maxRetries; i++)
            {
                Thread.Sleep(delayMs);
                try
                {
                    if (isWpsHost)
                    {
                        bool isProtected;
                        app = ExcelDnaUtil.GetApplicationNoWorkbook(true, out isProtected);
                    }
                    else
                    {
                        app = ExcelDnaUtil.GetApplicationNotProtectedNoThrow();
                    }
                }
                catch (Exception ex)
                {
                    Logger.ComAddIn.Info("Retry GetApplicationNotProtectedNoThrow failed. Attempt {0}/{1}. {2}", i + 1, maxRetries, ex);
                }

                if (app != null)
                {
                    return app;
                }
            }

            return null;
        }

        private static object TryAddComAddIn(object excelComAddIns, string progId, CultureInfo ci)
        {
            try
            {
                Logger.ComAddIn.Verbose("Attempting COMAddIns.Add({0}, false)", progId);
                return excelComAddIns.GetType().InvokeMember("Add", BindingFlags.InvokeMethod, null, excelComAddIns, new object[] { progId, false }, ci);
            }
            catch (Exception ex)
            {
                Logger.ComAddIn.Info("COMAddIns.Add(progId, false) failed for {0} - trying Add(progId). {1}", progId, ex);
                return excelComAddIns.GetType().InvokeMember("Add", BindingFlags.InvokeMethod, null, excelComAddIns, new object[] { progId }, ci);
            }
        }

        private static void TrySetComAddInConnect(object comAddIn, bool value, CultureInfo ci)
        {
            try
            {
                comAddIn.GetType().InvokeMember("Connect", BindingFlags.SetProperty, null, comAddIn, new object[] { value }, ci);
            }
            catch (Exception ex)
            {
                Logger.ComAddIn.Info("Setting COMAddIn.Connect={0} failed. {1}", value, ex);
                if (value)
                    throw;
            }
        }

        private static bool IsWpsComAddInLoaded(string progId)
        {
            lock (_wpsRetryLock)
            {
                int count;
                return _wpsRetryCounts.TryGetValue(progId, out count) && count < 0;
            }
        }

        private static void MarkWpsComAddInLoaded(string progId)
        {
            lock (_wpsRetryLock)
            {
                _wpsRetryCounts[progId] = -1;
            }
        }

        private static void ScheduleWpsComAddInRetry(ExcelComAddIn addIn, Guid clsId, string progId)
        {
            int attempt;
            lock (_wpsRetryLock)
            {
                int count;
                if (_wpsRetryCounts.TryGetValue(progId, out count))
                {
                    if (count < 0)
                        return;
                    if (count >= WpsComAddInRetryMax)
                    {
                        Logger.ComAddIn.Info("WPS: reached max COM Add-In retry attempts for {0}", progId);
                        return;
                    }
                    count++;
                    _wpsRetryCounts[progId] = count;
                    attempt = count;
                }
                else
                {
                    _wpsRetryCounts[progId] = 1;
                    attempt = 1;
                }
            }

            if (!SynchronizationManager.IsInstalled)
            {
                Logger.ComAddIn.Info("WPS: SynchronizationManager not installed; retry skipped for {0}", progId);
                return;
            }

            Logger.ComAddIn.Info("WPS: scheduling COM Add-In retry {0}/{1} for {2}", attempt, WpsComAddInRetryMax, progId);
            ThreadPool.QueueUserWorkItem(_ =>
            {
                Thread.Sleep(WpsComAddInRetryDelayMs);
                try
                {
                    ExcelAsyncUtil.QueueAsMacro(() => LoadComAddInInternal(addIn, clsId, progId));
                }
                catch (Exception ex)
                {
                    Logger.ComAddIn.Info("WPS: QueueAsMacro retry failed for {0}. {1}", progId, ex);
                }
            });
        }
    }

}
