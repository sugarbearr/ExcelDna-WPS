//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Excel-DNA is licensed under the zlib license. See LICENSE.txt for details.

using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using ExcelDna.Logging;

namespace ExcelDna.Integration.Rtd
{
    /// <summary>
    /// RTD降级配置和状态管理
    /// 当RTD功能不可用时（如低版本WPS），提供降级方案
    /// </summary>
    public static class RtdFallbackManager
    {
        #region Configuration

        /// <summary>
        /// 是否启用RTD降级功能
        /// 默认为true，允许在RTD失败时自动降级
        /// </summary>
        public static bool EnableFallback { get; set; } = true;

        /// <summary>
        /// 是否强制使用降级模式（跳过RTD尝试）
        /// 用于已知不支持RTD的环境
        /// </summary>
        public static bool ForceFallbackMode { get; set; } = false;

        /// <summary>
        /// RTD失败后的重试间隔（毫秒）
        /// 设为0表示永不重试RTD
        /// </summary>
        public static int RtdRetryIntervalMs { get; set; } = 0; // 默认不重试

        /// <summary>
        /// 降级模式下自动刷新的间隔（毫秒）
        /// 用于持续更新的Observable数据流
        /// 默认为0，表示不自动刷新（避免无限循环）
        /// </summary>
        public static int FallbackRefreshIntervalMs { get; set; } = 0; // 默认禁用自动刷新

        /// <summary>
        /// 缓存条目的过期时间（毫秒）
        /// 超过此时间未访问的缓存条目将被清理
        /// 设为0表示永不过期
        /// </summary>
        public static int CacheExpirationMs { get; set; } = 300000; // 默认5分钟

        /// <summary>
        /// 缓存清理间隔（毫秒）
        /// </summary>
        public static int CacheCleanupIntervalMs { get; set; } = 60000; // 默认1分钟

        #endregion

        #region State

        private static bool _rtdFailed = false;
        private static DateTime _rtdFailedTime = DateTime.MinValue;
        private static readonly object _stateLock = new object();
        private static bool _rtdTestCompleted = false;
        private static Timer _cacheCleanupTimer = null;
        private static Timer _volatileRefreshTimer = null;
        private static readonly object _timerLock = new object();
        private static readonly ConcurrentDictionary<AsyncCallInfo, FallbackCacheEntry> _refreshPending
            = new ConcurrentDictionary<AsyncCallInfo, FallbackCacheEntry>();

        /// <summary>
        /// RTD是否已经失败（用于快速跳过后续RTD尝试）
        /// </summary>
        public static bool RtdFailed
        {
            get
            {
                lock (_stateLock)
                {
                    if (!_rtdFailed)
                        return false;

                    // 检查是否应该重试RTD
                    if (RtdRetryIntervalMs > 0 &&
                        (DateTime.UtcNow - _rtdFailedTime).TotalMilliseconds > RtdRetryIntervalMs)
                    {
                        Logger.RtdServer.Info("RtdFallbackManager - Retry interval reached, will attempt RTD again");
                        _rtdFailed = false;
                        return false;
                    }

                    return true;
                }
            }
        }

        /// <summary>
        /// 标记RTD已失败，后续调用将使用降级方案
        /// </summary>
        internal static void MarkRtdFailed(string reason)
        {
            lock (_stateLock)
            {
                if (!_rtdFailed)
                {
                    _rtdFailed = true;
                    _rtdFailedTime = DateTime.UtcNow;
                    Logger.RtdServer.Warn("RtdFallbackManager - RTD marked as failed, fallback mode enabled. Reason: {0}", reason);

                    // 注意：暂时禁用定时器以避免无限循环问题
                    // 用户可以通过显式调用 StartTimers() 来启用
                    // StartTimers();
                }
            }
        }

        /// <summary>
        /// 重置RTD失败状态（用于测试或手动重试）
        /// </summary>
        public static void ResetRtdFailedState()
        {
            lock (_stateLock)
            {
                _rtdFailed = false;
                _rtdFailedTime = DateTime.MinValue;
                _rtdTestCompleted = false;
                Logger.RtdServer.Info("RtdFallbackManager - RTD failed state has been reset");
            }
        }

        /// <summary>
        /// 检查是否应该使用降级模式
        /// </summary>
        internal static bool ShouldUseFallback
        {
            get
            {
                if (ForceFallbackMode)
                    return true;

                if (!EnableFallback)
                    return false;

                return RtdFailed;
            }
        }

        #endregion

        #region Smart WPS Detection & RTD Testing

        /// <summary>
        /// 在加载时主动测试RTD支持
        /// 建议在AutoOpen中调用此方法
        /// </summary>
        /// <returns>true表示RTD可用，false表示需要使用降级模式</returns>
        public static bool TestRtdSupport()
        {
            lock (_stateLock)
            {
                if (_rtdTestCompleted)
                {
                    return !_rtdFailed;
                }

                try
                {
                    // 检测是否是WPS环境
                    if (ExcelDnaUtil.IsET)
                    {
                        Logger.RtdServer.Info("RtdFallbackManager - WPS environment detected, testing RTD support...");

                        // 获取WPS版本进行判断
                        var wpsVersion = GetWpsVersion();
                        if (wpsVersion != null)
                        {
                            Logger.RtdServer.Info("RtdFallbackManager - WPS version: {0}", wpsVersion);

                            // 根据已知的WPS版本兼容性进行判断
                            // WPS 2019及以后版本通常支持RTD较好
                            // 这里的版本号可以根据实际测试结果调整
                            if (wpsVersion.Major < 11)
                            {
                                Logger.RtdServer.Warn("RtdFallbackManager - WPS version {0} may not support RTD properly", wpsVersion);
                                MarkRtdFailed($"WPS version {wpsVersion} detected as potentially incompatible");
                                _rtdTestCompleted = true;
                                return false;
                            }
                        }

                        // 尝试实际的RTD测试调用
                        if (!TryTestRtdCall())
                        {
                            MarkRtdFailed("RTD test call failed in WPS environment");
                            _rtdTestCompleted = true;
                            return false;
                        }
                    }

                    _rtdTestCompleted = true;
                    Logger.RtdServer.Info("RtdFallbackManager - RTD support test passed");
                    return true;
                }
                catch (Exception ex)
                {
                    Logger.RtdServer.Error(ex, "RtdFallbackManager - RTD support test failed with exception");
                    MarkRtdFailed($"RTD test exception: {ex.Message}");
                    _rtdTestCompleted = true;
                    return false;
                }
            }
        }

        /// <summary>
        /// 获取WPS版本号
        /// </summary>
        private static Version GetWpsVersion()
        {
            try
            {
                var excelVersion = ExcelDnaUtil.ExcelVersion;
                // WPS的版本号格式可能不同，这里尝试解析
                return new Version((int)excelVersion, 0);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// 尝试进行一次简单的RTD测试调用
        /// </summary>
        private static bool TryTestRtdCall()
        {
            try
            {
                // 尝试检查RTD相关的Excel功能是否可用
                // 这是一个轻量级测试，不会实际创建RTD服务器
                var result = XlCall.TryExcel(XlCall.xlfGetWorkspace, out object workspace, 2);
                if (result != XlCall.XlReturn.XlReturnSuccess)
                {
                    return false;
                }

                // 如果基本的Excel API调用成功，我们暂时认为RTD可能可用
                // 实际的RTD测试会在第一次使用时进行
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 检测并配置WPS兼容模式
        /// 建议在AutoOpen中的SynchronizationManager安装后调用
        /// </summary>
        public static void ConfigureForWpsIfNeeded()
        {
            if (ExcelDnaUtil.IsET)
            {
                Logger.RtdServer.Info("RtdFallbackManager - Configuring for WPS environment");

                // 主动测试RTD支持
                TestRtdSupport();
            }
        }

        #endregion

        #region Fallback Cache with Expiration

        // 缓存结构：key = (functionName, parameters的哈希), value = (result, isComplete, timestamp)
        private static readonly ConcurrentDictionary<AsyncCallInfo, FallbackCacheEntry> _fallbackCache
            = new ConcurrentDictionary<AsyncCallInfo, FallbackCacheEntry>();

        // 跟踪持续更新的Observable（未完成的）
        private static readonly ConcurrentDictionary<AsyncCallInfo, FallbackCacheEntry> _activeObservables
            = new ConcurrentDictionary<AsyncCallInfo, FallbackCacheEntry>();

        internal class FallbackCacheEntry
        {
            public object Value { get; set; }
            public bool IsComplete { get; set; }
            public bool HasError { get; set; }
            public DateTime CreatedTime { get; set; }
            public DateTime LastAccessTime { get; set; }
            public DateTime LastUpdateTime { get; set; }
            public IDisposable Subscription { get; set; }
            public int UpdateCount { get; set; }
            public DateTime LastForceRefreshTime { get; set; }

            // 跟踪所有调用此函数的单元格引用
            private readonly object _callersLock = new object();
            private readonly HashSet<string> _callerKeys = new HashSet<string>();
            public List<ExcelReference> Callers { get; } = new List<ExcelReference>();

            public FallbackCacheEntry()
            {
                Value = ExcelError.ExcelErrorNA; // 初始值为 #N/A (计算中)
                IsComplete = false;
                HasError = false;
                CreatedTime = DateTime.UtcNow;
                LastAccessTime = DateTime.UtcNow;
                LastUpdateTime = DateTime.UtcNow;
                UpdateCount = 0;
                LastForceRefreshTime = DateTime.MinValue;
            }

            /// <summary>
            /// 检查缓存条目是否过期
            /// </summary>
            public bool IsExpired(int expirationMs)
            {
                if (expirationMs <= 0)
                    return false;

                // 未完成的Observable不过期
                if (!IsComplete)
                    return false;

                return (DateTime.UtcNow - LastAccessTime).TotalMilliseconds > expirationMs;
            }

            /// <summary>
            /// 更新最后访问时间
            /// </summary>
            public void Touch()
            {
                LastAccessTime = DateTime.UtcNow;
            }

            /// <summary>
            /// 添加调用者单元格引用（去重）
            /// </summary>
            public void AddCaller(ExcelReference caller)
            {
                if (caller == null) return;

                lock (_callersLock)
                {
                    // 使用字符串key去重
                    string key = GetCallerKey(caller);
                    if (!_callerKeys.Contains(key))
                    {
                        _callerKeys.Add(key);
                        Callers.Add(caller);
                    }
                }
            }

            /// <summary>
            /// 获取所有调用者的副本
            /// </summary>
            public List<ExcelReference> GetCallers()
            {
                lock (_callersLock)
                {
                    return new List<ExcelReference>(Callers);
                }
            }

            private static string GetCallerKey(ExcelReference caller)
            {
                return $"{caller.SheetId}:{caller.RowFirst},{caller.ColumnFirst}:{caller.RowLast},{caller.ColumnLast}";
            }
        }

        /// <summary>
        /// 使用降级模式处理Observable
        /// </summary>
        internal static object ProcessObservableFallback(
            string functionName,
            object parameters,
            ExcelObservableOptions options,
            ExcelObservableSource getObservable)
        {
            TryMarkVolatile();
            var callInfo = new AsyncCallInfo(functionName, parameters);

            // 获取当前调用者单元格引用
            ExcelReference caller = null;
            try
            {
                caller = XlCall.Excel(XlCall.xlfCaller) as ExcelReference;
            }
            catch
            {
                // 忽略获取调用者失败的情况
            }

            // 检查缓存
            if (_fallbackCache.TryGetValue(callInfo, out var entry))
            {
                entry.Touch(); // 更新访问时间
                entry.AddCaller(caller); // 记录调用者

                if (entry.IsComplete)
                {
                    Debug.Print("RtdFallback - Returning cached result for: {0}", functionName);
                    return entry.Value;
                }
                else
                {
                    Debug.Print("RtdFallback - Calculation in progress for: {0}", functionName);
                    // 对于持续更新的Observable，返回当前值（即使是#N/A）
                    return entry.Value;
                }
            }

            // 创建新的缓存条目并启动后台计算
            entry = new FallbackCacheEntry();
            entry.AddCaller(caller); // 记录调用者

            if (!_fallbackCache.TryAdd(callInfo, entry))
            {
                // 其他线程已经添加了，使用那个
                if (_fallbackCache.TryGetValue(callInfo, out entry))
                {
                    entry.Touch();
                    entry.AddCaller(caller);
                    return entry.IsComplete ? entry.Value : (object)ExcelError.ExcelErrorNA;
                }
                return ExcelError.ExcelErrorNA;
            }

            // 启动Observable订阅
            try
            {
                IExcelObservable observable = getObservable();
                var observer = new FallbackObserver(callInfo, entry);
                entry.Subscription = observable.Subscribe(observer);

                // 跟踪活跃的Observable
                _activeObservables[callInfo] = entry;

                Debug.Print("RtdFallback - Started calculation for: {0}", functionName);

                // 确保定时器已启动
                StartTimers();
            }
            catch (Exception ex)
            {
                entry.Value = ExcelIntegration.HandleUnhandledException(ex);
                entry.IsComplete = true;
                entry.HasError = true;
                Logger.RtdServer.Error(ex, "RtdFallback - Error starting observable for: {0}", functionName);
            }

            // 首次调用返回 #N/A 表示"计算中"
            return ExcelError.ExcelErrorNA;
        }

        /// <summary>
        /// 降级模式的Observer，将结果写入缓存
        /// </summary>
        private class FallbackObserver : IExcelObserver
        {
            private readonly AsyncCallInfo _callInfo;
            private readonly FallbackCacheEntry _entry;

            public FallbackObserver(AsyncCallInfo callInfo, FallbackCacheEntry entry)
            {
                _callInfo = callInfo;
                _entry = entry;
            }

            public void OnNext(object value)
            {
                _entry.Value = value;
                _entry.LastUpdateTime = DateTime.UtcNow;
                _entry.UpdateCount++;
                MarkRefreshPending(_callInfo, _entry);
                if (Interlocked.Exchange(ref _fallbackOnNextLogged, 1) == 0)
                {
                    Logger.RtdServer.Info("RtdFallback - OnNext received. ValueType={0}, Value={1}",
                        value == null ? "null" : value.GetType().FullName,
                        DescribeValue(value));
                }
                Debug.Print("RtdFallback - OnNext for: hash={0}, value={1}, updateCount={2}",
                    _callInfo.GetHashCode(), value, _entry.UpdateCount);

                // 触发Excel重算以获取新值
                ScheduleRecalculation();
            }

            public void OnCompleted()
            {
                _entry.IsComplete = true;
                MarkRefreshPending(_callInfo, _entry);
                if (Interlocked.Exchange(ref _fallbackOnCompletedLogged, 1) == 0)
                {
                    Logger.RtdServer.Info("RtdFallback - OnCompleted received for hash={0}", _callInfo.GetHashCode());
                }
                Debug.Print("RtdFallback - OnCompleted for: hash={0}", _callInfo.GetHashCode());

                // 清理订阅
                _entry.Subscription?.Dispose();
                _entry.Subscription = null;

                // 从活跃Observable列表中移除
                _activeObservables.TryRemove(_callInfo, out _);

                // 触发最终重算
                ScheduleRecalculation();
            }

            public void OnError(Exception exception)
            {
                _entry.Value = ExcelIntegration.HandleUnhandledException(exception);
                _entry.IsComplete = true;
                _entry.HasError = true;
                MarkRefreshPending(_callInfo, _entry);
                if (Interlocked.Exchange(ref _fallbackOnErrorLogged, 1) == 0)
                {
                    Logger.RtdServer.Warn("RtdFallback - OnError received. ExceptionType={0}, Message={1}",
                        exception == null ? "null" : exception.GetType().FullName,
                        exception?.Message);
                }
                Logger.RtdServer.Error(exception, "RtdFallback - OnError for: hash={0}", _callInfo.GetHashCode());

                // 清理订阅
                _entry.Subscription?.Dispose();
                _entry.Subscription = null;

                // 从活跃Observable列表中移除
                _activeObservables.TryRemove(_callInfo, out _);

                // 触发重算以显示错误
                ScheduleRecalculation();
            }

            private void ScheduleRecalculation()
            {
                // 使用全局节流的重算请求
                RequestBatchRecalculation();
            }
        }

        #endregion



        #region Timers & Throttled Recalculation

        // 重算节流控制
        private static bool _isRecalcPending = false;
        private static DateTime _lastRecalcTime = DateTime.MinValue;
        private static readonly object _recalcLock = new object();
        private static Timer _throttleTimer = null;
        private const int MIN_RECALC_INTERVAL_MS = 250; // 最小重算间隔
        private static int _recalcFailureLogged = 0;
        private static int _queueAsMacroFailureLogged = 0;
        private static int _volatileMarkFailureLogged = 0;
        private static int _fullCalcLogged = 0;
        private static int _recalcOutcomeLogged = 0;
        private static int _fallbackOnNextLogged = 0;
        private static int _fallbackOnCompletedLogged = 0;
        private static int _fallbackOnErrorLogged = 0;
        private static int _forceRefreshLogged = 0;
        private const int MIN_FORCE_REFRESH_INTERVAL_MS = 500;
        private const int MAX_FORCE_REFRESH_CALLERS = 100;

        /// <summary>
        /// 请求批量重算（带节流）
        /// </summary>
        internal static void RequestBatchRecalculation()
        {
            lock (_recalcLock)
            {
                if (_isRecalcPending) return; // 已经有计划的重算了

                var now = DateTime.UtcNow;
                var timeSinceLast = (now - _lastRecalcTime).TotalMilliseconds;

                if (timeSinceLast >= MIN_RECALC_INTERVAL_MS)
                {
                    // 距离上次重算已经足够久，直接调度
                    _isRecalcPending = true;
                    ScheduleMacro();
                }
                else
                {
                    // 需要延迟执行
                    _isRecalcPending = true;
                    if (_throttleTimer == null)
                    {
                        var delay = (int)(MIN_RECALC_INTERVAL_MS - timeSinceLast);
                        _throttleTimer = new Timer(_ =>
                        {
                            ScheduleMacro();
                            _throttleTimer?.Dispose();
                            _throttleTimer = null;
                        }, null, delay, Timeout.Infinite);
                    }
                }
            }
        }

        private static void ScheduleMacro()
        {
            try
            {
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    try
                    {
                        // 再次检查是否真的需要（理论上QueueAsMacro保证在主线程，但为了安全）
                        lock (_recalcLock)
                        {
                            _isRecalcPending = false;
                            _lastRecalcTime = DateTime.UtcNow;
                        }

                        // 执行重算 (只计算当前工作表)
                        TryRecalculate();
                    }
                    catch (Exception ex)
                    {
                        Debug.Print("RtdFallback - Batch recalculation error: {0}", ex.Message);
                    }
                });
            }
            catch (Exception ex)
            {
                Debug.Print("RtdFallback - QueueAsMacro error: {0}", ex.Message);
                if (Interlocked.Exchange(ref _queueAsMacroFailureLogged, 1) == 0)
                {
                    Logger.RtdServer.Warn("RtdFallback - QueueAsMacro failed; async recalculation may be unavailable. Error: {0}", ex.Message);
                }
                lock (_recalcLock) { _isRecalcPending = false; }
            }
        }

        private static void TryRecalculate()
        {
            try
            {
                object ignored;
                XlCall.XlReturn ret = XlCall.XlReturn.XlReturnFailed;
                string successMethod = null;
                bool anySuccess = false;

                if (ExcelDnaUtil.IsET)
                {
                    ret = XlCall.TryExcel(XlCall.xlcCalculateNow, out ignored);
                    if (ret == XlCall.XlReturn.XlReturnSuccess)
                    {
                        anySuccess = true;
                        if (successMethod == null)
                            successMethod = "xlcCalculateNow";
                    }

                    var retDoc = XlCall.TryExcel(XlCall.xlcCalculateDocument, out ignored);
                    if (retDoc == XlCall.XlReturn.XlReturnSuccess)
                    {
                        anySuccess = true;
                        if (successMethod == null)
                            successMethod = "xlcCalculateDocument";
                    }
                    if (retDoc != XlCall.XlReturn.XlReturnSuccess)
                    {
                        ret = retDoc;
                    }

                    // WPS may report success without updating - also try COM Calculate for reliability.
                    if (TryComCalculate(out var comMethod))
                    {
                        anySuccess = true;
                        if (successMethod == null)
                            successMethod = comMethod;
                    }

                    // Some WPS builds only respond to a full calculation.
                    if (Interlocked.Exchange(ref _fullCalcLogged, 1) == 0)
                    {
                        if (TryComCalculateFull(out var fullMethod))
                        {
                            anySuccess = true;
                            if (successMethod == null)
                                successMethod = fullMethod;
                        }
                    }
                    else if (!anySuccess)
                    {
                        if (TryComCalculateFull(out var fullMethod))
                        {
                            anySuccess = true;
                            if (successMethod == null)
                                successMethod = fullMethod;
                        }
                    }

                    if (anySuccess)
                    {
                        LogRecalcOutcomeOnce(true, successMethod, ret);
                        ProcessPendingRefresh();
                        return;
                    }
                }
                else
                {
                    var retDoc = XlCall.TryExcel(XlCall.xlcCalculateDocument, out ignored);
                    if (retDoc == XlCall.XlReturn.XlReturnSuccess)
                    {
                        LogRecalcOutcomeOnce(true, "xlcCalculateDocument", retDoc);
                        ProcessPendingRefresh();
                        return;
                    }

                    ret = XlCall.TryExcel(XlCall.xlcCalculateNow, out ignored);
                    if (ret == XlCall.XlReturn.XlReturnSuccess)
                    {
                        LogRecalcOutcomeOnce(true, "xlcCalculateNow", ret);
                        ProcessPendingRefresh();
                        return;
                    }
                }

                LogRecalcOutcomeOnce(false, null, ret);
                if (Interlocked.Exchange(ref _recalcFailureLogged, 1) == 0)
                {
                    Logger.RtdServer.Warn("RtdFallback - Recalculation failed. Last return: {0}", ret);
                }
            }
            catch (Exception ex)
            {
                if (Interlocked.Exchange(ref _recalcFailureLogged, 1) == 0)
                {
                    Logger.RtdServer.Warn("RtdFallback - Recalculation threw exception: {0}", ex.Message);
                }
            }
        }

        private static bool TryComCalculate(out string method)
        {
            method = null;
            try
            {
                object app = ExcelDnaUtil.Application;
                if (app == null)
                    return false;

                var culture = CultureInfo.GetCultureInfo(1033);
                if (TryInvokeCalculate(app, "Calculate", culture))
                {
                    method = "COM Calculate";
                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                if (Interlocked.Exchange(ref _recalcFailureLogged, 1) == 0)
                {
                    Logger.RtdServer.Warn("RtdFallback - COM Calculate failed: {0}", ex.Message);
                }
                return false;
            }
        }

        private static bool TryComCalculateFull(out string method)
        {
            method = null;
            try
            {
                object app = ExcelDnaUtil.Application;
                if (app == null)
                    return false;

                var culture = CultureInfo.GetCultureInfo(1033);
                if (TryInvokeCalculate(app, "CalculateFull", culture))
                {
                    method = "COM CalculateFull";
                    return true;
                }

                if (TryInvokeCalculate(app, "CalculateFullRebuild", culture))
                {
                    method = "COM CalculateFullRebuild";
                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                if (Interlocked.Exchange(ref _recalcFailureLogged, 1) == 0)
                {
                    Logger.RtdServer.Warn("RtdFallback - COM CalculateFull failed: {0}", ex.Message);
                }
                return false;
            }
        }

        private static bool TryInvokeCalculate(object app, string methodName, CultureInfo culture)
        {
            try
            {
                app.GetType().InvokeMember(methodName, BindingFlags.InvokeMethod, null, app, null, culture);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static void TryMarkVolatile()
        {
            try
            {
                var ret = XlCall.TryExcel(XlCall.xlfVolatile, out _, true);
                if (ret != XlCall.XlReturn.XlReturnSuccess)
                {
                    if (Interlocked.Exchange(ref _volatileMarkFailureLogged, 1) == 0)
                    {
                        Logger.RtdServer.Warn("RtdFallback - xlfVolatile returned {0}", ret);
                    }
                }
            }
            catch (Exception ex)
            {
                if (Interlocked.Exchange(ref _volatileMarkFailureLogged, 1) == 0)
                {
                    Logger.RtdServer.Warn("RtdFallback - xlfVolatile failed: {0}", ex.Message);
                }
            }
        }

        private static void MarkRefreshPending(AsyncCallInfo callInfo, FallbackCacheEntry entry)
        {
            if (!ExcelDnaUtil.IsET)
                return;

            if (entry == null)
                return;

            _refreshPending[callInfo] = entry;
        }

        private static void ProcessPendingRefresh()
        {
            if (!ExcelDnaUtil.IsET)
                return;

            if (_refreshPending.IsEmpty)
                return;

            int total = 0;
            int success = 0;
            int failed = 0;
            DateTime now = DateTime.UtcNow;

            foreach (var kvp in _refreshPending.ToArray())
            {
                if (!_refreshPending.TryRemove(kvp.Key, out var entry) || entry == null)
                    continue;

                if ((now - entry.LastForceRefreshTime).TotalMilliseconds < MIN_FORCE_REFRESH_INTERVAL_MS)
                    continue;

                entry.LastForceRefreshTime = now;

                var callers = entry.GetCallers();
                foreach (var caller in callers)
                {
                    if (caller == null)
                        continue;

                    total++;
                    if (total > MAX_FORCE_REFRESH_CALLERS)
                        break;

                    if (TryForceRefreshCaller(caller))
                        success++;
                    else
                        failed++;
                }

                if (total > MAX_FORCE_REFRESH_CALLERS)
                    break;
            }

            if (total > 0 && Interlocked.Exchange(ref _forceRefreshLogged, 1) == 0)
            {
                Logger.RtdServer.Info("RtdFallback - ForceRefresh attempted {0} cells (success {1}, failed {2})", total, success, failed);
            }
        }

        private static bool TryForceRefreshCaller(ExcelReference caller)
        {
            try
            {
                object app = ExcelDnaUtil.Application;
                if (app == null)
                    return false;

                object refTextObj = XlCall.Excel(XlCall.xlfReftext, caller, true);
                string refText = refTextObj as string;
                if (string.IsNullOrEmpty(refText))
                    return false;

                var culture = CultureInfo.GetCultureInfo(1033);
                object range = null;
                try
                {
                    range = app.GetType().InvokeMember("Range", BindingFlags.GetProperty, null, app, new object[] { refText }, culture);
                }
                catch
                {
                    range = app.GetType().InvokeMember("Range", BindingFlags.InvokeMethod, null, app, new object[] { refText }, culture);
                }

                if (range == null)
                    return false;

                Type rangeType = range.GetType();
                object formulaObj = rangeType.InvokeMember("Formula", BindingFlags.GetProperty, null, range, null, culture);
                string formulaText = formulaObj as string;
                if (string.IsNullOrEmpty(formulaText) || !formulaText.StartsWith("=", StringComparison.Ordinal))
                {
                    ReleaseComObject(range);
                    return false;
                }

                rangeType.InvokeMember("Formula", BindingFlags.SetProperty, null, range, new object[] { formulaObj }, culture);
                ReleaseComObject(range);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static void ReleaseComObject(object comObject)
        {
            try
            {
                if (comObject != null && Marshal.IsComObject(comObject))
                    Marshal.ReleaseComObject(comObject);
            }
            catch
            {
                // ignore
            }
        }

        private static void LogRecalcOutcomeOnce(bool success, string method, XlCall.XlReturn lastReturn)
        {
            if (Interlocked.Exchange(ref _recalcOutcomeLogged, 1) != 0)
                return;

            if (success)
            {
                Logger.RtdServer.Info("RtdFallback - Recalculation succeeded via {0}", string.IsNullOrEmpty(method) ? "unknown" : method);
            }
            else
            {
                Logger.RtdServer.Warn("RtdFallback - Recalculation failed. Last return: {0}", lastReturn);
            }
        }

        private static string DescribeValue(object value)
        {
            if (value == null)
                return "null";

            try
            {
                var text = Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
                if (text.Length > 200)
                    text = text.Substring(0, 200) + "...";
                return text;
            }
            catch
            {
                return value.GetType().FullName ?? "unknown";
            }
        }

        /// <summary>
        /// 启动缓存清理定时器
        /// </summary>
        private static void StartTimers()
        {
            lock (_timerLock)
            {
                // 启动缓存清理定时器
                if (_cacheCleanupTimer == null && CacheCleanupIntervalMs > 0)
                {
                    _cacheCleanupTimer = new Timer(
                        CleanupExpiredCache,
                        null,
                        CacheCleanupIntervalMs,
                        CacheCleanupIntervalMs);
                    Logger.RtdServer.Info("RtdFallbackManager - Cache cleanup timer started (interval: {0}ms)", CacheCleanupIntervalMs);
                }

                // 注意：旧的 VolatileRefreshTimer 已被移除，改用基于事件的 RequestBatchRecalculation
            }
        }

        /// <summary>
        /// 停止定时器
        /// </summary>
        public static void StopTimers()
        {
            lock (_timerLock)
            {
                if (_cacheCleanupTimer != null)
                {
                    _cacheCleanupTimer.Dispose();
                    _cacheCleanupTimer = null;
                    Logger.RtdServer.Info("RtdFallbackManager - Cache cleanup timer stopped");
                }

                if (_volatileRefreshTimer != null)
                {
                    _volatileRefreshTimer.Dispose();
                    _volatileRefreshTimer = null;
                    Logger.RtdServer.Info("RtdFallbackManager - Volatile refresh timer stopped");
                }
            }
        }

        /// <summary>
        /// 清理过期的缓存条目
        /// </summary>
        private static void CleanupExpiredCache(object state)
        {
            if (CacheExpirationMs <= 0)
                return;

            try
            {
                var expiredKeys = new List<AsyncCallInfo>();

                foreach (var kvp in _fallbackCache)
                {
                    if (kvp.Value.IsExpired(CacheExpirationMs))
                    {
                        expiredKeys.Add(kvp.Key);
                    }
                }

                foreach (var key in expiredKeys)
                {
                    if (_fallbackCache.TryRemove(key, out var entry))
                    {
                        entry.Subscription?.Dispose();
                        Debug.Print("RtdFallback - Expired cache entry removed: hash={0}", key.GetHashCode());
                    }
                }

                if (expiredKeys.Count > 0)
                {
                    Logger.RtdServer.Info("RtdFallbackManager - Cleaned up {0} expired cache entries", expiredKeys.Count);
                }
            }
            catch (Exception ex)
            {
                Debug.Print("RtdFallback - CleanupExpiredCache error: {0}", ex.Message);
            }
        }



        #endregion

        #region Cache Management

        /// <summary>
        /// 清理指定条目的缓存
        /// </summary>
        internal static void ClearCache(AsyncCallInfo callInfo)
        {
            if (_fallbackCache.TryRemove(callInfo, out var entry))
            {
                entry.Subscription?.Dispose();
                _activeObservables.TryRemove(callInfo, out _);
            }
        }

        /// <summary>
        /// 清理所有缓存
        /// </summary>
        public static void ClearAllCache()
        {
            foreach (var kvp in _fallbackCache)
            {
                kvp.Value.Subscription?.Dispose();
            }
            _fallbackCache.Clear();
            _activeObservables.Clear();
            Logger.RtdServer.Info("RtdFallbackManager - All cache cleared");
        }

        /// <summary>
        /// 缓存统计信息
        /// </summary>
        public class CacheStatistics
        {
            public int TotalEntries { get; set; }
            public int ActiveObservables { get; set; }
            public int CompletedEntries { get; set; }
        }

        /// <summary>
        /// 获取当前缓存统计信息
        /// </summary>
        public static CacheStatistics GetCacheStatistics()
        {
            return new CacheStatistics
            {
                TotalEntries = _fallbackCache.Count,
                ActiveObservables = _activeObservables.Count,
                CompletedEntries = _fallbackCache.Values.Count(e => e.IsComplete)
            };
        }

        #endregion
    }
}
