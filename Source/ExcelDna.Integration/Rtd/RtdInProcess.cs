//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Excel-DNA is licensed under the zlib license. See LICENSE.txt for details.

using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Threading;
using ExcelDna.Logging;

namespace ExcelDna.Integration.Rtd
{
    internal static class RtdInProcess
    {
        static readonly ConcurrentDictionary<string, RtdServerHost> Hosts = new ConcurrentDictionary<string, RtdServerHost>();

        internal static bool TryCreateObservable(string progId, string server, string[] topics, out IExcelObservable observable)
        {
            observable = null;
            if (string.IsNullOrEmpty(progId))
                return false;

            if (!RtdRegistration.TryGetRtdServerType(progId, out var rtdServerType))
                return false;

            observable = new RtdServerObservable(progId, server, topics ?? Array.Empty<string>());
            return true;
        }

        static RtdServerHost GetOrCreateHost(string progId, Type rtdServerType)
        {
            return Hosts.GetOrAdd(progId, _ => new RtdServerHost(progId, rtdServerType));
        }

        static void RemoveHost(string progId, RtdServerHost host)
        {
            Hosts.TryRemove(progId, out _);
            host.Dispose();
        }

        sealed class RtdServerObservable : IExcelObservable
        {
            readonly string _progId;
            readonly string _server;
            readonly string[] _topics;

            public RtdServerObservable(string progId, string server, string[] topics)
            {
                _progId = progId;
                _server = server;
                _topics = topics ?? Array.Empty<string>();
            }

            public IDisposable Subscribe(IExcelObserver observer)
            {
                if (observer == null)
                    throw new ArgumentNullException(nameof(observer));

                if (!RtdRegistration.TryGetRtdServerType(_progId, out var rtdServerType))
                {
                    observer.OnError(new InvalidOperationException("RTD server type not registered: " + _progId));
                    return new NoopDisposable();
                }

                var host = GetOrCreateHost(_progId, rtdServerType);
                return host.Subscribe(_topics, observer, _progId, RemoveHost);
            }
        }

        sealed class RtdServerHost : IRTDUpdateEvent, IDisposable
        {
            readonly string _progId;
            readonly IRtdServer _server;
            readonly object _lock = new object();
            readonly Dictionary<int, IExcelObserver> _observers = new Dictionary<int, IExcelObserver>();
            int _topicIdCounter;
            int _isRefreshing;
            int _heartbeatInterval = 15000;

            public RtdServerHost(string progId, Type rtdServerType)
            {
                _progId = progId;
                _server = (IRtdServer)Activator.CreateInstance(rtdServerType);
                int started = _server.ServerStart(this);
                if (started == 0)
                    throw new InvalidOperationException("RTD ServerStart failed for " + progId);

                Logger.RtdServer.Info("RtdInProcess - Started RTD host for {0}", _progId);
            }

            public IDisposable Subscribe(string[] topics, IExcelObserver observer, string hostKey, Action<string, RtdServerHost> onEmpty)
            {
                int topicId = Interlocked.Increment(ref _topicIdCounter);
                bool newValues = true;
                Array topicArray = topics ?? Array.Empty<string>();

                lock (_lock)
                {
                    _observers[topicId] = observer;
                }

                try
                {
                    object value = _server.ConnectData(topicId, ref topicArray, ref newValues);
                    observer.OnNext(value);
                }
                catch (Exception ex)
                {
                    Unsubscribe(topicId, observer, hostKey, onEmpty);
                    observer.OnError(ex);
                    return new NoopDisposable();
                }

                return new Subscription(this, topicId, hostKey, onEmpty);
            }

            public void UpdateNotify()
            {
                if (Interlocked.Exchange(ref _isRefreshing, 1) == 1)
                    return;

                try
                {
                    int topicCount = 0;
                    Array data = _server.RefreshData(ref topicCount);
                    if (data == null || topicCount <= 0)
                        return;

                    object[,] values = data as object[,];
                    if (values == null)
                        return;

                    for (int i = 0; i < topicCount; i++)
                    {
                        int topicId;
                        try
                        {
                            topicId = Convert.ToInt32(values[0, i]);
                        }
                        catch
                        {
                            continue;
                        }

                        IExcelObserver observer;
                        lock (_lock)
                        {
                            if (!_observers.TryGetValue(topicId, out observer))
                                continue;
                        }

                        observer.OnNext(values[1, i]);
                    }
                }
                catch (Exception ex)
                {
                    Logger.RtdServer.Warn("RtdInProcess UpdateNotify failed for {0}: {1}", _progId, ex.Message);
                }
                finally
                {
                    Interlocked.Exchange(ref _isRefreshing, 0);
                }
            }

            public int HeartbeatInterval
            {
                get => _heartbeatInterval;
                set => _heartbeatInterval = value;
            }

            public void Disconnect()
            {
                List<int> topicIds;
                lock (_lock)
                {
                    topicIds = new List<int>(_observers.Keys);
                }

                foreach (var id in topicIds)
                {
                    Unsubscribe(id, null, null, null);
                }
            }

            void Unsubscribe(int topicId, IExcelObserver observer, string hostKey, Action<string, RtdServerHost> onEmpty)
            {
                lock (_lock)
                {
                    _observers.Remove(topicId);
                }

                try
                {
                    _server.DisconnectData(topicId);
                }
                catch
                {
                    // Ignore disconnect errors.
                }

                if (hostKey != null && onEmpty != null)
                {
                    bool empty;
                    lock (_lock)
                    {
                        empty = _observers.Count == 0;
                    }
                    if (empty)
                    {
                        onEmpty(hostKey, this);
                    }
                }

                observer?.OnCompleted();
            }

            public void Dispose()
            {
                try
                {
                    _server.ServerTerminate();
                }
                catch
                {
                    // Ignore terminate errors.
                }
            }

            sealed class Subscription : IDisposable
            {
                readonly RtdServerHost _host;
                readonly int _topicId;
                readonly string _hostKey;
                readonly Action<string, RtdServerHost> _onEmpty;
                int _disposed;

                public Subscription(RtdServerHost host, int topicId, string hostKey, Action<string, RtdServerHost> onEmpty)
                {
                    _host = host;
                    _topicId = topicId;
                    _hostKey = hostKey;
                    _onEmpty = onEmpty;
                }

                public void Dispose()
                {
                    if (Interlocked.Exchange(ref _disposed, 1) != 0)
                        return;

                    _host.Unsubscribe(_topicId, null, _hostKey, _onEmpty);
                }
            }
        }

        sealed class NoopDisposable : IDisposable
        {
            public void Dispose() { }
        }
    }
}
