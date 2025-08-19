using System;
using System.Collections.Generic;
using System.Timers;

namespace Desktop_Fences
{
    public class TargetChecker
    {
        private readonly Timer _timer;
        private readonly Dictionary<string, (Action checkAction, bool isFolder)> _checkActions;
        private readonly object _lockObject = new object();

        public TargetChecker(double interval)
        {
            _timer = new Timer(interval);
            _timer.Elapsed += OnTimedEvent;
            _checkActions = new Dictionary<string, (Action, bool)>();
            Start(); // Start the timer immediately
        }

        public void Start()
        {
            _timer.Start();
        }

        public void Stop()
        {
            _timer.Stop();
        }

        public void AddCheckAction(string key, Action checkAction, bool isFolder)
        {
            if (checkAction == null) return;

            lock (_lockObject)
            {
                if (!_checkActions.ContainsKey(key))
                {
                    _checkActions.Add(key, (checkAction, isFolder));
                }
            }

            // Run immediately outside the lock to prevent potential deadlocks
            try
            {
                checkAction.Invoke();
            }
            catch (Exception ex)
            {
                LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.Error,
                    $"Error in immediate target check for {key}: {ex.Message}");
            }
        }
        public void RemoveCheckAction(string key)
        {
            lock (_lockObject)
            {
                if (_checkActions.ContainsKey(key))
                {
                    _checkActions.Remove(key);
                }
            }
        }

        private void OnTimedEvent(object sender, ElapsedEventArgs e)
        {
            try
            {
                List<(Action checkAction, bool isFolder)> actionsSnapshot;

                // Create a thread-safe snapshot of the actions
                // lock (_checkActions)
                lock (_lockObject)
                {
                    if (_checkActions.Count == 0)
                        return;

                    actionsSnapshot = new List<(Action, bool)>(_checkActions.Count);
                    foreach (var kvp in _checkActions)
                    {
                        actionsSnapshot.Add(kvp.Value);
                    }
                }

                // Execute actions outside the lock to prevent deadlocks
                foreach (var (checkAction, isFolder) in actionsSnapshot)
                {
                    try
                    {
                        checkAction?.Invoke();
                    }
                    catch (Exception actionEx)
                    {
                        // Log individual action errors but continue with other actions
                        LogManager.Log(LogManager.LogLevel.Warn, LogManager.LogCategory.Error,
                            $"Error in target check action: {actionEx.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                // Log timer event errors
                LogManager.Log(LogManager.LogLevel.Error, LogManager.LogCategory.Error,
                    $"Error in TargetChecker timer event: {ex.Message}");
            }
        }
    }
}