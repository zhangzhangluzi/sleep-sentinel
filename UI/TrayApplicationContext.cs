using SleepSentinel.Models;
using SleepSentinel.Services;
using System.Runtime.InteropServices;

namespace SleepSentinel.UI;

public sealed class TrayApplicationContext : ApplicationContext
{
    private const int DeferredWarmupDelayMilliseconds = 20000;
    private readonly PowerController _controller;
    private readonly FileLogger _logger;
    private readonly NotifyIcon _notifyIcon;
    private MainForm? _mainForm;
    private readonly Icon _appIcon;
    private readonly SettingsStore _settingsStore;
    private readonly EventWaitHandle _activationSignal;
    private readonly EventWaitHandle _takeoverSignal;
    private readonly RegisteredWaitHandle _activationWaitHandle;
    private readonly RegisteredWaitHandle _takeoverWaitHandle;
    private readonly TrayMessageWindow _trayMessageWindow;
    private readonly Control _uiInvoker;
    private readonly System.Threading.Timer _deferredWarmupTimer;
    private readonly DiagnosticReportService _diagnosticReportService;
    private readonly ToolStripMenuItem _followPowerPlanMenuItem;
    private readonly ToolStripMenuItem _keepAwakeMenuItem;
    private readonly ToolStripMenuItem _wakeTimersMenuItem;
    private readonly ToolStripMenuItem _standbyConnectivityMenuItem;
    private readonly ToolStripMenuItem _batteryFallbackMenuItem;
    private readonly ToolStripMenuItem _remoteWakeMenuItem;
    private readonly ToolStripMenuItem _startMinimizedMenuItem;
    private readonly ToolStripMenuItem _autostartMenuItem;
    private readonly EventHandler _stateChangedHandler;
    private int _deferredWarmupScheduled;
    private bool _isExiting;

    public TrayApplicationContext(
        PowerController controller,
        FileLogger logger,
        SettingsStore settingsStore,
        Icon appIcon,
        EventWaitHandle activationSignal,
        EventWaitHandle takeoverSignal)
    {
        _controller = controller;
        _logger = logger;
        _appIcon = (Icon)appIcon.Clone();
        _settingsStore = settingsStore;
        _activationSignal = activationSignal;
        _takeoverSignal = takeoverSignal;
        _diagnosticReportService = new DiagnosticReportService(settingsStore, logger, controller);
        _uiInvoker = new Control();
        _ = _uiInvoker.Handle;
        _trayMessageWindow = new TrayMessageWindow(OnTaskbarCreated);

        var menu = new ContextMenuStrip();
        menu.Items.Add("打开面板", null, (_, _) => ShowMainForm());
        menu.Items.Add(new ToolStripSeparator());

        _followPowerPlanMenuItem = new ToolStripMenuItem("遵循电源计划");
        _followPowerPlanMenuItem.Click += (_, _) => RunTrayAction("切换为遵循电源计划", () => _controller.SetPolicyMode(PowerPolicyMode.FollowPowerPlan));
        menu.Items.Add(_followPowerPlanMenuItem);

        _keepAwakeMenuItem = new ToolStripMenuItem("无限保持唤醒（类似 PowerToys Awake）");
        _keepAwakeMenuItem.Click += (_, _) => RunTrayAction("切换为无限保持唤醒", () => _controller.SetPolicyMode(PowerPolicyMode.KeepAwakeIndefinitely));
        menu.Items.Add(_keepAwakeMenuItem);

        menu.Items.Add(new ToolStripSeparator());

        _wakeTimersMenuItem = new ToolStripMenuItem("接管唤醒定时器");
        _wakeTimersMenuItem.Click += (_, _) =>
        {
            RunTrayAction("切换唤醒定时器接管", () =>
            {
                if (_controller.CurrentSettings.DisableWakeTimers)
                {
                    _controller.RestoreSoftwareWake();
                }
                else
                {
                    _controller.BlockSoftwareWake();
                }
            });
        };
        menu.Items.Add(_wakeTimersMenuItem);

        _standbyConnectivityMenuItem = new ToolStripMenuItem("接管待机联网");
        _standbyConnectivityMenuItem.Click += (_, _) =>
        {
            RunTrayAction("切换待机联网接管", () =>
            {
                if (_controller.CurrentSettings.DisableStandbyConnectivity)
                {
                    _controller.RestoreStandbyConnectivityWake();
                }
                else
                {
                    _controller.BlockStandbyConnectivityWake();
                }
            });
        };
        menu.Items.Add(_standbyConnectivityMenuItem);

        _batteryFallbackMenuItem = new ToolStripMenuItem("接管电池兜底休眠");
        _batteryFallbackMenuItem.Click += (_, _) =>
        {
            RunTrayAction("切换电池兜底休眠接管", () =>
            {
                if (_controller.CurrentSettings.EnforceBatteryStandbyHibernate)
                {
                    _controller.RestoreBatteryStandbyHibernateFallback();
                }
                else
                {
                    _controller.EnableBatteryStandbyHibernateFallback();
                }
            });
        };
        menu.Items.Add(_batteryFallbackMenuItem);

        _remoteWakeMenuItem = new ToolStripMenuItem("接管远控保活拦截");
        _remoteWakeMenuItem.Click += (_, _) =>
        {
            RunTrayAction("切换远控保活拦截接管", () =>
            {
                if (_controller.CurrentSettings.BlockKnownRemoteWakeRequests)
                {
                    _controller.RestoreKnownRemoteWakeRequests();
                }
                else
                {
                    _controller.BlockKnownRemoteWakeRequests();
                }
            });
        };
        menu.Items.Add(_remoteWakeMenuItem);

        menu.Items.Add(new ToolStripSeparator());

        _startMinimizedMenuItem = new ToolStripMenuItem("启动后仅驻留托盘");
        _startMinimizedMenuItem.Click += (_, _) => RunTrayAction("切换启动后仅驻留托盘", () => ToggleSetting(static settings => settings.StartMinimized = !settings.StartMinimized));
        menu.Items.Add(_startMinimizedMenuItem);

        _autostartMenuItem = new ToolStripMenuItem("开机自启");
        _autostartMenuItem.Click += (_, _) => RunTrayAction("切换开机自启", () => ToggleSetting(static settings => settings.StartWithWindows = !settings.StartWithWindows));
        menu.Items.Add(_autostartMenuItem);

        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add("记录唤醒诊断", null, (_, _) =>
        {
            RunTrayAction("记录唤醒诊断", () =>
            {
                var diagnostics = _controller.CollectFullWakeDiagnosticsText();
                _logger.Warn("用户从托盘菜单手动收集唤醒诊断。");
                _logger.Warn(diagnostics);
            });
        });
        menu.Items.Add("导出诊断报告", null, (_, _) =>
        {
            RunTrayAction("导出诊断报告", () =>
            {
                var path = _diagnosticReportService.Export();
                ShowTrayBalloon($"诊断报告已导出：{path}", ToolTipIcon.Info);
            });
        });
        menu.Items.Add("重新应用全部设置", null, (_, _) => RunTrayAction("重新应用全部设置", () => _controller.ReapplyAllManagedSettings()));

        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add("立即睡眠", null, (_, _) => RunTrayAction("立即睡眠", () => _controller.SleepNow()));
        menu.Items.Add("立即休眠", null, (_, _) => RunTrayAction("立即休眠", () => _controller.HibernateNow()));
        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add("退出", null, (_, _) => ExitThread());

        _notifyIcon = new NotifyIcon
        {
            Icon = _appIcon,
            Visible = true,
            Text = "SleepSentinel",
            ContextMenuStrip = menu
        };

        _notifyIcon.DoubleClick += (_, _) => ShowMainForm();
        _stateChangedHandler = (_, _) => RefreshTrayTextOnUiThread();
        _controller.StateChanged += _stateChangedHandler;
        _activationWaitHandle = ThreadPool.RegisterWaitForSingleObject(
            _activationSignal,
            static (state, _) => ((TrayApplicationContext)state!).OnExternalActivationRequested(),
            this,
            Timeout.Infinite,
            false);
        _takeoverWaitHandle = ThreadPool.RegisterWaitForSingleObject(
            _takeoverSignal,
            static (state, _) => ((TrayApplicationContext)state!).OnTakeoverRequested(),
            this,
            Timeout.Infinite,
            false);
        _deferredWarmupTimer = new System.Threading.Timer(
            static state => ((TrayApplicationContext)state!).OnDeferredWarmupTimerElapsed(),
            this,
            DeferredWarmupDelayMilliseconds,
            Timeout.Infinite);
        RefreshTrayText();

        if (!_controller.CurrentSettings.StartMinimized)
        {
            ShowMainForm();
        }
    }

    protected override void ExitThreadCore()
    {
        _isExiting = true;
        _controller.StateChanged -= _stateChangedHandler;
        _activationWaitHandle.Unregister(null);
        _takeoverWaitHandle.Unregister(null);
        _deferredWarmupTimer.Dispose();
        _trayMessageWindow.Dispose();
        _notifyIcon.Visible = false;
        _notifyIcon.Dispose();
        _mainForm?.Dispose();
        _uiInvoker.Dispose();
        _appIcon.Dispose();
        base.ExitThreadCore();
    }

    private void ShowMainForm()
    {
        var mainForm = EnsureMainForm();
        EnsureTrayIconVisible();
        mainForm.Show();
        mainForm.WindowState = FormWindowState.Normal;
        mainForm.RefreshFromController();
        mainForm.BringToFront();
        mainForm.Activate();
        StartDeferredWarmupInBackground();
    }

    private MainForm EnsureMainForm()
    {
        if (_mainForm is not null && !_mainForm.IsDisposed)
        {
            return _mainForm;
        }

        _mainForm = new MainForm(_controller, _logger, _settingsStore, _appIcon);
        _mainForm.FormClosed += (_, _) =>
        {
            if (_mainForm is not null && _mainForm.Visible)
            {
                _mainForm.Hide();
            }
        };

        return _mainForm;
    }

    private void OnDeferredWarmupTimerElapsed()
    {
        if (_isExiting)
        {
            return;
        }

        StartDeferredWarmupInBackground();
    }

    private void StartDeferredWarmupInBackground()
    {
        _deferredWarmupTimer.Change(Timeout.Infinite, Timeout.Infinite);
        if (_controller.StartupWarmupCompleted)
        {
            return;
        }

        if (System.Threading.Interlocked.Exchange(ref _deferredWarmupScheduled, 1) != 0)
        {
            return;
        }

        try
        {
            System.Threading.ThreadPool.UnsafeQueueUserWorkItem(
                static state =>
                {
                    var context = (TrayApplicationContext)state!;
                    try
                    {
                        context._controller.CompleteDeferredStartupValidation();
                    }
                    catch (Exception ex)
                    {
                        context._logger.Warn($"启动后后台校验失败：{ex.Message}");
                    }
                    finally
                    {
                        if (!context._controller.StartupWarmupCompleted)
                        {
                            System.Threading.Interlocked.Exchange(ref context._deferredWarmupScheduled, 0);
                        }
                    }
                },
                this);
        }
        catch (Exception ex)
        {
            System.Threading.Interlocked.Exchange(ref _deferredWarmupScheduled, 0);
            _logger.Warn($"启动后后台校验启动失败：{ex.Message}");
        }
    }

    private void RefreshTrayText()
    {
        var text = $"SleepSentinel - {_controller.CurrentStatus}";
        _notifyIcon.Text = text.Length > 63 ? text[..63] : text;
        _notifyIcon.BalloonTipTitle = "SleepSentinel";
        _notifyIcon.BalloonTipText = _controller.StartupWarmupCompleted
            ? _controller.CurrentRiskSummary
            : "SleepSentinel 已启动，详细状态会在后台补全。";
        _followPowerPlanMenuItem.Checked = _controller.CurrentSettings.PolicyMode == PowerPolicyMode.FollowPowerPlan;
        _keepAwakeMenuItem.Checked = _controller.CurrentSettings.PolicyMode == PowerPolicyMode.KeepAwakeIndefinitely;
        _wakeTimersMenuItem.Checked = _controller.CurrentSettings.DisableWakeTimers;
        _standbyConnectivityMenuItem.Checked = _controller.CurrentSettings.DisableStandbyConnectivity;
        _batteryFallbackMenuItem.Checked = _controller.CurrentSettings.EnforceBatteryStandbyHibernate;
        _remoteWakeMenuItem.Checked = _controller.CurrentSettings.BlockKnownRemoteWakeRequests;
        _startMinimizedMenuItem.Checked = _controller.CurrentSettings.StartMinimized;
        _autostartMenuItem.Checked = _controller.CurrentSettings.StartWithWindows;
    }

    private void RefreshTrayTextOnUiThread()
    {
        if (_isExiting)
        {
            return;
        }

        if (!_uiInvoker.IsHandleCreated)
        {
            return;
        }

        if (_uiInvoker.InvokeRequired)
        {
            try
            {
                _uiInvoker.BeginInvoke(new Action(RefreshTrayTextOnUiThread));
            }
            catch (InvalidOperationException)
            {
            }

            return;
        }

        RefreshTrayText();
    }

    private void ToggleSetting(Action<AppSettings> mutation)
    {
        var updatedSettings = _controller.CurrentSettings;
        mutation(updatedSettings);
        _controller.UpdateSettings(updatedSettings);
    }

    private void RunTrayAction(string actionName, Action action)
    {
        try
        {
            action();
        }
        catch (Exception ex)
        {
            _logger.Error($"托盘操作失败（{actionName}）：{ex.Message}");
            ShowTrayBalloon($"{actionName}失败：{ex.Message}", ToolTipIcon.Error);
        }
    }

    private void OnExternalActivationRequested()
    {
        if (_isExiting || !_uiInvoker.IsHandleCreated)
        {
            return;
        }

        try
        {
            _uiInvoker.BeginInvoke(new Action(() =>
            {
                if (_isExiting)
                {
                    return;
                }

                _logger.Info("收到新的启动请求，已唤回现有实例。");
                ShowMainForm();
            }));
        }
        catch (InvalidOperationException)
        {
        }
    }

    private void OnTakeoverRequested()
    {
        if (_isExiting || !_uiInvoker.IsHandleCreated)
        {
            return;
        }

        try
        {
            _uiInvoker.BeginInvoke(new Action(() =>
            {
                if (_isExiting)
                {
                    return;
                }

                _logger.Warn("收到更高权限实例的接管请求，正在退出当前实例。");
                ExitThread();
            }));
        }
        catch (InvalidOperationException)
        {
        }
    }

    private void OnTaskbarCreated()
    {
        if (_isExiting || !_uiInvoker.IsHandleCreated)
        {
            return;
        }

        try
        {
            _uiInvoker.BeginInvoke(new Action(() =>
            {
                if (_isExiting)
                {
                    return;
                }

                _logger.Warn("检测到任务栏已重建，正在恢复托盘图标。");
                EnsureTrayIconVisible();
            }));
        }
        catch (InvalidOperationException)
        {
        }
    }

    private void EnsureTrayIconVisible()
    {
        _notifyIcon.Visible = false;
        _notifyIcon.Icon = _appIcon;
        RefreshTrayText();
        _notifyIcon.Visible = true;
    }

    private void ShowTrayBalloon(string message, ToolTipIcon icon)
    {
        _notifyIcon.ShowBalloonTip(2000, "SleepSentinel", message, icon);
    }

    private sealed class TrayMessageWindow : NativeWindow, IDisposable
    {
        private readonly Action _taskbarCreatedCallback;
        private readonly int _taskbarCreatedMessage = RegisterWindowMessage("TaskbarCreated");

        public TrayMessageWindow(Action taskbarCreatedCallback)
        {
            _taskbarCreatedCallback = taskbarCreatedCallback;
            CreateHandle(new CreateParams());
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == _taskbarCreatedMessage)
            {
                _taskbarCreatedCallback();
            }

            base.WndProc(ref m);
        }

        public void Dispose()
        {
            if (Handle != IntPtr.Zero)
            {
                DestroyHandle();
            }
        }

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int RegisterWindowMessage(string lpString);
    }
}
