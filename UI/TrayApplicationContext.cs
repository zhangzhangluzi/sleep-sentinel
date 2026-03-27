using SleepSentinel.Models;
using SleepSentinel.Services;
using System.Runtime.InteropServices;

namespace SleepSentinel.UI;

public sealed class TrayApplicationContext : ApplicationContext
{
    private readonly PowerController _controller;
    private readonly FileLogger _logger;
    private readonly NotifyIcon _notifyIcon;
    private readonly MainForm _mainForm;
    private readonly Icon _appIcon;
    private readonly SettingsStore _settingsStore;
    private readonly EventWaitHandle _activationSignal;
    private readonly RegisteredWaitHandle _activationWaitHandle;
    private readonly TrayMessageWindow _trayMessageWindow;
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
    private bool _isExiting;

    public TrayApplicationContext(
        PowerController controller,
        FileLogger logger,
        SettingsStore settingsStore,
        Icon appIcon,
        EventWaitHandle activationSignal)
    {
        _controller = controller;
        _logger = logger;
        _appIcon = (Icon)appIcon.Clone();
        _settingsStore = settingsStore;
        _activationSignal = activationSignal;
        _diagnosticReportService = new DiagnosticReportService(settingsStore, logger, controller);

        _mainForm = new MainForm(controller, logger, settingsStore, _appIcon);
        _mainForm.FormClosed += (_, _) =>
        {
            if (_mainForm.Visible)
            {
                _mainForm.Hide();
            }
        };
        _ = _mainForm.Handle;
        _trayMessageWindow = new TrayMessageWindow(OnTaskbarCreated);

        var menu = new ContextMenuStrip();
        menu.Items.Add("打开面板", null, (_, _) => ShowMainForm());
        menu.Items.Add(new ToolStripSeparator());

        _followPowerPlanMenuItem = new ToolStripMenuItem("遵循电源计划");
        _followPowerPlanMenuItem.Click += (_, _) => _controller.SetPolicyMode(PowerPolicyMode.FollowPowerPlan);
        menu.Items.Add(_followPowerPlanMenuItem);

        _keepAwakeMenuItem = new ToolStripMenuItem("无限保持唤醒（类似 PowerToys Awake）");
        _keepAwakeMenuItem.Click += (_, _) => _controller.SetPolicyMode(PowerPolicyMode.KeepAwakeIndefinitely);
        menu.Items.Add(_keepAwakeMenuItem);

        menu.Items.Add(new ToolStripSeparator());

        _wakeTimersMenuItem = new ToolStripMenuItem("接管唤醒定时器");
        _wakeTimersMenuItem.Click += (_, _) =>
        {
            if (_controller.CurrentSettings.DisableWakeTimers)
            {
                _controller.RestoreSoftwareWake();
            }
            else
            {
                _controller.BlockSoftwareWake();
            }
        };
        menu.Items.Add(_wakeTimersMenuItem);

        _standbyConnectivityMenuItem = new ToolStripMenuItem("接管待机联网");
        _standbyConnectivityMenuItem.Click += (_, _) =>
        {
            if (_controller.CurrentSettings.DisableStandbyConnectivity)
            {
                _controller.RestoreStandbyConnectivityWake();
            }
            else
            {
                _controller.BlockStandbyConnectivityWake();
            }
        };
        menu.Items.Add(_standbyConnectivityMenuItem);

        _batteryFallbackMenuItem = new ToolStripMenuItem("接管电池兜底休眠");
        _batteryFallbackMenuItem.Click += (_, _) =>
        {
            if (_controller.CurrentSettings.EnforceBatteryStandbyHibernate)
            {
                _controller.RestoreBatteryStandbyHibernateFallback();
            }
            else
            {
                _controller.EnableBatteryStandbyHibernateFallback();
            }
        };
        menu.Items.Add(_batteryFallbackMenuItem);

        _remoteWakeMenuItem = new ToolStripMenuItem("接管远控保活拦截");
        _remoteWakeMenuItem.Click += (_, _) =>
        {
            if (_controller.CurrentSettings.BlockKnownRemoteWakeRequests)
            {
                _controller.RestoreKnownRemoteWakeRequests();
            }
            else
            {
                _controller.BlockKnownRemoteWakeRequests();
            }
        };
        menu.Items.Add(_remoteWakeMenuItem);

        menu.Items.Add(new ToolStripSeparator());

        _startMinimizedMenuItem = new ToolStripMenuItem("启动后仅驻留托盘");
        _startMinimizedMenuItem.Click += (_, _) => ToggleSetting(static settings => settings.StartMinimized = !settings.StartMinimized);
        menu.Items.Add(_startMinimizedMenuItem);

        _autostartMenuItem = new ToolStripMenuItem("开机自启");
        _autostartMenuItem.Click += (_, _) => ToggleSetting(static settings => settings.StartWithWindows = !settings.StartWithWindows);
        menu.Items.Add(_autostartMenuItem);

        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add("记录唤醒诊断", null, (_, _) =>
        {
            var diagnostics = _controller.CollectFullWakeDiagnosticsText();
            _logger.Warn("用户从托盘菜单手动收集唤醒诊断。");
            _logger.Warn(diagnostics);
        });
        menu.Items.Add("导出诊断报告", null, (_, _) =>
        {
            var path = _diagnosticReportService.Export();
            ShowTrayBalloon($"诊断报告已导出：{path}", ToolTipIcon.Info);
        });
        menu.Items.Add("重新应用全部设置", null, (_, _) => _controller.ReapplyAllManagedSettings());

        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add("立即睡眠", null, (_, _) => _controller.SleepNow());
        menu.Items.Add("立即休眠", null, (_, _) => _controller.HibernateNow());
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
        _trayMessageWindow.Dispose();
        _notifyIcon.Visible = false;
        _notifyIcon.Dispose();
        _mainForm.Dispose();
        _appIcon.Dispose();
        base.ExitThreadCore();
    }

    private void ShowMainForm()
    {
        EnsureTrayIconVisible();
        _mainForm.Show();
        _mainForm.WindowState = FormWindowState.Normal;
        _mainForm.BringToFront();
        _mainForm.Activate();
    }

    private void RefreshTrayText()
    {
        var text = $"SleepSentinel - {_controller.CurrentStatus}";
        _notifyIcon.Text = text.Length > 63 ? text[..63] : text;
        _notifyIcon.BalloonTipTitle = "SleepSentinel";
        _notifyIcon.BalloonTipText = _controller.CurrentRiskSummary;
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
        if (_isExiting || _mainForm.IsDisposed)
        {
            return;
        }

        if (!_mainForm.IsHandleCreated)
        {
            return;
        }

        if (_mainForm.InvokeRequired)
        {
            _mainForm.BeginInvoke(new Action(RefreshTrayTextOnUiThread));
            return;
        }

        RefreshTrayText();
    }

    private void ToggleSetting(Action<AppSettings> mutation)
    {
        var updatedSettings = _settingsStore.Load();
        mutation(updatedSettings);
        _controller.UpdateSettings(updatedSettings);
    }

    private void OnExternalActivationRequested()
    {
        if (_isExiting || _mainForm.IsDisposed || !_mainForm.IsHandleCreated)
        {
            return;
        }

        try
        {
            _mainForm.BeginInvoke(new Action(() =>
            {
                if (_isExiting || _mainForm.IsDisposed)
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

    private void OnTaskbarCreated()
    {
        if (_isExiting || _mainForm.IsDisposed || !_mainForm.IsHandleCreated)
        {
            return;
        }

        try
        {
            _mainForm.BeginInvoke(new Action(() =>
            {
                if (_isExiting || _mainForm.IsDisposed)
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
