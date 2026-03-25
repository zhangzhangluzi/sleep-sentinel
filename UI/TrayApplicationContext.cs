using SleepSentinel.Services;
using System.Runtime.InteropServices;

namespace SleepSentinel.UI;

public sealed class TrayApplicationContext : ApplicationContext
{
    private readonly PowerController _controller;
    private readonly FileLogger _logger;
    private readonly SettingsStore _settingsStore;
    private readonly DiagnosticReportService _diagnosticReportService;
    private readonly NotifyIcon _notifyIcon;
    private readonly MainForm _mainForm;
    private readonly Icon _appIcon;
    private readonly EventWaitHandle _activationSignal;
    private readonly RegisteredWaitHandle _activationWaitHandle;
    private readonly TrayMessageWindow _trayMessageWindow;
    private readonly ToolStripMenuItem _statusMenuItem;
    private readonly ToolStripMenuItem _followPowerPlanMenuItem;
    private readonly ToolStripMenuItem _keepAwakeMenuItem;
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
        _settingsStore = settingsStore;
        _diagnosticReportService = new DiagnosticReportService(settingsStore, logger, controller);
        _appIcon = (Icon)appIcon.Clone();
        _activationSignal = activationSignal;

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
        _statusMenuItem = new ToolStripMenuItem("状态：正在初始化")
        {
            Enabled = false
        };
        menu.Items.Add(_statusMenuItem);
        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add("打开面板", null, (_, _) => ShowMainForm());
        menu.Items.Add(new ToolStripSeparator());

        _followPowerPlanMenuItem = new ToolStripMenuItem("遵循电源计划")
        {
            CheckOnClick = false
        };
        _followPowerPlanMenuItem.Click += (_, _) => _controller.SetPolicyMode(Models.PowerPolicyMode.FollowPowerPlan);
        menu.Items.Add(_followPowerPlanMenuItem);

        _keepAwakeMenuItem = new ToolStripMenuItem("无限保持唤醒（类似 PowerToys Awake）")
        {
            CheckOnClick = false
        };
        _keepAwakeMenuItem.Click += (_, _) => _controller.SetPolicyMode(Models.PowerPolicyMode.KeepAwakeIndefinitely);
        menu.Items.Add(_keepAwakeMenuItem);

        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add("立即睡眠", null, (_, _) => _controller.SleepNow());
        menu.Items.Add("立即休眠", null, (_, _) => _controller.HibernateNow());
        menu.Items.Add("记录唤醒诊断", null, (_, _) => RecordWakeDiagnosticsFromTray());
        menu.Items.Add("导出诊断报告", null, (_, _) => ExportDiagnosticReportFromTray());

        menu.Items.Add(new ToolStripSeparator());
        _startMinimizedMenuItem = new ToolStripMenuItem("启动后仅驻留托盘")
        {
            CheckOnClick = false
        };
        _startMinimizedMenuItem.Click += (_, _) => ToggleStartMinimized();
        menu.Items.Add(_startMinimizedMenuItem);

        _autostartMenuItem = new ToolStripMenuItem("开机自启")
        {
            CheckOnClick = false
        };
        _autostartMenuItem.Click += (_, _) => ToggleAutostart();
        menu.Items.Add(_autostartMenuItem);

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
        _notifyIcon.BalloonTipText = _controller.CurrentStatus;
        _statusMenuItem.Text = $"状态：{_controller.CurrentStatus}";
        _followPowerPlanMenuItem.Checked = _controller.CurrentSettings.PolicyMode == Models.PowerPolicyMode.FollowPowerPlan;
        _keepAwakeMenuItem.Checked = _controller.CurrentSettings.PolicyMode == Models.PowerPolicyMode.KeepAwakeIndefinitely;
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

    private void ToggleStartMinimized()
    {
        var settings = _settingsStore.Load();
        settings.StartMinimized = !settings.StartMinimized;
        _controller.UpdateSettings(settings);
    }

    private void ToggleAutostart()
    {
        var settings = _settingsStore.Load();
        settings.StartWithWindows = !settings.StartWithWindows;
        _controller.UpdateSettings(settings);
    }

    private void RecordWakeDiagnosticsFromTray()
    {
        _logger.Warn("用户从托盘手动收集唤醒诊断。");
        _logger.Warn(_controller.CollectWakeDiagnostics());
        _notifyIcon.ShowBalloonTip(2500, "SleepSentinel", "已记录唤醒诊断到日志。", ToolTipIcon.Info);
    }

    private void ExportDiagnosticReportFromTray()
    {
        try
        {
            var path = _diagnosticReportService.Export();
            _notifyIcon.ShowBalloonTip(3000, "SleepSentinel", $"诊断报告已导出：{path}", ToolTipIcon.Info);
        }
        catch (Exception ex)
        {
            _logger.Error($"从托盘导出诊断报告失败：{ex.Message}");
            MessageBox.Show($"导出失败：{ex.Message}", "SleepSentinel", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
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
