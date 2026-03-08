using SleepSentinel.Services;

namespace SleepSentinel.UI;

public sealed class TrayApplicationContext : ApplicationContext
{
    private readonly PowerController _controller;
    private readonly FileLogger _logger;
    private readonly NotifyIcon _notifyIcon;
    private readonly MainForm _mainForm;
    private readonly Icon _appIcon;
    private readonly ToolStripMenuItem _keepAwakeMenu;
    private readonly ToolStripMenuItem _followPowerPlanMenuItem;
    private readonly ToolStripMenuItem _keepAwakeMenuItem;
    private readonly EventHandler _stateChangedHandler;

    public TrayApplicationContext(PowerController controller, FileLogger logger, SettingsStore settingsStore, Icon appIcon)
    {
        _controller = controller;
        _logger = logger;
        _appIcon = (Icon)appIcon.Clone();

        _mainForm = new MainForm(controller, logger, settingsStore, _appIcon);
        _mainForm.FormClosed += (_, _) =>
        {
            if (_mainForm.Visible)
            {
                _mainForm.Hide();
            }
        };

        var menu = new ContextMenuStrip();
        menu.Items.Add("打开面板", null, (_, _) => ShowMainForm());
        menu.Items.Add(new ToolStripSeparator());

        _keepAwakeMenu = new ToolStripMenuItem("保持唤醒");

        _followPowerPlanMenuItem = new ToolStripMenuItem("遵循电源计划")
        {
            CheckOnClick = false
        };
        _followPowerPlanMenuItem.Click += (_, _) => _controller.SetPolicyMode(Models.PowerPolicyMode.FollowPowerPlan);
        _keepAwakeMenu.DropDownItems.Add(_followPowerPlanMenuItem);

        _keepAwakeMenuItem = new ToolStripMenuItem("无限保持唤醒（类似 PowerToys Awake）")
        {
            CheckOnClick = false
        };
        _keepAwakeMenuItem.Click += (_, _) => _controller.SetPolicyMode(Models.PowerPolicyMode.KeepAwakeIndefinitely);
        _keepAwakeMenu.DropDownItems.Add(_keepAwakeMenuItem);
        menu.Items.Add(_keepAwakeMenu);

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
        RefreshTrayText();

        if (!_controller.CurrentSettings.StartMinimized)
        {
            ShowMainForm();
        }
    }

    protected override void ExitThreadCore()
    {
        _controller.StateChanged -= _stateChangedHandler;
        _notifyIcon.Visible = false;
        _notifyIcon.Dispose();
        _mainForm.Dispose();
        _appIcon.Dispose();
        base.ExitThreadCore();
    }

    private void ShowMainForm()
    {
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
        _keepAwakeMenu.Text = $"保持唤醒: {_controller.CurrentStatus}";
        _followPowerPlanMenuItem.Checked = _controller.CurrentSettings.PolicyMode == Models.PowerPolicyMode.FollowPowerPlan;
        _keepAwakeMenuItem.Checked = _controller.CurrentSettings.PolicyMode == Models.PowerPolicyMode.KeepAwakeIndefinitely;
    }

    private void RefreshTrayTextOnUiThread()
    {
        if (_mainForm.IsDisposed)
        {
            return;
        }

        if (_mainForm.IsHandleCreated && _mainForm.InvokeRequired)
        {
            _mainForm.BeginInvoke(new Action(RefreshTrayTextOnUiThread));
            return;
        }

        RefreshTrayText();
    }
}
