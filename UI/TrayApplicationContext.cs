using SleepSentinel.Services;

namespace SleepSentinel.UI;

public sealed class TrayApplicationContext : ApplicationContext
{
    private readonly PowerController _controller;
    private readonly FileLogger _logger;
    private readonly NotifyIcon _notifyIcon;
    private readonly MainForm _mainForm;
    private readonly Icon _appIcon;

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
        _controller.StateChanged += (_, _) => RefreshTrayText();
        RefreshTrayText();

        if (!_controller.CurrentSettings.StartMinimized)
        {
            ShowMainForm();
        }
    }

    protected override void ExitThreadCore()
    {
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
    }
}
