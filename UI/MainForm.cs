using SleepSentinel.Models;
using SleepSentinel.Services;

namespace SleepSentinel.UI;

public sealed class MainForm : Form
{
    private readonly PowerController _controller;
    private readonly FileLogger _logger;
    private readonly SettingsStore _settingsStore;
    private readonly DiagnosticReportService _diagnosticReportService;
    private readonly ComboBox _policyModeComboBox;
    private readonly CheckBox _resumeProtectionCheckbox;
    private readonly ComboBox _resumeActionComboBox;
    private readonly NumericUpDown _resumeDelayInput;
    private readonly CheckBox _onlyUnattendedWakeCheckbox;
    private readonly CheckBox _disableWakeTimersCheckbox;
    private readonly CheckBox _startMinimizedCheckbox;
    private readonly CheckBox _autostartCheckbox;
    private readonly Label _statusLabel;
    private readonly TextBox _logTextBox;
    private readonly EventHandler _stateChangedHandler;
    private readonly Icon _appIcon;

    public MainForm(PowerController controller, FileLogger logger, SettingsStore settingsStore, Icon appIcon)
    {
        _controller = controller;
        _logger = logger;
        _settingsStore = settingsStore;
        _diagnosticReportService = new DiagnosticReportService(settingsStore, logger, controller);
        _appIcon = (Icon)appIcon.Clone();

        Text = "SleepSentinel";
        Width = 860;
        Height = 640;
        MinimumSize = new Size(760, 540);
        StartPosition = FormStartPosition.CenterScreen;
        Icon = _appIcon;

        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 4,
            Padding = new Padding(16)
        };
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

        var intro = new Label
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            Text = "让电脑在“该睡就睡”时不要被其他软件长期唤醒，同时保留类似 PowerToys Awake 的无限保持唤醒模式。"
        };
        root.Controls.Add(intro);

        var settingsPanel = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            ColumnCount = 2,
            AutoSize = true,
            Padding = new Padding(0, 12, 0, 12)
        };
        settingsPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 220));
        settingsPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

        _policyModeComboBox = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Dock = DockStyle.Left, Width = 280 };
        _policyModeComboBox.Items.AddRange(new object[]
        {
            "遵循电源计划",
            "无限保持唤醒（类似 PowerToys Awake）"
        });

        _resumeProtectionCheckbox = new CheckBox { AutoSize = true, Text = "恢复后自动重新进入睡眠/休眠，避免被软件唤醒后常驻亮机" };

        _resumeActionComboBox = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Dock = DockStyle.Left, Width = 180 };
        _resumeActionComboBox.Items.AddRange(new object[] { "睡眠", "休眠" });

        _resumeDelayInput = new NumericUpDown
        {
            Minimum = 3,
            Maximum = 600,
            Value = 8,
            Width = 100
        };

        _onlyUnattendedWakeCheckbox = new CheckBox
        {
            AutoSize = true,
            Text = "仅在人工行为时跳过；人工包括键盘、鼠标、开盖、解锁、控制台/远程接管、登录"
        };

        _disableWakeTimersCheckbox = new CheckBox
        {
            AutoSize = true,
            Text = "同时关闭当前电源计划的唤醒定时器（AC/DC）"
        };

        _startMinimizedCheckbox = new CheckBox { AutoSize = true, Text = "启动后仅驻留托盘" };
        _autostartCheckbox = new CheckBox { AutoSize = true, Text = "开机自启" };

        AddSettingRow(settingsPanel, 0, "工作模式", _policyModeComboBox);
        AddSettingRow(settingsPanel, 1, "恢复保护", _resumeProtectionCheckbox);

        var resumeOptionsFlow = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Left };
        resumeOptionsFlow.Controls.Add(new Label { Text = "恢复后执行", AutoSize = true, Padding = new Padding(0, 7, 8, 0) });
        resumeOptionsFlow.Controls.Add(_resumeActionComboBox);
        resumeOptionsFlow.Controls.Add(new Label { Text = "延迟秒数", AutoSize = true, Padding = new Padding(16, 7, 8, 0) });
        resumeOptionsFlow.Controls.Add(_resumeDelayInput);
        AddSettingRow(settingsPanel, 2, "保护细节", resumeOptionsFlow);
        AddSettingRow(settingsPanel, 3, "唤醒过滤", _onlyUnattendedWakeCheckbox);
        AddSettingRow(settingsPanel, 4, "电源计划", _disableWakeTimersCheckbox);

        var startupFlow = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Left };
        startupFlow.Controls.Add(_startMinimizedCheckbox);
        startupFlow.Controls.Add(_autostartCheckbox);
        AddSettingRow(settingsPanel, 5, "启动行为", startupFlow);

        root.Controls.Add(settingsPanel);

        var actionsFlow = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            Padding = new Padding(0, 0, 0, 12)
        };

        var saveButton = new Button { Text = "保存设置", AutoSize = true };
        saveButton.Click += (_, _) => SaveSettings();
        var sleepButton = new Button { Text = "立即睡眠", AutoSize = true };
        sleepButton.Click += (_, _) => _controller.SleepNow();
        var hibernateButton = new Button { Text = "立即休眠", AutoSize = true };
        hibernateButton.Click += (_, _) => _controller.HibernateNow();
        var refreshLogButton = new Button { Text = "刷新日志", AutoSize = true };
        refreshLogButton.Click += (_, _) => LoadLogs();
        var diagnosticsButton = new Button { Text = "记录唤醒诊断", AutoSize = true };
        diagnosticsButton.Click += (_, _) =>
        {
            _logger.Warn("用户手动收集唤醒诊断。");
            _logger.Warn(_controller.CollectWakeDiagnostics());
            LoadLogs();
        };
        var blockSoftwareWakeButton = new Button { Text = "禁止软件唤醒", AutoSize = true };
        blockSoftwareWakeButton.Click += (_, _) =>
        {
            _controller.BlockSoftwareWake();
            ApplySettingsToUi(_controller.CurrentSettings);
            UpdateStatus();
            LoadLogs();
        };
        var reapplyWakeTimerButton = new Button { Text = "重新应用唤醒定时器策略", AutoSize = true };
        reapplyWakeTimerButton.Click += (_, _) =>
        {
            _controller.ReapplyWakeTimerPolicy();
            UpdateStatus();
            LoadLogs();
        };
        var exportReportButton = new Button { Text = "导出诊断报告", AutoSize = true };
        exportReportButton.Click += (_, _) => ExportDiagnosticReport();

        actionsFlow.Controls.Add(saveButton);
        actionsFlow.Controls.Add(sleepButton);
        actionsFlow.Controls.Add(hibernateButton);
        actionsFlow.Controls.Add(refreshLogButton);
        actionsFlow.Controls.Add(diagnosticsButton);
        actionsFlow.Controls.Add(blockSoftwareWakeButton);
        actionsFlow.Controls.Add(reapplyWakeTimerButton);
        actionsFlow.Controls.Add(exportReportButton);

        _statusLabel = new Label { AutoSize = true, Dock = DockStyle.Top, Padding = new Padding(0, 0, 0, 12) };

        root.Controls.Add(actionsFlow);
        root.Controls.Add(_statusLabel);

        _logTextBox = new TextBox
        {
            Multiline = true,
            ReadOnly = true,
            ScrollBars = ScrollBars.Vertical,
            Dock = DockStyle.Fill,
            Font = new Font(FontFamily.GenericMonospace, 9.5f)
        };
        root.Controls.Add(_logTextBox);

        Controls.Add(root);

        _stateChangedHandler = (_, _) => UpdateStatus();

        _logger.LogWritten += OnLogWritten;
        _controller.StateChanged += _stateChangedHandler;

        ApplySettingsToUi(_controller.CurrentSettings);
        UpdateStatus();
        LoadLogs();
    }

    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        if (e.CloseReason == CloseReason.UserClosing)
        {
            e.Cancel = true;
            Hide();
        }

        base.OnFormClosing(e);
    }

    protected override void OnResize(EventArgs e)
    {
        if (WindowState == FormWindowState.Minimized && Visible)
        {
            Hide();
        }

        base.OnResize(e);
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            _logger.LogWritten -= OnLogWritten;
            _controller.StateChanged -= _stateChangedHandler;
            _appIcon.Dispose();
        }

        base.Dispose(disposing);
    }

    private void SaveSettings()
    {
        var settings = _settingsStore.Load();
        settings.PolicyMode = _policyModeComboBox.SelectedIndex == 1
            ? PowerPolicyMode.KeepAwakeIndefinitely
            : PowerPolicyMode.FollowPowerPlan;
        settings.ResumeProtectionEnabled = _resumeProtectionCheckbox.Checked;
        settings.ResumeProtectionMode = _resumeActionComboBox.SelectedIndex == 1
            ? ResumeProtectionMode.Hibernate
            : ResumeProtectionMode.Sleep;
        settings.ResumeProtectionOnlyForUnattendedWake = _onlyUnattendedWakeCheckbox.Checked;
        settings.DisableWakeTimers = _disableWakeTimersCheckbox.Checked;
        settings.ResumeProtectionDelaySeconds = (int)_resumeDelayInput.Value;
        settings.StartMinimized = _startMinimizedCheckbox.Checked;
        settings.StartWithWindows = _autostartCheckbox.Checked;

        _controller.UpdateSettings(settings);
        ApplySettingsToUi(settings);
        UpdateStatus();
        LoadLogs();
    }

    private void ApplySettingsToUi(AppSettings settings)
    {
        _policyModeComboBox.SelectedIndex = settings.PolicyMode == PowerPolicyMode.KeepAwakeIndefinitely ? 1 : 0;
        _resumeProtectionCheckbox.Checked = settings.ResumeProtectionEnabled;
        _resumeActionComboBox.SelectedIndex = settings.ResumeProtectionMode == ResumeProtectionMode.Hibernate ? 1 : 0;
        _onlyUnattendedWakeCheckbox.Checked = settings.ResumeProtectionOnlyForUnattendedWake;
        _disableWakeTimersCheckbox.Checked = settings.DisableWakeTimers;
        _resumeDelayInput.Value = Math.Clamp(settings.ResumeProtectionDelaySeconds, 3, 600);
        _startMinimizedCheckbox.Checked = settings.StartMinimized;
        _autostartCheckbox.Checked = settings.StartWithWindows;
    }

    private void UpdateStatus()
    {
        if (IsDisposed)
        {
            return;
        }

        if (InvokeRequired)
        {
            BeginInvoke(new Action(UpdateStatus));
            return;
        }

        var settings = _controller.CurrentSettings;
        _statusLabel.Text =
            $"当前状态：{_controller.CurrentStatus}{Environment.NewLine}" +
            $"保护规则：{_controller.CurrentProtectionRuleSummary}{Environment.NewLine}" +
            $"最近一次唤醒判定：{settings.LastWakeSummary}{Environment.NewLine}" +
            $"唤醒定时器策略：{settings.WakeTimerPolicySummary}{Environment.NewLine}" +
            $"配置文件：{_settingsStore.SettingsPath}{Environment.NewLine}" +
            $"日志目录：{_logger.LogDirectory}{Environment.NewLine}" +
            $"上次挂起：{FormatTime(settings.LastSuspendUtc)} | 上次恢复：{FormatTime(settings.LastResumeUtc)}";
    }

    private void LoadLogs()
    {
        _logTextBox.Lines = _logger.ReadRecent().ToArray();
        _logTextBox.SelectionStart = _logTextBox.TextLength;
        _logTextBox.ScrollToCaret();
    }

    private void OnLogWritten(object? sender, string line)
    {
        if (IsDisposed)
        {
            return;
        }

        if (InvokeRequired)
        {
            BeginInvoke(new Action(() => OnLogWritten(sender, line)));
            return;
        }

        _logTextBox.AppendText(line + Environment.NewLine);
    }

    private void ExportDiagnosticReport()
    {
        try
        {
            var path = _diagnosticReportService.Export();
            UpdateStatus();
            LoadLogs();
            MessageBox.Show($"诊断报告已导出：\n{path}", "SleepSentinel", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            _logger.Error($"导出诊断报告失败：{ex.Message}");
            MessageBox.Show($"导出失败：{ex.Message}", "SleepSentinel", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private static void AddSettingRow(TableLayoutPanel panel, int rowIndex, string labelText, Control control)
    {
        panel.RowCount = Math.Max(panel.RowCount, rowIndex + 1);
        panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        panel.Controls.Add(new Label
        {
            Text = labelText,
            AutoSize = true,
            Padding = new Padding(0, 7, 12, 7)
        }, 0, rowIndex);
        panel.Controls.Add(control, 1, rowIndex);
    }

    private static string FormatTime(DateTimeOffset? value)
    {
        return value?.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss") ?? "无";
    }
}
