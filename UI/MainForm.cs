using SleepSentinel.Models;
using SleepSentinel.Services;

namespace SleepSentinel.UI;

public sealed class MainForm : Form
{
    private const int MinimumWindowWidth = 1100;
    private const int MinimumWindowHeight = 760;
    private const int PreferredWindowWidth = 1800;
    private const int PreferredWindowHeight = 1180;
    private readonly PowerController _controller;
    private readonly FileLogger _logger;
    private readonly SettingsStore _settingsStore;
    private readonly DiagnosticReportService _diagnosticReportService;
    private readonly ComboBox _policyModeComboBox;
    private readonly CheckBox _resumeProtectionCheckbox;
    private readonly RadioButton _resumeSleepRadioButton;
    private readonly RadioButton _resumeHibernateRadioButton;
    private readonly RadioButton _resumeLockScreenRadioButton;
    private readonly NumericUpDown _resumeDelayInput;
    private readonly CheckBox _onlyUnattendedWakeCheckbox;
    private readonly CheckBox _disableWakeTimersCheckbox;
    private readonly CheckBox _disableStandbyConnectivityCheckbox;
    private readonly CheckBox _disableWiFiDirectAdaptersCheckbox;
    private readonly CheckBox _enforceBatteryStandbyHibernateCheckbox;
    private readonly NumericUpDown _batteryStandbyHibernateTimeoutInput;
    private readonly CheckBox _blockKnownRemoteWakeCheckbox;
    private readonly TextBox _customRemoteWakeTextBox;
    private readonly CheckBox _startMinimizedCheckbox;
    private readonly CheckBox _autostartCheckbox;
    private readonly Label _statusLabel;
    private readonly Label _wakeTimerQuickStateLabel;
    private readonly Label _standbyConnectivityQuickStateLabel;
    private readonly Label _wifiDirectQuickStateLabel;
    private readonly Label _batteryStandbyHibernateQuickStateLabel;
    private readonly Label _remoteWakeQuickStateLabel;
    private readonly Label _customRemoteWakeHintLabel;
    private readonly TextBox _detailsTextBox;
    private readonly TextBox _diagnosticsTextBox;
    private readonly TextBox _logTextBox;
    private readonly EventHandler _stateChangedHandler;
    private readonly Icon _appIcon;
    private bool _suppressInteractiveToggleEvents;

    public MainForm(PowerController controller, FileLogger logger, SettingsStore settingsStore, Icon appIcon)
    {
        _controller = controller;
        _logger = logger;
        _settingsStore = settingsStore;
        _diagnosticReportService = new DiagnosticReportService(settingsStore, logger, controller);
        _appIcon = (Icon)appIcon.Clone();

        Text = "SleepSentinel";
        MinimumSize = new Size(MinimumWindowWidth, MinimumWindowHeight);
        StartPosition = FormStartPosition.Manual;
        Icon = _appIcon;
        ApplyInitialWindowBounds(_controller.CurrentSettings);

        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 5,
            Padding = new Padding(16)
        };
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

        var intro = new Label
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            Text = "让电脑在“该睡就睡”时不要被其他软件长期唤醒，同时保留类似 PowerToys Awake 的无限保持唤醒模式。设置现在会自动生效。"
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

        _policyModeComboBox = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 280 };
        _policyModeComboBox.Items.AddRange(
        [
            "遵循电源计划",
            "无限保持唤醒（类似 PowerToys Awake）"
        ]);
        _policyModeComboBox.SelectedIndexChanged += (_, _) => ApplyUiSettingsImmediately();

        _resumeProtectionCheckbox = new CheckBox
        {
            AutoSize = true,
            Text = "恢复后自动执行睡眠/休眠/锁屏，避免被软件唤醒后常驻亮机"
        };
        _resumeProtectionCheckbox.CheckedChanged += (_, _) => ApplyUiSettingsImmediately();

        _resumeSleepRadioButton = CreateResumeActionRadioButton("睡眠");
        _resumeHibernateRadioButton = CreateResumeActionRadioButton("休眠");
        _resumeLockScreenRadioButton = CreateResumeActionRadioButton("锁屏");

        _resumeDelayInput = new NumericUpDown
        {
            Minimum = 3,
            Maximum = 600,
            Width = 100
        };
        _resumeDelayInput.ValueChanged += (_, _) => ApplyUiSettingsImmediately();

        _onlyUnattendedWakeCheckbox = new CheckBox
        {
            AutoSize = true,
            Text = "仅在人工行为时跳过；人工包括键盘、鼠标、开盖、解锁、控制台/远程接管、登录"
        };
        _onlyUnattendedWakeCheckbox.CheckedChanged += (_, _) => ApplyUiSettingsImmediately();

        _disableWakeTimersCheckbox = new CheckBox
        {
            AutoSize = true,
            Text = "关闭当前电源计划的唤醒定时器（AC/DC）"
        };
        _disableWakeTimersCheckbox.CheckedChanged += (_, _) =>
        {
            if (_suppressInteractiveToggleEvents)
            {
                return;
            }

            if (_disableWakeTimersCheckbox.Checked)
            {
                _controller.BlockSoftwareWake();
            }
            else
            {
                _controller.RestoreSoftwareWake();
            }

            SyncUiFromController();
        };

        _disableStandbyConnectivityCheckbox = new CheckBox
        {
            AutoSize = true,
            Text = "关闭待机状态下的网络连接（AC/DC，减少 Windows Update/更新协调器在合盖待机时拉活）"
        };
        _disableStandbyConnectivityCheckbox.CheckedChanged += (_, _) =>
        {
            if (_suppressInteractiveToggleEvents)
            {
                return;
            }

            if (_disableStandbyConnectivityCheckbox.Checked)
            {
                _controller.BlockStandbyConnectivityWake();
            }
            else
            {
                _controller.RestoreStandbyConnectivityWake();
            }

            SyncUiFromController();
        };

        _disableWiFiDirectAdaptersCheckbox = new CheckBox
        {
            AutoSize = true,
            Text = "禁用 Microsoft Wi-Fi Direct 虚拟适配器（降低 S0 待机恢复异常；影响无线投屏/移动热点/附近共享）"
        };
        _disableWiFiDirectAdaptersCheckbox.CheckedChanged += (_, _) =>
        {
            if (_suppressInteractiveToggleEvents)
            {
                return;
            }

            if (_disableWiFiDirectAdaptersCheckbox.Checked)
            {
                _controller.DisableWiFiDirectAdapters();
            }
            else
            {
                _controller.RestoreWiFiDirectAdapters();
            }

            SyncUiFromController();
        };

        _batteryStandbyHibernateTimeoutInput = new NumericUpDown
        {
            Minimum = 3,
            Maximum = 240,
            Width = 90
        };
        _batteryStandbyHibernateTimeoutInput.ValueChanged += (_, _) =>
        {
            UpdateBatteryStandbyHibernateCheckboxText();
            ApplyUiSettingsImmediately();
        };

        _enforceBatteryStandbyHibernateCheckbox = new CheckBox
        {
            AutoSize = true
        };
        _enforceBatteryStandbyHibernateCheckbox.CheckedChanged += (_, _) =>
        {
            if (_suppressInteractiveToggleEvents)
            {
                return;
            }

            if (_enforceBatteryStandbyHibernateCheckbox.Checked)
            {
                _controller.EnableBatteryStandbyHibernateFallback();
            }
            else
            {
                _controller.RestoreBatteryStandbyHibernateFallback();
            }

            SyncUiFromController();
        };

        _blockKnownRemoteWakeCheckbox = new CheckBox
        {
            AutoSize = true,
            Text = "拦截常见远程软件的保持唤醒请求（ToDesk、向日葵、GameViewer/UU、AnyDesk、TeamViewer、RustDesk）"
        };
        _blockKnownRemoteWakeCheckbox.CheckedChanged += (_, _) =>
        {
            if (_suppressInteractiveToggleEvents)
            {
                return;
            }

            if (_blockKnownRemoteWakeCheckbox.Checked)
            {
                _controller.BlockKnownRemoteWakeRequests();
            }
            else
            {
                _controller.RestoreKnownRemoteWakeRequests();
            }

            SyncUiFromController();
        };

        _customRemoteWakeTextBox = new TextBox
        {
            Multiline = true,
            Width = 460,
            Height = 54,
            ScrollBars = ScrollBars.Vertical,
            AcceptsReturn = true
        };
        _customRemoteWakeTextBox.Leave += (_, _) => ApplyCustomRemoteEntriesFromUi();

        _startMinimizedCheckbox = new CheckBox { AutoSize = true, Text = "启动后仅驻留托盘" };
        _startMinimizedCheckbox.CheckedChanged += (_, _) => ApplyUiSettingsImmediately();
        _autostartCheckbox = new CheckBox { AutoSize = true, Text = "开机自启" };
        _autostartCheckbox.CheckedChanged += (_, _) => ApplyUiSettingsImmediately();

        AddSettingRow(settingsPanel, 0, "工作模式", _policyModeComboBox);
        AddSettingRow(settingsPanel, 1, "恢复保护", _resumeProtectionCheckbox);

        var resumeOptionsFlow = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Left, WrapContents = false };
        var resumeActionFlow = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = new Padding(0) };
        resumeActionFlow.Controls.Add(_resumeSleepRadioButton);
        resumeActionFlow.Controls.Add(_resumeHibernateRadioButton);
        resumeActionFlow.Controls.Add(_resumeLockScreenRadioButton);
        resumeOptionsFlow.Controls.Add(new Label { Text = "恢复后执行", AutoSize = true, Padding = new Padding(0, 7, 8, 0) });
        resumeOptionsFlow.Controls.Add(resumeActionFlow);
        resumeOptionsFlow.Controls.Add(new Label { Text = "延迟秒数", AutoSize = true, Padding = new Padding(16, 7, 8, 0) });
        resumeOptionsFlow.Controls.Add(_resumeDelayInput);
        AddSettingRow(settingsPanel, 2, "保护细节", resumeOptionsFlow);
        AddSettingRow(settingsPanel, 3, "唤醒过滤", _onlyUnattendedWakeCheckbox);

        var wakeTimerActions = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Left, WrapContents = false };
        var reapplyWakeTimerButton = new Button { Text = "重新应用", AutoSize = true };
        reapplyWakeTimerButton.Click += (_, _) =>
        {
            _controller.ReapplyWakeTimerPolicy();
            SyncUiFromController();
        };
        _wakeTimerQuickStateLabel = new Label { AutoSize = true, Padding = new Padding(12, 7, 0, 0) };
        wakeTimerActions.Controls.Add(_disableWakeTimersCheckbox);
        wakeTimerActions.Controls.Add(reapplyWakeTimerButton);
        wakeTimerActions.Controls.Add(_wakeTimerQuickStateLabel);
        AddSettingRow(settingsPanel, 4, "软件唤醒", wakeTimerActions);

        var standbyConnectivityActions = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Left, WrapContents = false };
        var reapplyStandbyConnectivityButton = new Button { Text = "重新应用", AutoSize = true };
        reapplyStandbyConnectivityButton.Click += (_, _) =>
        {
            _controller.ReapplyStandbyConnectivityPolicy();
            SyncUiFromController();
        };
        _standbyConnectivityQuickStateLabel = new Label { AutoSize = true, Padding = new Padding(12, 7, 0, 0) };
        standbyConnectivityActions.Controls.Add(_disableStandbyConnectivityCheckbox);
        standbyConnectivityActions.Controls.Add(reapplyStandbyConnectivityButton);
        standbyConnectivityActions.Controls.Add(_standbyConnectivityQuickStateLabel);
        AddSettingRow(settingsPanel, 5, "待机联网", standbyConnectivityActions);

        var wifiDirectActions = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Left, WrapContents = false };
        var reapplyWiFiDirectButton = new Button { Text = "重新应用", AutoSize = true };
        reapplyWiFiDirectButton.Click += (_, _) =>
        {
            _controller.ReapplyWiFiDirectAdapterPolicy();
            SyncUiFromController();
        };
        _wifiDirectQuickStateLabel = new Label { AutoSize = true, Padding = new Padding(12, 7, 0, 0) };
        wifiDirectActions.Controls.Add(_disableWiFiDirectAdaptersCheckbox);
        wifiDirectActions.Controls.Add(reapplyWiFiDirectButton);
        wifiDirectActions.Controls.Add(_wifiDirectQuickStateLabel);
        AddSettingRow(settingsPanel, 6, "无线稳态", wifiDirectActions);

        var batteryStandbyHibernateActions = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Left, WrapContents = false };
        var reapplyBatteryStandbyHibernateButton = new Button { Text = "重新应用", AutoSize = true };
        reapplyBatteryStandbyHibernateButton.Click += (_, _) =>
        {
            ApplyUiSettingsImmediately();
            _controller.ReapplyBatteryStandbyHibernatePolicy();
            SyncUiFromController();
        };
        _batteryStandbyHibernateQuickStateLabel = new Label { AutoSize = true, Padding = new Padding(12, 7, 0, 0) };
        batteryStandbyHibernateActions.Controls.Add(_enforceBatteryStandbyHibernateCheckbox);
        batteryStandbyHibernateActions.Controls.Add(new Label { Text = "分钟数", AutoSize = true, Padding = new Padding(16, 7, 8, 0) });
        batteryStandbyHibernateActions.Controls.Add(_batteryStandbyHibernateTimeoutInput);
        batteryStandbyHibernateActions.Controls.Add(reapplyBatteryStandbyHibernateButton);
        batteryStandbyHibernateActions.Controls.Add(_batteryStandbyHibernateQuickStateLabel);
        AddSettingRow(settingsPanel, 7, "电池兜底", batteryStandbyHibernateActions);

        var remoteWakeActions = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Left, WrapContents = false };
        var reapplyRemoteWakeButton = new Button { Text = "重新应用", AutoSize = true };
        reapplyRemoteWakeButton.Click += (_, _) =>
        {
            ApplyCustomRemoteEntriesFromUi();
            _controller.ReapplyKnownRemoteWakePolicy();
            SyncUiFromController();
        };
        var suggestRemoteWakeButton = new Button { Text = "自动建议", AutoSize = true };
        suggestRemoteWakeButton.Click += (_, _) => SuggestCustomRemoteWakeEntries();
        _remoteWakeQuickStateLabel = new Label { AutoSize = true, Padding = new Padding(12, 7, 0, 0) };
        remoteWakeActions.Controls.Add(_blockKnownRemoteWakeCheckbox);
        remoteWakeActions.Controls.Add(reapplyRemoteWakeButton);
        remoteWakeActions.Controls.Add(suggestRemoteWakeButton);
        remoteWakeActions.Controls.Add(_remoteWakeQuickStateLabel);
        AddSettingRow(settingsPanel, 8, "远控拦截", remoteWakeActions);

        var customRemoteWakeActions = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Left, FlowDirection = FlowDirection.TopDown };
        var customRemoteWakeToolbar = new FlowLayoutPanel { AutoSize = true, WrapContents = false };
        var applyCustomRemoteWakeButton = new Button { Text = "应用名单", AutoSize = true };
        applyCustomRemoteWakeButton.Click += (_, _) => ApplyCustomRemoteEntriesFromUi();
        _customRemoteWakeHintLabel = new Label { AutoSize = true, Padding = new Padding(12, 7, 0, 0) };
        customRemoteWakeToolbar.Controls.Add(applyCustomRemoteWakeButton);
        customRemoteWakeToolbar.Controls.Add(_customRemoteWakeHintLabel);
        customRemoteWakeActions.Controls.Add(_customRemoteWakeTextBox);
        customRemoteWakeActions.Controls.Add(customRemoteWakeToolbar);
        AddSettingRow(settingsPanel, 9, "自定义远控", customRemoteWakeActions);

        var startupFlow = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Left, WrapContents = false };
        startupFlow.Controls.Add(_startMinimizedCheckbox);
        startupFlow.Controls.Add(_autostartCheckbox);
        AddSettingRow(settingsPanel, 10, "启动行为", startupFlow);

        root.Controls.Add(settingsPanel);

        var actionsFlow = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            Padding = new Padding(0, 0, 0, 12)
        };

        var reapplyAllButton = new Button { Text = "重新应用全部设置", AutoSize = true };
        reapplyAllButton.Click += (_, _) =>
        {
            ApplyUiSettingsImmediately();
            _controller.ReapplyAllManagedSettings();
            SyncUiFromController(includeDiagnostics: false);
        };
        var sleepButton = new Button { Text = "立即睡眠", AutoSize = true };
        sleepButton.Click += (_, _) => _controller.SleepNow();
        var hibernateButton = new Button { Text = "立即休眠", AutoSize = true };
        hibernateButton.Click += (_, _) => _controller.HibernateNow();
        var refreshLogButton = new Button { Text = "刷新日志", AutoSize = true };
        refreshLogButton.Click += (_, _) => LoadLogs();
        var diagnosticsButton = new Button { Text = "记录唤醒诊断", AutoSize = true };
        diagnosticsButton.Click += (_, _) => RefreshDiagnostics(logSnapshot: true);
        var exportReportButton = new Button { Text = "导出诊断报告", AutoSize = true };
        exportReportButton.Click += (_, _) => ExportDiagnosticReport();

        actionsFlow.Controls.Add(reapplyAllButton);
        actionsFlow.Controls.Add(sleepButton);
        actionsFlow.Controls.Add(hibernateButton);
        actionsFlow.Controls.Add(refreshLogButton);
        actionsFlow.Controls.Add(diagnosticsButton);
        actionsFlow.Controls.Add(exportReportButton);
        root.Controls.Add(actionsFlow);

        _statusLabel = new Label
        {
            AutoSize = true,
            Dock = DockStyle.Top,
            Padding = new Padding(0, 0, 0, 12)
        };
        root.Controls.Add(_statusLabel);

        var tabs = new TabControl
        {
            Dock = DockStyle.Fill
        };

        _detailsTextBox = CreateReadOnlyOutputTextBox();
        _diagnosticsTextBox = CreateReadOnlyOutputTextBox();
        _logTextBox = CreateReadOnlyOutputTextBox();
        _logTextBox.Font = new Font(FontFamily.GenericMonospace, 9.5f);

        tabs.TabPages.Add(CreateTabPage("状态详情", _detailsTextBox));
        tabs.TabPages.Add(CreateTabPage("诊断摘要", _diagnosticsTextBox));
        tabs.TabPages.Add(CreateTabPage("运行日志", _logTextBox));
        root.Controls.Add(tabs);

        Controls.Add(root);

        _stateChangedHandler = (_, _) => SyncUiFromController(includeDiagnostics: false);

        _logger.LogWritten += OnLogWritten;
        _controller.StateChanged += _stateChangedHandler;

        ApplySettingsToUi(_controller.CurrentSettings);
        RefreshStatusAndDetails();
        _diagnosticsTextBox.Text = "最近诊断尚未刷新。点击“记录唤醒诊断”后，这里会显示 powercfg、事件日志和 SleepStudy 摘要。";
        LoadLogs();
    }

    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        SaveWindowBounds();

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
            SaveWindowBounds();
            Hide();
        }

        base.OnResize(e);
    }

    protected override void OnResizeEnd(EventArgs e)
    {
        SaveWindowBounds();
        base.OnResizeEnd(e);
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

    private void ApplyUiSettingsImmediately()
    {
        if (_suppressInteractiveToggleEvents)
        {
            return;
        }

        var settings = CreateSettingsFromUi();
        _controller.UpdateSettings(settings);
        SyncUiFromController(includeDiagnostics: false);
    }

    private void ApplyCustomRemoteEntriesFromUi()
    {
        if (_suppressInteractiveToggleEvents)
        {
            return;
        }

        _controller.UpdateCustomRemoteWakeEntries(ParseCustomRemoteWakeEntries(_customRemoteWakeTextBox.Text));
        SyncUiFromController(includeDiagnostics: false);
    }

    private void SuggestCustomRemoteWakeEntries()
    {
        var suggestions = _controller.SuggestCustomRemoteWakeEntries();
        if (suggestions.Count == 0)
        {
            MessageBox.Show("当前 requests / 运行进程 / 运行服务里没有发现新的远控候选项。", "SleepSentinel", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        var mergedEntries = ParseCustomRemoteWakeEntries(_customRemoteWakeTextBox.Text)
            .Concat(suggestions)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(static entry => entry, StringComparer.OrdinalIgnoreCase)
            .ToArray();

        _customRemoteWakeTextBox.Text = string.Join(Environment.NewLine, mergedEntries);
        ApplyCustomRemoteEntriesFromUi();
        MessageBox.Show($"已追加 {suggestions.Count} 条候选项到自定义远控名单。", "SleepSentinel", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    private AppSettings CreateSettingsFromUi()
    {
        var settings = _controller.CurrentSettings;
        settings.PolicyMode = _policyModeComboBox.SelectedIndex == 1
            ? PowerPolicyMode.KeepAwakeIndefinitely
            : PowerPolicyMode.FollowPowerPlan;
        settings.ResumeProtectionEnabled = _resumeProtectionCheckbox.Checked;
        settings.ResumeProtectionMode = GetResumeProtectionModeFromUi();
        settings.ResumeProtectionOnlyForUnattendedWake = _onlyUnattendedWakeCheckbox.Checked;
        settings.ResumeProtectionDelaySeconds = (int)_resumeDelayInput.Value;
        settings.DisableWakeTimers = _disableWakeTimersCheckbox.Checked;
        settings.DisableStandbyConnectivity = _disableStandbyConnectivityCheckbox.Checked;
        settings.DisableWiFiDirectAdapters = _disableWiFiDirectAdaptersCheckbox.Checked;
        settings.EnforceBatteryStandbyHibernate = _enforceBatteryStandbyHibernateCheckbox.Checked;
        settings.BatteryStandbyHibernateTimeoutSeconds = checked((int)_batteryStandbyHibernateTimeoutInput.Value * 60);
        settings.BlockKnownRemoteWakeRequests = _blockKnownRemoteWakeCheckbox.Checked;
        settings.CustomRemoteWakeEntries = ParseCustomRemoteWakeEntries(_customRemoteWakeTextBox.Text).ToList();
        settings.StartMinimized = _startMinimizedCheckbox.Checked;
        settings.StartWithWindows = _autostartCheckbox.Checked;
        return settings;
    }

    private void ApplySettingsToUi(AppSettings settings)
    {
        _suppressInteractiveToggleEvents = true;
        try
        {
            _policyModeComboBox.SelectedIndex = settings.PolicyMode == PowerPolicyMode.KeepAwakeIndefinitely ? 1 : 0;
            _resumeProtectionCheckbox.Checked = settings.ResumeProtectionEnabled;
            SetResumeProtectionModeOnUi(settings.ResumeProtectionMode);
            _onlyUnattendedWakeCheckbox.Checked = settings.ResumeProtectionOnlyForUnattendedWake;
            _resumeDelayInput.Value = Math.Clamp(settings.ResumeProtectionDelaySeconds, 3, 600);
            _disableWakeTimersCheckbox.Checked = settings.DisableWakeTimers;
            _disableStandbyConnectivityCheckbox.Checked = settings.DisableStandbyConnectivity;
            _disableWiFiDirectAdaptersCheckbox.Checked = settings.DisableWiFiDirectAdapters;
            _batteryStandbyHibernateTimeoutInput.Value = Math.Clamp(settings.BatteryStandbyHibernateTimeoutSeconds / 60, 3, 240);
            _enforceBatteryStandbyHibernateCheckbox.Checked = settings.EnforceBatteryStandbyHibernate;
            _blockKnownRemoteWakeCheckbox.Checked = settings.BlockKnownRemoteWakeRequests;
            _customRemoteWakeTextBox.Text = string.Join(Environment.NewLine, settings.CustomRemoteWakeEntries);
            _startMinimizedCheckbox.Checked = settings.StartMinimized;
            _autostartCheckbox.Checked = settings.StartWithWindows;
        }
        finally
        {
            _suppressInteractiveToggleEvents = false;
        }

        UpdateBatteryStandbyHibernateCheckboxText();
        _wakeTimerQuickStateLabel.Text = _controller.CurrentWakeTimerQuickState;
        _standbyConnectivityQuickStateLabel.Text = _controller.CurrentStandbyConnectivityQuickState;
        _wifiDirectQuickStateLabel.Text = _controller.CurrentWiFiDirectQuickState;
        _batteryStandbyHibernateQuickStateLabel.Text = _controller.CurrentBatteryStandbyHibernateQuickState;
        _remoteWakeQuickStateLabel.Text = _controller.CurrentRemoteWakeQuickState;
        _customRemoteWakeHintLabel.Text = settings.CustomRemoteWakeEntries.Count == 0
            ? "当前：未添加自定义条目"
            : $"当前：{settings.CustomRemoteWakeEntries.Count} 条自定义规则";
    }

    private RadioButton CreateResumeActionRadioButton(string text)
    {
        var radioButton = new RadioButton
        {
            AutoSize = true,
            Text = text,
            Margin = new Padding(0, 6, 16, 0)
        };
        radioButton.CheckedChanged += (_, _) =>
        {
            if (_suppressInteractiveToggleEvents || !radioButton.Checked)
            {
                return;
            }

            ApplyUiSettingsImmediately();
        };

        return radioButton;
    }

    private ResumeProtectionMode GetResumeProtectionModeFromUi()
    {
        if (_resumeLockScreenRadioButton.Checked)
        {
            return ResumeProtectionMode.LockScreen;
        }

        return _resumeHibernateRadioButton.Checked
            ? ResumeProtectionMode.Hibernate
            : ResumeProtectionMode.Sleep;
    }

    private void SetResumeProtectionModeOnUi(ResumeProtectionMode mode)
    {
        _resumeSleepRadioButton.Checked = mode == ResumeProtectionMode.Sleep;
        _resumeHibernateRadioButton.Checked = mode == ResumeProtectionMode.Hibernate;
        _resumeLockScreenRadioButton.Checked = mode == ResumeProtectionMode.LockScreen;
    }

    private void RefreshStatusAndDetails()
    {
        if (IsDisposed)
        {
            return;
        }

        if (InvokeRequired)
        {
            if (!IsHandleCreated)
            {
                return;
            }

            try
            {
                BeginInvoke(new Action(RefreshStatusAndDetails));
            }
            catch (InvalidOperationException)
            {
            }

            return;
        }

        var settings = _controller.CurrentSettings;
        _statusLabel.Text =
            $"当前状态：{_controller.CurrentStatus}{Environment.NewLine}" +
            $"风险提示：{_controller.CurrentRiskSummary}{Environment.NewLine}" +
            $"权限状态：{_controller.CurrentCapabilitySummary}{Environment.NewLine}" +
            $"最近一次唤醒：{settings.LastWakeSummary}{Environment.NewLine}" +
            $"最近证据：{settings.LastWakeEvidenceSummary}";

        _detailsTextBox.Text =
            $"保护规则：{_controller.CurrentProtectionRuleSummary}{Environment.NewLine}" +
            $"电源计划：{_controller.CurrentPowerPlanSummary}{Environment.NewLine}" +
            $"远控名单：{_controller.CurrentManagedRemoteEntriesSummary}{Environment.NewLine}" +
            $"唤醒定时器策略：{settings.WakeTimerPolicySummary}{Environment.NewLine}" +
            $"待机联网策略：{settings.StandbyConnectivityPolicySummary}{Environment.NewLine}" +
            $"Wi-Fi Direct 策略：{settings.WiFiDirectAdapterPolicySummary}{Environment.NewLine}" +
            $"电池兜底策略：{settings.BatteryStandbyHibernatePolicySummary}{Environment.NewLine}" +
            $"远控拦截策略：{settings.KnownRemoteWakePolicySummary}{Environment.NewLine}" +
            $"启动策略：{settings.AutostartPolicySummary}{Environment.NewLine}" +
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
            if (!IsHandleCreated)
            {
                return;
            }

            try
            {
                BeginInvoke(new Action(() => OnLogWritten(sender, line)));
            }
            catch (InvalidOperationException)
            {
            }

            return;
        }

        _logTextBox.AppendText(line + Environment.NewLine);
    }

    private void RefreshDiagnostics(bool logSnapshot)
    {
        try
        {
            var snapshot = _controller.CollectWakeDiagnosticSnapshot(includePowerRequests: true, includeSleepStudy: true);
            var text = _controller.FormatWakeDiagnosticSnapshot(snapshot, includePowerRequests: true, includeSleepStudy: true);
            _diagnosticsTextBox.Text = text;

            if (logSnapshot)
            {
                _logger.Warn("用户手动收集唤醒诊断。");
                _logger.Warn(text);
            }

            RefreshStatusAndDetails();
            LoadLogs();
        }
        catch (Exception ex)
        {
            _logger.Error($"刷新诊断失败：{ex.Message}");
            MessageBox.Show($"刷新诊断失败：{ex.Message}", "SleepSentinel", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void ExportDiagnosticReport()
    {
        try
        {
            var path = _diagnosticReportService.Export();
            RefreshStatusAndDetails();
            LoadLogs();
            MessageBox.Show($"诊断报告已导出：\n{path}", "SleepSentinel", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            _logger.Error($"导出诊断报告失败：{ex.Message}");
            MessageBox.Show($"导出失败：{ex.Message}", "SleepSentinel", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void SyncUiFromController(bool includeDiagnostics = false)
    {
        if (IsDisposed)
        {
            return;
        }

        if (InvokeRequired)
        {
            if (!IsHandleCreated)
            {
                return;
            }

            try
            {
                BeginInvoke(new Action(() => SyncUiFromController(includeDiagnostics)));
            }
            catch (InvalidOperationException)
            {
            }

            return;
        }

        ApplySettingsToUi(_controller.CurrentSettings);
        RefreshStatusAndDetails();
        LoadLogs();
        if (includeDiagnostics)
        {
            RefreshDiagnostics(logSnapshot: false);
        }
    }

    private void UpdateBatteryStandbyHibernateCheckboxText()
    {
        _enforceBatteryStandbyHibernateCheckbox.Text =
            $"电池供电下在待机 {_batteryStandbyHibernateTimeoutInput.Value} 分钟后自动转入休眠（仅 DC，防止合盖后一夜耗尽）";
    }

    private void ApplyInitialWindowBounds(AppSettings settings)
    {
        var workingArea = Screen.FromPoint(Cursor.Position).WorkingArea;
        var defaultWidth = Math.Min(Math.Max(MinimumWindowWidth, (int)(workingArea.Width * 0.86)), Math.Min(PreferredWindowWidth, workingArea.Width));
        var defaultHeight = Math.Min(Math.Max(MinimumWindowHeight, (int)(workingArea.Height * 0.88)), Math.Min(PreferredWindowHeight, workingArea.Height));

        if (settings.WindowBoundsCaptured
            && settings.WindowWidth >= MinimumWindowWidth
            && settings.WindowHeight >= MinimumWindowHeight)
        {
            var storedBounds = new Rectangle(settings.WindowX, settings.WindowY, settings.WindowWidth, settings.WindowHeight);
            Bounds = FitBoundsToWorkingArea(storedBounds, workingArea);
            return;
        }

        Size = new Size(defaultWidth, defaultHeight);
        Location = new Point(
            workingArea.Left + Math.Max(0, (workingArea.Width - defaultWidth) / 2),
            workingArea.Top + Math.Max(0, (workingArea.Height - defaultHeight) / 2));
    }

    private void SaveWindowBounds()
    {
        if (WindowState == FormWindowState.Minimized)
        {
            return;
        }

        var bounds = WindowState == FormWindowState.Normal ? Bounds : RestoreBounds;
        if (bounds.Width < MinimumWindowWidth || bounds.Height < MinimumWindowHeight)
        {
            return;
        }

        _settingsStore.Update(settings =>
        {
            settings.WindowBoundsCaptured = true;
            settings.WindowWidth = bounds.Width;
            settings.WindowHeight = bounds.Height;
            settings.WindowX = bounds.X;
            settings.WindowY = bounds.Y;
        });
    }

    private static IReadOnlyList<string> ParseCustomRemoteWakeEntries(string text)
    {
        return text
            .Split(["\r\n", "\n"], StringSplitOptions.RemoveEmptyEntries)
            .Select(static line => line.Trim())
            .Where(static line => !string.IsNullOrWhiteSpace(line))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
    }

    private static TabPage CreateTabPage(string title, Control content)
    {
        var page = new TabPage(title);
        page.Controls.Add(content);
        return page;
    }

    private static TextBox CreateReadOnlyOutputTextBox()
    {
        return new TextBox
        {
            Multiline = true,
            ReadOnly = true,
            ScrollBars = ScrollBars.Vertical,
            Dock = DockStyle.Fill
        };
    }

    private static Rectangle FitBoundsToWorkingArea(Rectangle bounds, Rectangle workingArea)
    {
        var width = Math.Min(Math.Max(bounds.Width, MinimumWindowWidth), workingArea.Width);
        var height = Math.Min(Math.Max(bounds.Height, MinimumWindowHeight), workingArea.Height);
        var x = Math.Min(Math.Max(bounds.X, workingArea.Left), workingArea.Right - width);
        var y = Math.Min(Math.Max(bounds.Y, workingArea.Top), workingArea.Bottom - height);
        return new Rectangle(x, y, width, height);
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
