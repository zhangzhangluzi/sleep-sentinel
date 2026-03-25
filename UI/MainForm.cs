using SleepSentinel.Models;
using SleepSentinel.Services;

namespace SleepSentinel.UI;

public sealed class MainForm : Form
{
    private static readonly Color ShellColor = Color.FromArgb(244, 247, 248);
    private static readonly Color SidebarColor = Color.FromArgb(234, 242, 237);
    private static readonly Color CardColor = Color.White;
    private static readonly Color BorderColor = Color.FromArgb(223, 229, 232);
    private static readonly Color TextColor = Color.FromArgb(43, 52, 55);
    private static readonly Color MutedTextColor = Color.FromArgb(88, 96, 100);
    private static readonly Color PrimaryColor = Color.FromArgb(27, 109, 36);
    private static readonly Color PrimarySoftColor = Color.FromArgb(220, 244, 225);
    private static readonly Color AccentColor = Color.FromArgb(77, 98, 108);
    private static readonly Color AccentSoftColor = Color.FromArgb(224, 233, 238);
    private static readonly Color WarningColor = Color.FromArgb(169, 57, 0);
    private static readonly Color WarningSoftColor = Color.FromArgb(252, 236, 227);

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
    private readonly CheckBox _disableStandbyConnectivityCheckbox;
    private readonly CheckBox _enforceBatteryStandbyHibernateCheckbox;
    private readonly CheckBox _blockKnownRemoteWakeCheckbox;
    private readonly CheckBox _startMinimizedCheckbox;
    private readonly CheckBox _autostartCheckbox;
    private readonly Label _overviewStatusChipLabel;
    private readonly Label _overviewModeValueLabel;
    private readonly Label _overviewWakeValueLabel;
    private readonly Label _overviewRuleValueLabel;
    private readonly Label _wakeTimerQuickStateLabel;
    private readonly Label _standbyConnectivityQuickStateLabel;
    private readonly Label _batteryStandbyHibernateQuickStateLabel;
    private readonly Label _remoteWakeQuickStateLabel;
    private readonly Label _wakeTimerSummaryLabel;
    private readonly Label _standbyConnectivitySummaryLabel;
    private readonly Label _batteryStandbyHibernateSummaryLabel;
    private readonly Label _remoteWakeSummaryLabel;
    private readonly TextBox _logTextBox;
    private readonly TextBox _wakeDiagnosticsTextBox;
    private readonly TextBox _powerRequestDiagnosticsTextBox;
    private readonly Dictionary<string, Label> _statusValueLabels = [];
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
        Width = 1120;
        Height = 780;
        MinimumSize = new Size(980, 700);
        StartPosition = FormStartPosition.CenterScreen;
        Icon = _appIcon;
        BackColor = ShellColor;
        Font = new Font("Segoe UI", 9F, FontStyle.Regular);

        _policyModeComboBox = new ComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDownList,
            Width = 250
        };
        _policyModeComboBox.Items.AddRange([
            "遵循电源计划",
            "无限保持唤醒（类似 PowerToys Awake）"
        ]);

        _resumeProtectionCheckbox = CreateSettingsCheckBox("恢复后自动重新进入睡眠/休眠，避免被软件唤醒后长期常驻亮机");
        _resumeActionComboBox = new ComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDownList,
            Width = 110
        };
        _resumeActionComboBox.Items.AddRange(["睡眠", "休眠"]);
        _resumeDelayInput = new NumericUpDown
        {
            Minimum = 3,
            Maximum = 600,
            Value = 8,
            Width = 90
        };
        _onlyUnattendedWakeCheckbox = CreateSettingsCheckBox("仅在人工行为时跳过；人工包括键盘、鼠标、开盖、解锁、控制台/远程接管、登录");

        _disableWakeTimersCheckbox = CreateSettingsCheckBox("关闭当前电源计划的唤醒定时器（AC/DC）");
        _disableWakeTimersCheckbox.CheckedChanged += (_, _) => OnPolicyToggleChanged(
            _disableWakeTimersCheckbox,
            _controller.BlockSoftwareWake,
            _controller.RestoreSoftwareWake);

        _disableStandbyConnectivityCheckbox = CreateSettingsCheckBox("关闭待机状态下的网络连接，减少 Windows Update / 更新协调器在合盖待机时拉活");
        _disableStandbyConnectivityCheckbox.CheckedChanged += (_, _) => OnPolicyToggleChanged(
            _disableStandbyConnectivityCheckbox,
            _controller.BlockStandbyConnectivityWake,
            _controller.RestoreStandbyConnectivityWake);

        _enforceBatteryStandbyHibernateCheckbox = CreateSettingsCheckBox("电池供电下在待机 10 分钟后自动转入休眠（仅 DC，防止合盖后一夜耗尽）");
        _enforceBatteryStandbyHibernateCheckbox.CheckedChanged += (_, _) => OnPolicyToggleChanged(
            _enforceBatteryStandbyHibernateCheckbox,
            _controller.EnableBatteryStandbyHibernateFallback,
            _controller.RestoreBatteryStandbyHibernateFallback);

        _blockKnownRemoteWakeCheckbox = CreateSettingsCheckBox("拦截常见远程软件的保持唤醒请求（ToDesk、向日葵、GameViewer/UU、AnyDesk、TeamViewer、RustDesk）");
        _blockKnownRemoteWakeCheckbox.CheckedChanged += (_, _) => OnPolicyToggleChanged(
            _blockKnownRemoteWakeCheckbox,
            _controller.BlockKnownRemoteWakeRequests,
            _controller.RestoreKnownRemoteWakeRequests);

        _startMinimizedCheckbox = CreateSettingsCheckBox("启动后仅驻留托盘");
        _autostartCheckbox = CreateSettingsCheckBox("开机自启");

        _overviewStatusChipLabel = CreateChipLabel();
        _overviewModeValueLabel = CreateValueLabel();
        _overviewWakeValueLabel = CreateValueLabel();
        _overviewRuleValueLabel = CreateBodyLabel();
        _wakeTimerQuickStateLabel = CreateChipLabel();
        _standbyConnectivityQuickStateLabel = CreateChipLabel();
        _batteryStandbyHibernateQuickStateLabel = CreateChipLabel();
        _remoteWakeQuickStateLabel = CreateChipLabel();
        _wakeTimerSummaryLabel = CreateBodyLabel();
        _standbyConnectivitySummaryLabel = CreateBodyLabel();
        _batteryStandbyHibernateSummaryLabel = CreateBodyLabel();
        _remoteWakeSummaryLabel = CreateBodyLabel();

        _logTextBox = new TextBox
        {
            Multiline = true,
            ReadOnly = true,
            ScrollBars = ScrollBars.Vertical,
            Dock = DockStyle.Fill,
            BorderStyle = BorderStyle.None,
            BackColor = Color.FromArgb(238, 243, 245),
            ForeColor = TextColor,
            Font = new Font(FontFamily.GenericMonospace, 9.2f),
            Margin = new Padding(0)
        };

        _wakeDiagnosticsTextBox = CreateDiagnosticTextBox();
        _powerRequestDiagnosticsTextBox = CreateDiagnosticTextBox();

        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 1,
            Padding = new Padding(14),
            BackColor = ShellColor
        };
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 260F));
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        root.Controls.Add(BuildSidebar(), 0, 0);
        root.Controls.Add(BuildMainArea(), 1, 0);
        Controls.Add(root);

        _stateChangedHandler = (_, _) =>
        {
            UpdateStatus();
            RefreshDiagnosticsView();
        };

        _logger.LogWritten += OnLogWritten;
        _controller.StateChanged += _stateChangedHandler;

        ApplySettingsToUi(_controller.CurrentSettings);
        UpdateStatus();
        RefreshDiagnosticsView();
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

    private Control BuildSidebar()
    {
        var panel = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = SidebarColor,
            Padding = new Padding(18),
            Margin = new Padding(0, 0, 14, 0)
        };

        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 6,
            BackColor = Color.Transparent
        };
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        var title = new Label
        {
            Text = "SleepSentinel",
            AutoSize = true,
            Font = new Font("Segoe UI", 17F, FontStyle.Bold),
            ForeColor = TextColor,
            Margin = new Padding(0, 0, 0, 4)
        };
        var subtitle = new Label
        {
            Text = "让电脑在“该睡就睡”时更像一个稳定、可控的系统守门员。",
            AutoSize = true,
            MaximumSize = new Size(210, 0),
            ForeColor = MutedTextColor,
            Margin = new Padding(0, 0, 0, 18)
        };
        var actionHeader = new Label
        {
            Text = "快捷操作",
            AutoSize = true,
            Font = new Font("Segoe UI", 10.5F, FontStyle.Bold),
            ForeColor = TextColor,
            Margin = new Padding(0, 0, 0, 8)
        };

        var actions = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            ColumnCount = 1,
            RowCount = 6,
            AutoSize = true,
            BackColor = Color.Transparent
        };
        actions.Controls.Add(CreateSidebarButton("保存常规设置", true, SaveSettings), 0, 0);
        actions.Controls.Add(CreateSidebarButton("立即睡眠", false, () => _controller.SleepNow()), 0, 1);
        actions.Controls.Add(CreateSidebarButton("立即休眠", false, () => _controller.HibernateNow()), 0, 2);
        actions.Controls.Add(CreateSidebarButton("记录唤醒诊断", false, RecordWakeDiagnostics), 0, 3);
        actions.Controls.Add(CreateSidebarButton("导出诊断报告", false, ExportDiagnosticReport), 0, 4);
        actions.Controls.Add(CreateSidebarButton("刷新日志", false, LoadLogs), 0, 5);

        var noteCard = CreateSidebarNoteCard(
            "交互说明",
            "软件唤醒、待机联网、电池兜底、远控拦截会立即生效；工作模式、恢复保护和启动行为需要点击“保存常规设置”。");
        var footer = new Label
        {
            Text = "托盘会持续显示当前状态；双击图标可随时拉回控制台。",
            AutoSize = true,
            MaximumSize = new Size(210, 0),
            ForeColor = MutedTextColor,
            Margin = new Padding(0, 16, 0, 0)
        };

        layout.Controls.Add(title, 0, 0);
        layout.Controls.Add(subtitle, 0, 1);
        layout.Controls.Add(actionHeader, 0, 2);
        layout.Controls.Add(actions, 0, 3);
        layout.Controls.Add(noteCard, 0, 4);
        layout.Controls.Add(footer, 0, 5);
        panel.Controls.Add(layout);
        return panel;
    }

    private Control BuildMainArea()
    {
        var shell = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = ShellColor
        };

        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 2,
            BackColor = Color.Transparent
        };
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

        var intro = new Label
        {
            Text = "保护控制台把“工作模式、恢复保护、四层唤醒拦截、诊断和日志”拆开，让你先看到风险，再决定怎么收口。",
            AutoSize = true,
            MaximumSize = new Size(780, 0),
            ForeColor = MutedTextColor,
            Margin = new Padding(0, 2, 0, 12)
        };

        var tabs = new TabControl
        {
            Dock = DockStyle.Fill,
            Appearance = TabAppearance.Normal,
            SizeMode = TabSizeMode.Fixed,
            ItemSize = new Size(120, 32)
        };
        tabs.Controls.Add(BuildProtectionTab());
        tabs.Controls.Add(BuildStatusTab());
        tabs.Controls.Add(BuildDiagnosticsTab());
        tabs.Controls.Add(BuildLogsTab());

        layout.Controls.Add(intro, 0, 0);
        layout.Controls.Add(tabs, 0, 1);
        shell.Controls.Add(layout);
        return shell;
    }

    private TabPage BuildProtectionTab()
    {
        var page = CreatePage("保护控制");
        var scroll = new Panel
        {
            Dock = DockStyle.Fill,
            AutoScroll = true,
            BackColor = ShellColor,
            Padding = new Padding(0, 0, 8, 8)
        };

        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            ColumnCount = 1,
            AutoSize = true,
            BackColor = Color.Transparent
        };
        layout.Controls.Add(BuildOverviewCard(), 0, 0);
        layout.Controls.Add(BuildProtectionGrid(), 0, 1);

        scroll.Controls.Add(layout);
        page.Controls.Add(scroll);
        return page;
    }

    private TabPage BuildStatusTab()
    {
        var page = CreatePage("系统状态");
        var scroll = new Panel
        {
            Dock = DockStyle.Fill,
            AutoScroll = true,
            BackColor = ShellColor,
            Padding = new Padding(0, 0, 8, 8)
        };

        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            ColumnCount = 1,
            AutoSize = true,
            BackColor = Color.Transparent
        };
        layout.Controls.Add(CreateDetailCard(
            "运行概览",
            "先看当前状态、保护规则和最近一次唤醒判断。",
            ("当前状态", "CurrentStatus"),
            ("保护规则", "ProtectionRule"),
            ("最近一次唤醒判定", "LastWakeSummary")), 0, 0);
        layout.Controls.Add(CreateDetailCard(
            "策略详情",
            "四层策略分开显示，避免再出现“合并成一句话看不出谁在生效”的问题。",
            ("唤醒定时器策略", "WakeTimerPolicy"),
            ("待机联网策略", "StandbyConnectivityPolicy"),
            ("电池兜底策略", "BatteryFallbackPolicy"),
            ("远控拦截策略", "RemoteWakePolicy")), 0, 1);
        layout.Controls.Add(CreateDetailCard(
            "路径与时间",
            "保留配置、日志和最近挂起 / 恢复时间，方便继续排障。",
            ("配置文件", "SettingsPath"),
            ("日志目录", "LogDirectory"),
            ("上次挂起", "LastSuspend"),
            ("上次恢复", "LastResume")), 0, 2);

        scroll.Controls.Add(layout);
        page.Controls.Add(scroll);
        return page;
    }

    private TabPage BuildDiagnosticsTab()
    {
        var page = CreatePage("诊断工具");

        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 3,
            Padding = new Padding(8),
            BackColor = ShellColor
        };
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 50F));

        var toolbar = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            Margin = new Padding(0, 0, 0, 12)
        };
        toolbar.Controls.Add(CreateCompactButton("记录唤醒诊断", true, RecordWakeDiagnostics));
        toolbar.Controls.Add(CreateCompactButton("刷新诊断视图", false, RefreshDiagnosticsView));
        toolbar.Controls.Add(CreateCompactButton("导出诊断报告", false, ExportDiagnosticReport));

        root.Controls.Add(toolbar, 0, 0);
        root.Controls.Add(CreateTextCard(
            "唤醒诊断",
            "对应 powercfg /lastwake 与 /waketimers 的原始输出。",
            _wakeDiagnosticsTextBox), 0, 1);
        root.Controls.Add(CreateTextCard(
            "电源请求与覆盖",
            "对应 powercfg /requests 与 /requestsoverride，便于判断谁在拖住睡眠。",
            _powerRequestDiagnosticsTextBox), 0, 2);

        page.Controls.Add(root);
        return page;
    }

    private TabPage BuildLogsTab()
    {
        var page = CreatePage("运行日志");
        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 2,
            Padding = new Padding(8),
            BackColor = ShellColor
        };
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

        var toolbar = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            Margin = new Padding(0, 0, 0, 12)
        };
        toolbar.Controls.Add(CreateCompactButton("刷新日志", false, LoadLogs));
        toolbar.Controls.Add(CreateCompactButton("导出诊断报告", false, ExportDiagnosticReport));

        var logCard = CreateCardShell(
            "实时日志",
            "日志被降到高级区，但仍然保持可滚动、可刷新、可用于排障。",
            null,
            out var body);
        var logHost = new Panel
        {
            Dock = DockStyle.Fill,
            Height = 420,
            Padding = new Padding(12),
            BackColor = Color.FromArgb(238, 243, 245)
        };
        _logTextBox.Dock = DockStyle.Fill;
        logHost.Controls.Add(_logTextBox);
        body.Controls.Add(logHost, 0, 0);

        root.Controls.Add(toolbar, 0, 0);
        root.Controls.Add(logCard, 0, 1);
        page.Controls.Add(root);
        return page;
    }

    private Control BuildOverviewCard()
    {
        var card = CreateCardShell(
            "保护总览",
            "先判断当前是否安全，再去调整每一层策略。",
            null,
            out var body);

        var headerRow = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            ColumnCount = 2,
            AutoSize = true,
            Margin = new Padding(0, 0, 0, 12)
        };
        headerRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        headerRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        headerRow.Controls.Add(_overviewStatusChipLabel, 0, 0);
        headerRow.Controls.Add(CreateCompactButton("保存常规设置", true, SaveSettings), 1, 0);

        var stats = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            ColumnCount = 2,
            AutoSize = true,
            Margin = new Padding(0, 0, 0, 12)
        };
        stats.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
        stats.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
        stats.Controls.Add(CreateStatPanel("当前模式", _overviewModeValueLabel), 0, 0);
        stats.Controls.Add(CreateStatPanel("最近一次唤醒判定", _overviewWakeValueLabel), 1, 0);

        var ruleHost = new Panel
        {
            Dock = DockStyle.Top,
            BackColor = AccentSoftColor,
            Padding = new Padding(14),
            Margin = new Padding(0),
            AutoSize = true
        };
        var ruleLayout = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            ColumnCount = 1,
            AutoSize = true,
            BackColor = Color.Transparent
        };
        ruleLayout.Controls.Add(CreateSectionLabel("保护规则"), 0, 0);
        _overviewRuleValueLabel.Margin = new Padding(0);
        ruleLayout.Controls.Add(_overviewRuleValueLabel, 0, 1);
        ruleHost.Controls.Add(ruleLayout);

        body.Controls.Add(headerRow, 0, 0);
        body.Controls.Add(stats, 0, 1);
        body.Controls.Add(ruleHost, 0, 2);
        return card;
    }

    private Control BuildProtectionGrid()
    {
        var grid = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            ColumnCount = 2,
            RowCount = 3,
            AutoSize = true,
            Margin = new Padding(0)
        };
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50F));

        grid.Controls.Add(BuildModeCard(), 0, 0);
        grid.Controls.Add(BuildResumeProtectionCard(), 1, 0);
        grid.Controls.Add(BuildTogglePolicyCard(
            "软件唤醒",
            "关闭当前电源计划的唤醒定时器。",
            _disableWakeTimersCheckbox,
            _wakeTimerQuickStateLabel,
            _wakeTimerSummaryLabel,
            _controller.ReapplyWakeTimerPolicy), 0, 1);
        grid.Controls.Add(BuildTogglePolicyCard(
            "待机联网",
            "减少现代待机期间的后台联网拉活。",
            _disableStandbyConnectivityCheckbox,
            _standbyConnectivityQuickStateLabel,
            _standbyConnectivitySummaryLabel,
            _controller.ReapplyStandbyConnectivityPolicy), 1, 1);
        grid.Controls.Add(BuildTogglePolicyCard(
            "电池兜底",
            "电池待机一段时间后自动转入休眠。",
            _enforceBatteryStandbyHibernateCheckbox,
            _batteryStandbyHibernateQuickStateLabel,
            _batteryStandbyHibernateSummaryLabel,
            _controller.ReapplyBatteryStandbyHibernatePolicy), 0, 2);
        grid.Controls.Add(BuildTogglePolicyCard(
            "远控拦截",
            "拦截常见远程软件的 DISPLAY / SYSTEM / AWAYMODE 保活请求。",
            _blockKnownRemoteWakeCheckbox,
            _remoteWakeQuickStateLabel,
            _remoteWakeSummaryLabel,
            _controller.ReapplyKnownRemoteWakePolicy), 1, 2);

        return grid;
    }

    private Control BuildModeCard()
    {
        var card = CreateCardShell(
            "工作模式与启动",
            "工作模式、启动后行为和开机自启统一在这里调整，再由“保存常规设置”生效。",
            null,
            out var body);

        body.Controls.Add(CreateSectionLabel("工作模式"), 0, 0);
        body.Controls.Add(_policyModeComboBox, 0, 1);
        body.Controls.Add(CreateSectionLabel("启动行为"), 0, 2);
        body.Controls.Add(_startMinimizedCheckbox, 0, 3);
        body.Controls.Add(_autostartCheckbox, 0, 4);
        body.Controls.Add(CreateBodyLabel("这些选项会跟恢复保护一起保存，避免误触后立即改动全局行为。"), 0, 5);
        return card;
    }

    private Control BuildResumeProtectionCard()
    {
        var card = CreateCardShell(
            "恢复保护",
            "把“异常恢复后怎么办”说清楚，不再只剩一句模糊摘要。",
            null,
            out var body);

        var detailsFlow = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            Margin = new Padding(0, 0, 0, 8),
            WrapContents = true
        };
        detailsFlow.Controls.Add(new Label
        {
            Text = "恢复后执行",
            AutoSize = true,
            Padding = new Padding(0, 7, 8, 0),
            ForeColor = MutedTextColor
        });
        detailsFlow.Controls.Add(_resumeActionComboBox);
        detailsFlow.Controls.Add(new Label
        {
            Text = "延迟秒数",
            AutoSize = true,
            Padding = new Padding(16, 7, 8, 0),
            ForeColor = MutedTextColor
        });
        detailsFlow.Controls.Add(_resumeDelayInput);

        body.Controls.Add(_resumeProtectionCheckbox, 0, 0);
        body.Controls.Add(detailsFlow, 0, 1);
        body.Controls.Add(_onlyUnattendedWakeCheckbox, 0, 2);
        body.Controls.Add(CreateBodyLabel("人工行为包括键盘、鼠标、开盖、解锁、控制台/远程接管和登录。"), 0, 3);
        return card;
    }

    private Control BuildTogglePolicyCard(
        string title,
        string subtitle,
        CheckBox toggle,
        Label stateChip,
        Label summaryLabel,
        Action reapplyAction)
    {
        var card = CreateCardShell(title, subtitle, stateChip, out var body);
        summaryLabel.Margin = new Padding(0, 4, 0, 0);

        body.Controls.Add(toggle, 0, 0);
        body.Controls.Add(CreateCompactButton("重新应用", false, () =>
        {
            reapplyAction();
            ApplySettingsToUi(_controller.CurrentSettings);
            UpdateStatus();
            LoadLogs();
        }), 0, 1);
        body.Controls.Add(summaryLabel, 0, 2);
        return card;
    }

    private Panel CreateDetailCard(string title, string subtitle, params (string Key, string FieldKey)[] rows)
    {
        var card = CreateCardShell(title, subtitle, null, out var body);

        var table = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            ColumnCount = 2,
            AutoSize = true
        };
        table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 150F));
        table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

        for (var index = 0; index < rows.Length; index++)
        {
            var key = new Label
            {
                Text = rows[index].Key,
                AutoSize = true,
                ForeColor = MutedTextColor,
                Padding = new Padding(0, 4, 12, 10),
                Margin = new Padding(0)
            };
            var value = CreateValueLabel();
            value.Margin = new Padding(0, 4, 0, 10);
            value.MaximumSize = new Size(590, 0);
            _statusValueLabels[rows[index].FieldKey] = value;

            table.Controls.Add(key, 0, index);
            table.Controls.Add(value, 1, index);
        }

        body.Controls.Add(table, 0, 0);
        return card;
    }

    private Panel CreateTextCard(string title, string subtitle, TextBox textBox)
    {
        var card = CreateCardShell(title, subtitle, null, out var body);
        var host = new Panel
        {
            Dock = DockStyle.Fill,
            Padding = new Padding(12),
            BackColor = Color.FromArgb(238, 243, 245),
            Height = 220
        };
        textBox.Dock = DockStyle.Fill;
        host.Controls.Add(textBox);
        body.Controls.Add(host, 0, 0);
        return card;
    }

    private static TabPage CreatePage(string text)
    {
        return new TabPage
        {
            Text = text,
            Padding = new Padding(8),
            BackColor = ShellColor
        };
    }

    private static Panel CreateCardShell(string title, string subtitle, Label? stateChip, out TableLayoutPanel body)
    {
        var card = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = CardColor,
            Padding = new Padding(16),
            Margin = new Padding(0, 0, 12, 12),
            BorderStyle = BorderStyle.FixedSingle,
            AutoSize = true
        };

        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 2,
            AutoSize = true,
            BackColor = Color.Transparent
        };

        var header = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            ColumnCount = 2,
            AutoSize = true,
            Margin = new Padding(0, 0, 0, 10),
            BackColor = Color.Transparent
        };
        header.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
        header.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

        var titleBlock = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 2,
            AutoSize = true,
            BackColor = Color.Transparent
        };
        titleBlock.Controls.Add(new Label
        {
            Text = title,
            AutoSize = true,
            Font = new Font("Segoe UI", 11F, FontStyle.Bold),
            ForeColor = TextColor,
            Margin = new Padding(0, 0, 0, 4)
        }, 0, 0);
        titleBlock.Controls.Add(new Label
        {
            Text = subtitle,
            AutoSize = true,
            MaximumSize = new Size(360, 0),
            ForeColor = MutedTextColor,
            Margin = new Padding(0)
        }, 0, 1);

        header.Controls.Add(titleBlock, 0, 0);
        if (stateChip is not null)
        {
            header.Controls.Add(stateChip, 1, 0);
        }

        body = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            ColumnCount = 1,
            AutoSize = true,
            BackColor = Color.Transparent
        };

        layout.Controls.Add(header, 0, 0);
        layout.Controls.Add(body, 0, 1);
        card.Controls.Add(layout);
        return card;
    }

    private Panel CreateStatPanel(string title, Label valueLabel)
    {
        var panel = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = Color.FromArgb(238, 243, 245),
            Padding = new Padding(14),
            Margin = new Padding(0, 0, 12, 0),
            AutoSize = true
        };

        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            ColumnCount = 1,
            AutoSize = true,
            BackColor = Color.Transparent
        };
        layout.Controls.Add(new Label
        {
            Text = title,
            AutoSize = true,
            ForeColor = MutedTextColor,
            Font = new Font("Segoe UI", 8.5F, FontStyle.Bold),
            Margin = new Padding(0, 0, 0, 6)
        }, 0, 0);
        valueLabel.Margin = new Padding(0);
        layout.Controls.Add(valueLabel, 0, 1);
        panel.Controls.Add(layout);
        return panel;
    }

    private static CheckBox CreateSettingsCheckBox(string text)
    {
        return new CheckBox
        {
            AutoSize = true,
            Text = text,
            MaximumSize = new Size(340, 0),
            Margin = new Padding(0, 0, 0, 8)
        };
    }

    private static Label CreateSectionLabel(string text)
    {
        return new Label
        {
            Text = text,
            AutoSize = true,
            Font = new Font("Segoe UI", 9F, FontStyle.Bold),
            ForeColor = AccentColor,
            Margin = new Padding(0, 0, 0, 6)
        };
    }

    private static Label CreateChipLabel()
    {
        return new Label
        {
            AutoSize = true,
            Padding = new Padding(10, 5, 10, 5),
            Margin = new Padding(0)
        };
    }

    private static Label CreateValueLabel()
    {
        return new Label
        {
            AutoSize = true,
            MaximumSize = new Size(420, 0),
            Font = new Font("Segoe UI", 10F, FontStyle.Bold),
            ForeColor = TextColor
        };
    }

    private static Label CreateBodyLabel(string? text = null)
    {
        return new Label
        {
            Text = text ?? string.Empty,
            AutoSize = true,
            MaximumSize = new Size(360, 0),
            ForeColor = MutedTextColor,
            Margin = new Padding(0)
        };
    }

    private static TextBox CreateDiagnosticTextBox()
    {
        return new TextBox
        {
            Multiline = true,
            ReadOnly = true,
            ScrollBars = ScrollBars.Vertical,
            BorderStyle = BorderStyle.None,
            BackColor = Color.FromArgb(238, 243, 245),
            ForeColor = TextColor,
            Font = new Font(FontFamily.GenericMonospace, 9.2f)
        };
    }

    private Button CreateSidebarButton(string text, bool primary, Action onClick)
    {
        var button = new Button
        {
            Text = text,
            Dock = DockStyle.Top,
            Height = 40,
            FlatStyle = FlatStyle.Flat,
            Margin = new Padding(0, 0, 0, 8),
            BackColor = primary ? PrimaryColor : CardColor,
            ForeColor = primary ? Color.White : TextColor
        };
        button.FlatAppearance.BorderSize = primary ? 0 : 1;
        button.FlatAppearance.BorderColor = BorderColor;
        button.Click += (_, _) => onClick();
        return button;
    }

    private Button CreateCompactButton(string text, bool primary, Action onClick)
    {
        var button = new Button
        {
            Text = text,
            AutoSize = true,
            FlatStyle = FlatStyle.Flat,
            Padding = new Padding(12, 6, 12, 6),
            Margin = new Padding(0, 0, 8, 0),
            BackColor = primary ? PrimaryColor : CardColor,
            ForeColor = primary ? Color.White : TextColor
        };
        button.FlatAppearance.BorderSize = primary ? 0 : 1;
        button.FlatAppearance.BorderColor = BorderColor;
        button.Click += (_, _) => onClick();
        return button;
    }

    private Control CreateSidebarNoteCard(string title, string body)
    {
        var card = new Panel
        {
            Dock = DockStyle.Bottom,
            BackColor = CardColor,
            Padding = new Padding(14),
            Margin = new Padding(0, 18, 0, 0),
            BorderStyle = BorderStyle.FixedSingle,
            AutoSize = true
        };

        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            ColumnCount = 1,
            AutoSize = true,
            BackColor = Color.Transparent
        };
        layout.Controls.Add(new Label
        {
            Text = title,
            AutoSize = true,
            Font = new Font("Segoe UI", 9F, FontStyle.Bold),
            ForeColor = TextColor,
            Margin = new Padding(0, 0, 0, 6)
        }, 0, 0);
        layout.Controls.Add(new Label
        {
            Text = body,
            AutoSize = true,
            MaximumSize = new Size(200, 0),
            ForeColor = MutedTextColor
        }, 0, 1);
        card.Controls.Add(layout);
        return card;
    }

    private void OnPolicyToggleChanged(CheckBox checkBox, Action enableAction, Action disableAction)
    {
        if (_suppressInteractiveToggleEvents)
        {
            return;
        }

        if (checkBox.Checked)
        {
            enableAction();
        }
        else
        {
            disableAction();
        }

        ApplySettingsToUi(_controller.CurrentSettings);
        UpdateStatus();
        LoadLogs();
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
        _suppressInteractiveToggleEvents = true;
        try
        {
            _policyModeComboBox.SelectedIndex = settings.PolicyMode == PowerPolicyMode.KeepAwakeIndefinitely ? 1 : 0;
            _resumeProtectionCheckbox.Checked = settings.ResumeProtectionEnabled;
            _resumeActionComboBox.SelectedIndex = settings.ResumeProtectionMode == ResumeProtectionMode.Hibernate ? 1 : 0;
            _onlyUnattendedWakeCheckbox.Checked = settings.ResumeProtectionOnlyForUnattendedWake;
            _resumeDelayInput.Value = Math.Clamp(settings.ResumeProtectionDelaySeconds, 3, 600);
            _startMinimizedCheckbox.Checked = settings.StartMinimized;
            _autostartCheckbox.Checked = settings.StartWithWindows;
            _disableWakeTimersCheckbox.Checked = settings.DisableWakeTimers;
            _disableStandbyConnectivityCheckbox.Checked = settings.DisableStandbyConnectivity;
            _enforceBatteryStandbyHibernateCheckbox.Checked = settings.EnforceBatteryStandbyHibernate;
            _blockKnownRemoteWakeCheckbox.Checked = settings.BlockKnownRemoteWakeRequests;
        }
        finally
        {
            _suppressInteractiveToggleEvents = false;
        }

        _resumeActionComboBox.Enabled = settings.ResumeProtectionEnabled;
        _resumeDelayInput.Enabled = settings.ResumeProtectionEnabled;
        _onlyUnattendedWakeCheckbox.Enabled = settings.ResumeProtectionEnabled;

        UpdateChip(_wakeTimerQuickStateLabel, settings.DisableWakeTimers ? "已拦截" : "未接管", settings.DisableWakeTimers ? ChipKind.Active : ChipKind.Neutral);
        UpdateChip(_standbyConnectivityQuickStateLabel, settings.DisableStandbyConnectivity ? "已拦截" : "未接管", settings.DisableStandbyConnectivity ? ChipKind.Active : ChipKind.Neutral);
        UpdateChip(_batteryStandbyHibernateQuickStateLabel, settings.EnforceBatteryStandbyHibernate ? "已兜底" : "未接管", settings.EnforceBatteryStandbyHibernate ? ChipKind.Active : ChipKind.Neutral);
        UpdateChip(_remoteWakeQuickStateLabel, settings.BlockKnownRemoteWakeRequests ? "已拦截" : "未接管", settings.BlockKnownRemoteWakeRequests ? ChipKind.Active : ChipKind.Neutral);

        _wakeTimerSummaryLabel.Text = settings.WakeTimerPolicySummary;
        _standbyConnectivitySummaryLabel.Text = settings.StandbyConnectivityPolicySummary;
        _batteryStandbyHibernateSummaryLabel.Text = settings.BatteryStandbyHibernatePolicySummary;
        _remoteWakeSummaryLabel.Text = settings.KnownRemoteWakePolicySummary;
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
        var hasProtection =
            settings.ResumeProtectionEnabled
            || settings.DisableWakeTimers
            || settings.DisableStandbyConnectivity
            || settings.EnforceBatteryStandbyHibernate
            || settings.BlockKnownRemoteWakeRequests;

        if (settings.PolicyMode == PowerPolicyMode.KeepAwakeIndefinitely)
        {
            UpdateChip(_overviewStatusChipLabel, "当前保持唤醒", ChipKind.Accent);
        }
        else if (hasProtection)
        {
            UpdateChip(_overviewStatusChipLabel, "系统正在保护中", ChipKind.Active);
        }
        else
        {
            UpdateChip(_overviewStatusChipLabel, "基础模式运行中", ChipKind.Neutral);
        }

        _overviewModeValueLabel.Text = settings.PolicyMode == PowerPolicyMode.KeepAwakeIndefinitely
            ? "无限保持唤醒"
            : "遵循电源计划";
        _overviewWakeValueLabel.Text = settings.LastWakeSummary;
        _overviewRuleValueLabel.Text = _controller.CurrentProtectionRuleSummary;

        SetDetailValue("CurrentStatus", _controller.CurrentStatus);
        SetDetailValue("ProtectionRule", _controller.CurrentProtectionRuleSummary);
        SetDetailValue("LastWakeSummary", settings.LastWakeSummary);
        SetDetailValue("WakeTimerPolicy", settings.WakeTimerPolicySummary);
        SetDetailValue("StandbyConnectivityPolicy", settings.StandbyConnectivityPolicySummary);
        SetDetailValue("BatteryFallbackPolicy", settings.BatteryStandbyHibernatePolicySummary);
        SetDetailValue("RemoteWakePolicy", settings.KnownRemoteWakePolicySummary);
        SetDetailValue("SettingsPath", _settingsStore.SettingsPath);
        SetDetailValue("LogDirectory", _logger.LogDirectory);
        SetDetailValue("LastSuspend", FormatTime(settings.LastSuspendUtc));
        SetDetailValue("LastResume", FormatTime(settings.LastResumeUtc));
    }

    private void RefreshDiagnosticsView()
    {
        if (IsDisposed)
        {
            return;
        }

        if (InvokeRequired)
        {
            BeginInvoke(new Action(RefreshDiagnosticsView));
            return;
        }

        _wakeDiagnosticsTextBox.Text = _controller.CollectWakeDiagnostics();
        _powerRequestDiagnosticsTextBox.Text = _controller.CollectPowerRequestDiagnostics();
    }

    private void LoadLogs()
    {
        if (_logTextBox.IsDisposed)
        {
            return;
        }

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

    private void RecordWakeDiagnostics()
    {
        _logger.Warn("用户手动收集唤醒诊断。");
        _logger.Warn(_controller.CollectWakeDiagnostics());
        RefreshDiagnosticsView();
        LoadLogs();
    }

    private void ExportDiagnosticReport()
    {
        try
        {
            var path = _diagnosticReportService.Export();
            UpdateStatus();
            RefreshDiagnosticsView();
            LoadLogs();
            MessageBox.Show($"诊断报告已导出：\n{path}", "SleepSentinel", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            _logger.Error($"导出诊断报告失败：{ex.Message}");
            MessageBox.Show($"导出失败：{ex.Message}", "SleepSentinel", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void SetDetailValue(string key, string value)
    {
        if (_statusValueLabels.TryGetValue(key, out var label))
        {
            label.Text = value;
        }
    }

    private static void UpdateChip(Label label, string text, ChipKind kind)
    {
        label.Text = text;
        label.ForeColor = kind switch
        {
            ChipKind.Active => PrimaryColor,
            ChipKind.Accent => AccentColor,
            ChipKind.Warning => WarningColor,
            _ => TextColor
        };
        label.BackColor = kind switch
        {
            ChipKind.Active => PrimarySoftColor,
            ChipKind.Accent => AccentSoftColor,
            ChipKind.Warning => WarningSoftColor,
            _ => Color.FromArgb(236, 240, 242)
        };
    }

    private static string FormatTime(DateTimeOffset? value)
    {
        return value?.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss") ?? "无";
    }

    private enum ChipKind
    {
        Neutral,
        Active,
        Accent,
        Warning
    }
}
