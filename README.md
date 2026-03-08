# SleepSentinel

SleepSentinel 是一个 Windows 托盘常驻工具，目标是解决两类场景：

- 电脑该按电源计划休眠时，尽量不要因为某些软件的唤醒行为长期亮机。
- 需要临时常亮时，提供类似 PowerToys Awake 的“无限保持唤醒”模式。

## 功能

- 托盘驻留，双击图标打开控制面板
- 两种模式
  - `遵循电源计划`
  - `无限保持唤醒（类似 PowerToys Awake）`
- 恢复保护
  - 监听系统 `Suspend/Resume`
  - 记录 `powercfg /lastwake` 和 `powercfg /waketimers`
  - 支持“仅对疑似非人工唤醒执行保护”
  - 在恢复后延迟几秒，自动再次睡眠或休眠
- 可选接管当前电源计划，把 `Wake Timers` 设为禁用
- 本地日志
- 一键导出诊断报告
- 开机自启
- 小体积单文件发布和 Inno Setup 安装脚本

## 重要说明

这个工具可以做到的是：

- 在 `无限保持唤醒` 模式下阻止系统按空闲策略自动睡眠。
- 在 `遵循电源计划` 模式下，系统一旦被外部程序、唤醒定时器或设备唤醒，自动记录原因并再次睡眠/休眠。
- 可以根据 `powercfg` 输出做启发式判断，疑似人工唤醒时跳过自动回睡。
- 可选对当前电源计划执行 `powercfg /setacvalueindex scheme_current sub_sleep rtcwake 0` 和 `setdcvalueindex ... rtcwake 0`，直接禁用唤醒定时器。
- 可把当前设置、最近日志和 `powercfg` 唤醒诊断导出为文本报告，便于排查或发给别人分析。

这个工具做不到的是：

- 从操作系统底层 100% 阻止所有硬件/驱动/BIOS 层面的唤醒。

如果某些设备或 BIOS 唤醒源非常强，仍需要额外检查：

- `powercfg /devicequery wake_armed`
- BIOS 里的 RTC Wake / Wake on LAN
- 电源计划中的唤醒定时器

## Windows 构建

在 Windows PowerShell 或命令行里进入项目目录后执行：

```powershell
dotnet restore
dotnet publish .\SleepSentinel.csproj -c Release -r win-x64 --self-contained false /p:PublishSingleFile=true
```

输出目录通常在：

```text
bin\Release\net8.0-windows\win-x64\publish\
```

## 安装包

如果已安装 Inno Setup，打开 [SleepSentinel.iss](/mnt/c/Users/zhang/projects/SleepSentinel/SleepSentinel.iss) 编译即可生成安装程序。

GitHub Release 也会自动附带：

- `SleepSentinel-Setup-win-x64.exe`
- 正常安装、开始菜单快捷方式、卸载入口

## GitHub 下载

仓库接入了 GitHub Actions：

- 每次推送 `main` 都会自动构建 Windows x64 版本并上传到 Actions Artifacts
- 推送形如 `v1.0.0` 的标签时，会自动创建 GitHub Release 并附带 `SleepSentinel-win-x64.zip`

当前 Release 采用 `framework-dependent` 小体积单文件发布，因此目标电脑需要安装 `.NET 8 Desktop Runtime`。

## 诊断报告

面板中的“导出诊断报告”会生成一个文本文件，默认输出到：

```text
%LocalAppData%\SleepSentinel\reports\
```

报告内容包括：

- 当前应用设置
- 最近日志
- `powercfg /lastwake`
- `powercfg /waketimers`

## 后续可扩展

- 加入 Windows 事件日志来源分析
- 增加“禁用当前电源计划唤醒定时器”的可选开关
