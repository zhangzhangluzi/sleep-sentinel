namespace SleepSentinel.Models;

public sealed class PowerPlanRestoreSnapshot
{
    public string PlanName { get; set; } = string.Empty;
    public bool WakeTimerCaptured { get; set; }
    public int WakeTimerAcValue { get; set; }
    public int WakeTimerDcValue { get; set; }
    public bool StandbyConnectivityCaptured { get; set; }
    public int StandbyConnectivityAcValue { get; set; }
    public int StandbyConnectivityDcValue { get; set; }
    public int DisconnectedStandbyModeAcValue { get; set; }
    public int DisconnectedStandbyModeDcValue { get; set; }
    public bool BatteryStandbyHibernateCaptured { get; set; }
    public int BatteryStandbyHibernateAcValue { get; set; }
    public int BatteryStandbyHibernateDcValue { get; set; }
}
