namespace PlanningToolkit.Excel.Excel;

internal sealed class ExcelApplicationState : IDisposable
{
    private const int CalculationManual = -4135;
    private readonly dynamic _application;
    private readonly bool _screenUpdating;
    private readonly bool _enableEvents;
    private readonly bool _displayAlerts;
    private readonly int? _calculation;
    private readonly object? _statusBar;
    private readonly string _calculationBehavior;
    private bool _disposed;

    public ExcelApplicationState(dynamic application, string statusMessage, string calculationBehavior)
    {
        _application = application ?? throw new ArgumentNullException(nameof(application));
        _screenUpdating = application.ScreenUpdating;
        _enableEvents = application.EnableEvents;
        _displayAlerts = application.DisplayAlerts;
        _calculation = null;
        _statusBar = application.StatusBar;
        _calculationBehavior = calculationBehavior;

        application.ScreenUpdating = false;
        application.EnableEvents = false;
        application.DisplayAlerts = false;
        try
        {
            if ((int)application.Workbooks.Count > 0)
            {
                _calculation = (int)application.Calculation;
                application.Calculation = CalculationManual;
            }
        }
        catch
        {
            // Excel can reject Calculation changes when no workbook is open,
            // during cell edit mode, or while another add-in controls calculation.
            _calculation = null;
        }
        application.StatusBar = statusMessage;
    }

    public void Dispose()
    {
        if (_disposed)
            return;

        if (_calculation.HasValue)
            TryRestore(() => _application.Calculation = _calculationBehavior switch
            {
                "KeepAutomatic" => -4105,
                "KeepManual" => CalculationManual,
                _ => _calculation.Value
            });
        TryRestore(() => _application.DisplayAlerts = _displayAlerts);
        TryRestore(() => _application.EnableEvents = _enableEvents);
        TryRestore(() => _application.ScreenUpdating = _screenUpdating);
        TryRestore(() => _application.StatusBar = _statusBar ?? false);
        _disposed = true;
    }

    private static void TryRestore(Action action)
    {
        try
        {
            action();
        }
        catch
        {
            // Excel might be closing. Restoration is best-effort during shutdown.
        }
    }
}
