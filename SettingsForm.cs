using System.Drawing;
using System.Windows.Forms;
using PlanningToolkit.Core;
using PlanningToolkit.Infrastructure.Settings;

namespace PlanningToolkit.Excel.UI;

internal sealed class SettingsForm : Form
{
    private readonly ISettingsStore _store;
    private readonly PropertyGrid _propertyGrid;

    public SettingsForm(AppSettings current, ISettingsStore store)
    {
        ArgumentNullException.ThrowIfNull(current);
        _store = store ?? throw new ArgumentNullException(nameof(store));

        Text = "Planning Toolkit Settings";
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.Sizable;
        MinimumSize = new Size(620, 520);
        ClientSize = new Size(700, 610);
        Font = SystemFonts.MessageBoxFont;

        _propertyGrid = new PropertyGrid
        {
            Dock = DockStyle.Fill,
            SelectedObject = Copy(current),
            PropertySort = PropertySort.Categorized,
            ToolbarVisible = true,
            HelpVisible = true
        };

        var footer = new FlowLayoutPanel
        {
            Dock = DockStyle.Bottom,
            Height = 48,
            FlowDirection = FlowDirection.RightToLeft,
            Padding = new Padding(8)
        };

        var save = new Button { Text = "Save", AutoSize = true };
        save.Click += SaveClicked;
        var cancel = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };
        var defaults = new Button { Text = "Restore Defaults", AutoSize = true };
        defaults.Click += (_, _) => _propertyGrid.SelectedObject = new AppSettings();

        footer.Controls.Add(save);
        footer.Controls.Add(cancel);
        footer.Controls.Add(defaults);

        Controls.Add(_propertyGrid);
        Controls.Add(footer);
        CancelButton = cancel;
    }

    private void SaveClicked(object? sender, EventArgs eventArgs)
    {
        if (_propertyGrid.SelectedObject is not AppSettings settings)
            return;

        var errors = settings.Validate();
        if (errors.Count > 0)
        {
            MessageBox.Show(
                string.Join(Environment.NewLine, errors),
                "Invalid Settings",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            return;
        }

        try
        {
            _store.Save(settings);
            DialogResult = DialogResult.OK;
            Close();
        }
        catch (Exception exception)
        {
            MessageBox.Show(
                $"Settings could not be saved.\n\n{exception.Message}",
                "Planning Toolkit",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
    }

    private static AppSettings Copy(AppSettings source) => new()
    {
        SchemaVersion = source.SchemaVersion,
        DateFormat = source.DateFormat,
        CurrencyCode = source.CurrencyCode,
        WorkingHoursPerDay = source.WorkingHoursPerDay,
        WorkingDaysPerWeek = source.WorkingDaysPerWeek,
        CriticalFloatThresholdDays = source.CriticalFloatThresholdDays,
        LongDurationThresholdDays = source.LongDurationThresholdDays,
        LagThresholdDays = source.LagThresholdDays,
        GanttBaselineColor = source.GanttBaselineColor,
        GanttCurrentColor = source.GanttCurrentColor,
        GanttCriticalColor = source.GanttCriticalColor,
        WbsLevel1Color = source.WbsLevel1Color,
        WbsLevel2Color = source.WbsLevel2Color,
        DefaultOutputFolder = source.DefaultOutputFolder,
        CalculationBehavior = source.CalculationBehavior,
        LoggingLevel = source.LoggingLevel
    };
}
