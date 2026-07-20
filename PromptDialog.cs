using System.Drawing;
using System.Windows.Forms;

namespace PlanningToolkit.Excel.UI;

internal sealed class PromptDialog : Form
{
    private readonly TextBox _input;

    private PromptDialog(string title, string prompt, string defaultValue)
    {
        Text = title;
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MinimizeBox = false;
        MaximizeBox = false;
        ShowInTaskbar = false;
        ClientSize = new Size(430, 135);
        Font = SystemFonts.MessageBoxFont;

        var label = new Label
        {
            Text = prompt,
            AutoSize = false,
            Location = new Point(12, 12),
            Size = new Size(406, 28)
        };

        _input = new TextBox
        {
            Text = defaultValue,
            Location = new Point(12, 43),
            Size = new Size(406, 23)
        };

        var ok = new Button
        {
            Text = "OK",
            DialogResult = DialogResult.OK,
            Location = new Point(262, 91),
            Size = new Size(75, 28)
        };

        var cancel = new Button
        {
            Text = "Cancel",
            DialogResult = DialogResult.Cancel,
            Location = new Point(343, 91),
            Size = new Size(75, 28)
        };

        Controls.AddRange(new Control[] { label, _input, ok, cancel });
        AcceptButton = ok;
        CancelButton = cancel;
    }

    public static string? Show(string title, string prompt, string defaultValue)
    {
        using var dialog = new PromptDialog(title, prompt, defaultValue);
        return dialog.ShowDialog() == DialogResult.OK ? dialog._input.Text : null;
    }
}

