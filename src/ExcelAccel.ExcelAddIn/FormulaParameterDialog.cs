using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;
using ExcelAccel.Application.Formulas;
using ExcelAccel.ExcelAddIn.Reliability;

namespace ExcelAccel.ExcelAddIn;

internal static class FormulaParameterDialog
{
    public static bool TryGetSpacingInterval(string title, out int interval)
    {
        interval = 0;
        if (!TryShow(title, new[] { Field("Interval (1-10,000)", "2") }, false, out var values, out _)) return false;
        while (!int.TryParse(values[0], NumberStyles.None, CultureInfo.InvariantCulture, out interval) || interval < 1 || interval > 10000)
        {
            ShowValidation("Interval must be a whole number from 1 through 10,000.");
            if (!TryShow(title, new[] { Field("Interval (1-10,000)", values[0]) }, false, out values, out _)) return false;
        }
        return true;
    }

    public static bool TryGetNumericSequence(out double start, out double step, out SequenceFillDirection direction)
    {
        start = 0; step = 0; direction = SequenceFillDirection.Down;
        var fields = new[] { Field("Invariant start", "1"), Field("Invariant step", "1") };
        while (true)
        {
            if (!TryShow("Fill Numeric Sequence", fields, true, out var values, out direction)) return false;
            if (double.TryParse(values[0], NumberStyles.Float, CultureInfo.InvariantCulture, out start) &&
                double.TryParse(values[1], NumberStyles.Float, CultureInfo.InvariantCulture, out step) &&
                !double.IsNaN(start) && !double.IsInfinity(start) && !double.IsNaN(step) && !double.IsInfinity(step)) return true;
            ShowValidation("Start and step must be finite invariant numbers, using '.' as the decimal separator.");
            fields = new[] { Field("Invariant start", values[0]), Field("Invariant step", values[1]) };
        }
    }

    public static bool TryGetDateSequence(out DateTime start, out int stepDays, out SequenceFillDirection direction)
    {
        start = default; stepDays = 0; direction = SequenceFillDirection.Down;
        var fields = new[] { Field("Start date (yyyy-MM-dd)", DateTime.Today.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)), Field("Step in whole days", "1") };
        while (true)
        {
            if (!TryShow("Fill Date Sequence", fields, true, out var values, out direction)) return false;
            if (DateTime.TryParseExact(values[0], "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out start) &&
                int.TryParse(values[1], NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out stepDays)) return true;
            ShowValidation("Use an exact yyyy-MM-dd date and a whole-number day step.");
            fields = new[] { Field("Start date (yyyy-MM-dd)", values[0]), Field("Step in whole days", values[1]) };
        }
    }

    private static KeyValuePair<string, string> Field(string label, string value) => new KeyValuePair<string, string>(label, value);

    private static bool TryShow(string title, IReadOnlyList<KeyValuePair<string, string>> fields, bool includeDirection,
        out string[] values, out SequenceFillDirection direction)
    {
        using (var form = new Form
        {
            Text = title, StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog,
            MinimizeBox = false, MaximizeBox = false, ShowInTaskbar = false, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink,
            AccessibleName = title,
        })
        using (var layout = new TableLayoutPanel { AutoSize = true, ColumnCount = 2, Padding = new Padding(12) })
        {
            var inputs = new List<TextBox>();
            for (var index = 0; index < fields.Count; index++)
            {
                var label = new Label { Text = fields[index].Key, AutoSize = true, Anchor = AnchorStyles.Left };
                var input = new TextBox { Text = fields[index].Value, Width = 220, AccessibleName = fields[index].Key };
                layout.Controls.Add(label, 0, index); layout.Controls.Add(input, 1, index); inputs.Add(input);
            }
            ComboBox? directionBox = null;
            if (includeDirection)
            {
                directionBox = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 220, AccessibleName = "Fill direction" };
                directionBox.Items.AddRange(new object[] { "Down (column-major)", "Right (row-major)" }); directionBox.SelectedIndex = 0;
                layout.Controls.Add(new Label { Text = "Direction", AutoSize = true, Anchor = AnchorStyles.Left }, 0, fields.Count);
                layout.Controls.Add(directionBox, 1, fields.Count);
            }
            var buttons = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.RightToLeft, Dock = DockStyle.Fill };
            var ok = new Button { Text = "Plan", DialogResult = DialogResult.OK, AutoSize = true };
            var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, AutoSize = true };
            buttons.Controls.Add(ok); buttons.Controls.Add(cancel);
            layout.Controls.Add(buttons, 0, fields.Count + (includeDirection ? 1 : 0)); layout.SetColumnSpan(buttons, 2);
            form.Controls.Add(layout); form.AcceptButton = ok; form.CancelButton = cancel;
            var owner = ExcelWindowOwner.TryCreate();
            var result = owner is null ? form.ShowDialog() : form.ShowDialog(owner);
            values = inputs.ConvertAll(input => input.Text).ToArray();
            direction = directionBox?.SelectedIndex == 1 ? SequenceFillDirection.Right : SequenceFillDirection.Down;
            return result == DialogResult.OK;
        }
    }

    private static void ShowValidation(string message)
    {
        var owner = ExcelWindowOwner.TryCreate();
        if (owner is null) MessageBox.Show(message, "ExcelAccel parameters", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        else MessageBox.Show(owner, message, "ExcelAccel parameters", MessageBoxButtons.OK, MessageBoxIcon.Warning);
    }
}
