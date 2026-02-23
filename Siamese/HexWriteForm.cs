using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

public class HexWriteForm : Form
{
    #region Fields
    private const int TOTAL_COLS = 32;
    private const int MAX_ROWS = 10;
    private const int COL_WIDTH = 38;
    private const int ROW_HEADER_WIDTH = 55;

    private readonly Font GridFont = new Font("Consolas", 10f);

    private DataGridView grid;
    private Panel buttonPanel;
    private Button btnOK;
    private Button btnCancel;

    private readonly int totalBytes;
    private readonly int segmentsPerGrid;
    private readonly int numRows;
    private readonly byte[] initialData;

    public byte[] ResultData { get; private set; }
    #endregion

    #region Ctor
    public HexWriteForm(int totalBytes, int segmentsPerGrid, byte[] initial = null)
    {
        this.totalBytes = totalBytes;
        this.segmentsPerGrid = segmentsPerGrid;
        initialData = initial;

        numRows = segmentsPerGrid > 0
            ? Math.Min((int)Math.Ceiling((double)totalBytes / segmentsPerGrid), MAX_ROWS)
            : 0;

        Text = "Hex Data Editor";
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        ShowInTaskbar = false;

        BuildGrid();
        ApplyColumnVisibility();
        BuildButtons();

        Controls.Add(grid);
        Controls.Add(buttonPanel);

        ApplyEditableRange(initialData);
        FitWindowToGrid();

        grid.EditingControlShowing += Grid_EditingControlShowing;
        grid.KeyDown += Grid_KeyDown;

        AcceptButton = btnOK;
        CancelButton = btnCancel;
    }
    #endregion

    #region Public
    public void SetReadOnlyMode()
    {
        grid.ReadOnly = true;

        for (int idx = 0; idx < totalBytes; idx++)
        {
            if (segmentsPerGrid == 0) continue;
            int r = idx / segmentsPerGrid;
            int c = idx % segmentsPerGrid;
            if (r < grid.Rows.Count && c < grid.Columns.Count)
            {
                var cell = grid.Rows[r].Cells[c];
                cell.Style.BackColor = SystemColors.Control;
                cell.Style.ForeColor = SystemColors.GrayText;
            }
        }

        if (btnOK != null)
        {
            btnOK.Enabled = false;
            btnOK.Visible = false;
        }

        if (btnCancel != null) btnCancel.Text = "Close";
    }
    #endregion

    #region UI
    private void BuildGrid()
    {
        grid = new DataGridView
        {
            Dock = DockStyle.Top,
            BorderStyle = BorderStyle.FixedSingle,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            AllowUserToResizeColumns = false,
            AllowUserToResizeRows = false,
            MultiSelect = true,
            SelectionMode = DataGridViewSelectionMode.CellSelect,
            RowHeadersVisible = true,
            RowHeadersWidth = ROW_HEADER_WIDTH,
            AutoGenerateColumns = false,
            ScrollBars = ScrollBars.None,
            Font = GridFont
        };

        grid.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
        grid.DefaultCellStyle.Font = GridFont;
        grid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
        grid.ColumnHeadersDefaultCellStyle.Font = GridFont;
        grid.RowHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
        grid.RowHeadersDefaultCellStyle.Font = GridFont;

        grid.Columns.Clear();
        for (int c = 0; c < TOTAL_COLS; c++)
        {
            grid.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = $"S{c + 1}",
                SortMode = DataGridViewColumnSortMode.NotSortable,
                Width = COL_WIDTH,
                Visible = false
            });
        }

        grid.Rows.Clear();
        if (numRows > 0)
        {
            grid.Rows.Add(numRows);
            for (int r = 0; r < numRows; r++)
            {
                grid.Rows[r].HeaderCell.Value = $"G{r + 1}";
                grid.Rows[r].Height = 24;
            }
        }

        foreach (DataGridViewRow row in grid.Rows)
            foreach (DataGridViewCell cell in row.Cells)
            {
                cell.Value = "--";
                cell.ReadOnly = true;
                cell.Style.BackColor = SystemColors.Control;
                cell.Style.ForeColor = SystemColors.GrayText;
            }
    }

    private void ApplyColumnVisibility()
    {
        for (int c = 0; c < TOTAL_COLS; c++)
            if (grid.Columns.Count > c) grid.Columns[c].Visible = c < segmentsPerGrid;
    }

    private void BuildButtons()
    {
        buttonPanel = new Panel { Dock = DockStyle.Bottom, Height = 50, Padding = new Padding(10) };

        btnOK = new Button { Text = "Write", DialogResult = DialogResult.OK, Width = 90, Height = 30 };
        btnOK.Click += BtnOK_Click;

        btnCancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Width = 90, Height = 30 };

        buttonPanel.Controls.Add(btnOK);
        buttonPanel.Controls.Add(btnCancel);

        buttonPanel.Resize += (s, e) =>
        {
            int spacing = 8;
            btnCancel.Location = new Point(buttonPanel.Width - btnCancel.Width - 12, (buttonPanel.Height - btnCancel.Height) / 2);
            btnOK.Location = new Point(btnCancel.Left - btnOK.Width - spacing, btnCancel.Top);
        };
    }

    private void FitWindowToGrid()
    {
        int totalW = grid.RowHeadersWidth + 2;
        foreach (DataGridViewColumn col in grid.Columns)
            if (col.Visible) totalW += col.Width;

        int headerH = grid.ColumnHeadersVisible ? grid.ColumnHeadersHeight : 0;
        int rowsH = 0;
        foreach (DataGridViewRow row in grid.Rows) rowsH += row.Height;
        int totalH = headerH + rowsH + 2;

        grid.Size = new Size(totalW, totalH);
        ClientSize = new Size(totalW, totalH + buttonPanel.Height);
    }
    #endregion

    #region Data
    private void ApplyEditableRange(byte[] initialRange)
    {
        if (totalBytes <= 0 || segmentsPerGrid == 0) return;

        for (int idx = 0; idx < totalBytes; idx++)
        {
            int r = idx / segmentsPerGrid;
            int c = idx % segmentsPerGrid;

            string value = (initialRange != null && idx < initialRange.Length)
                ? initialRange[idx].ToString("X2")
                : "00";

            if (r < grid.Rows.Count && c < grid.Columns.Count)
            {
                var cell = grid.Rows[r].Cells[c];
                cell.Value = value;
                cell.ReadOnly = false;
                cell.Style.BackColor = Color.White;
                cell.Style.ForeColor = Color.Black;
            }
        }

        if (grid.Rows.Count > 0 && grid.Columns.Count > 0)
        {
            grid.ClearSelection();
            grid.CurrentCell = grid.Rows[0].Cells[0];
            grid.Rows[0].Cells[0].Selected = true;
        }
    }
    #endregion

    #region Input
    private static bool IsHexByte(string s) => Regex.IsMatch(s, @"\A[0-9A-Fa-f]{1,2}\z");

    private void Grid_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
    {
        if (e.Control is TextBox tb)
        {
            tb.CharacterCasing = CharacterCasing.Upper;
            tb.MaxLength = 2;
            tb.KeyPress -= Tb_KeyPressHexOnly;
            tb.KeyPress += Tb_KeyPressHexOnly;
        }
    }

    private void Tb_KeyPressHexOnly(object sender, KeyPressEventArgs e)
    {
        char ch = e.KeyChar;
        if (char.IsControl(ch)) return;

        bool isHexChar = (ch >= '0' && ch <= '9') ||
                         (ch >= 'A' && ch <= 'F') ||
                         (ch >= 'a' && ch <= 'f');

        e.Handled = !isHexChar;
    }

    private void Grid_KeyDown(object sender, KeyEventArgs e)
    {
        if (segmentsPerGrid == 0) return;

        if (e.Control && e.KeyCode == Keys.A)
        {
            grid.ClearSelection();
            for (int idx = 0; idx < totalBytes; idx++)
            {
                int r = idx / segmentsPerGrid;
                int c = idx % segmentsPerGrid;
                grid.Rows[r].Cells[c].Selected = true;
            }
            e.SuppressKeyPress = true;
            return;
        }

        if (e.Control && e.KeyCode == Keys.V)
        {
            string text = Clipboard.GetText();
            if (string.IsNullOrWhiteSpace(text)) return;

            var tokens = Regex.Split(text.Trim(), @"[\s,;]+").Where(t => !string.IsNullOrEmpty(t)).ToList();
            if (tokens.Count == 0) return;

            var startCell = grid.CurrentCell;
            if (startCell == null) return;

            var selectedCells = grid.SelectedCells.Cast<DataGridViewCell>().ToList();
            if (selectedCells.Count > 1 && tokens.Count == 1 && IsHexByte(tokens[0]))
            {
                string v = tokens[0].PadLeft(2, '0').ToUpper();
                foreach (var cell in selectedCells)
                {
                    if (cell.ReadOnly) continue;
                    cell.Value = v;
                }
            }
            else
            {
                int startIdx = startCell.RowIndex * segmentsPerGrid + startCell.ColumnIndex;
                int currentIdx = startIdx;
                int tokenIdx = 0;
                while (tokenIdx < tokens.Count && currentIdx < totalBytes)
                {
                    int r = currentIdx / segmentsPerGrid;
                    int c = currentIdx % segmentsPerGrid;
                    var cell = grid.Rows[r].Cells[c];

                    if (!cell.ReadOnly && IsHexByte(tokens[tokenIdx]))
                    {
                        cell.Value = tokens[tokenIdx].PadLeft(2, '0').ToUpper();
                        tokenIdx++;
                    }
                    currentIdx++;
                }
            }
            e.SuppressKeyPress = true;
        }
    }
    #endregion

    #region Result
    private void BtnOK_Click(object sender, EventArgs e)
    {
        if (totalBytes <= 0 || segmentsPerGrid == 0)
        {
            ResultData = Array.Empty<byte>();
            return;
        }

        var byteList = new List<byte>(totalBytes);

        for (int idx = 0; idx < totalBytes; idx++)
        {
            int r = idx / segmentsPerGrid;
            int c = idx % segmentsPerGrid;
            var cell = grid.Rows[r].Cells[c];
            string valueStr = (cell.Value?.ToString() ?? "00").Trim();

            if (!IsHexByte(valueStr))
            {
                MessageBox.Show($"잘못된 입력값: '{valueStr}' (at index: {idx})", "입력 오류",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                DialogResult = DialogResult.None;
                return;
            }
            byteList.Add(Convert.ToByte(valueStr, 16));
        }

        ResultData = byteList.ToArray();
    }
    #endregion
}
