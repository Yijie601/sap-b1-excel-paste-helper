using System.ComponentModel;
using SapB1ExcelHelper.Services;

namespace SapB1ExcelHelper;

public sealed class SupplierMappingForm : Form
{
    private readonly SupplierMappingService _mappingService;
    private readonly BindingList<MappingRow> _rows;
    private readonly DataGridView _grid;

    public SupplierMappingForm(SupplierMappingService mappingService)
    {
        _mappingService = mappingService;
        _rows = new BindingList<MappingRow>(mappingService.GetAll()
            .Select(entry => new MappingRow
            {
                SupplierName = entry.SupplierName,
                SapCode = entry.SapCode
            })
            .ToList());

        Text = "Supplier Mapping";
        StartPosition = FormStartPosition.CenterParent;
        Size = new Size(760, 560);
        MinimumSize = new Size(650, 450);
        Font = new Font("Segoe UI", 10F);

        var title = new Label
        {
            Text = "Supplier Mapping",
            Font = new Font("Segoe UI Semibold", 17F),
            AutoSize = true,
            Location = new Point(20, 16)
        };
        var help = new Label
        {
            Text = "Matching ignores letter case and trims leading/trailing spaces. No fuzzy matching is used.",
            ForeColor = Color.FromArgb(85, 91, 101),
            AutoSize = true,
            Location = new Point(23, 52)
        };

        _grid = new DataGridView
        {
            Location = new Point(22, 84),
            Size = new Size(700, 365),
            Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
            AutoGenerateColumns = false,
            AllowUserToAddRows = true,
            AllowUserToDeleteRows = true,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            MultiSelect = true,
            DataSource = _rows,
            BackgroundColor = Color.White,
            BorderStyle = BorderStyle.FixedSingle
        };
        _grid.Columns.Add(new DataGridViewTextBoxColumn
        {
            DataPropertyName = nameof(MappingRow.SupplierName),
            HeaderText = "Supplier Name",
            AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
            FillWeight = 70
        });
        _grid.Columns.Add(new DataGridViewTextBoxColumn
        {
            DataPropertyName = nameof(MappingRow.SapCode),
            HeaderText = "SAP Code",
            AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
            FillWeight = 30
        });

        var importButton = CreateButton("Import CSV", 22);
        importButton.Click += (_, _) => ImportCsv();
        var exportButton = CreateButton("Export CSV", 142);
        exportButton.Click += (_, _) => ExportCsv();
        var deleteButton = CreateButton("Delete Selected", 262, 145);
        deleteButton.Click += (_, _) => DeleteSelected();

        var saveButton = CreateButton("Save", 570, 105);
        saveButton.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        saveButton.BackColor = Color.FromArgb(31, 106, 68);
        saveButton.ForeColor = Color.White;
        saveButton.FlatAppearance.BorderSize = 0;
        saveButton.Click += (_, _) => SaveAndClose();
        var cancelButton = CreateButton("Cancel", 617, 110);
        cancelButton.Location = new Point(617, 465);
        cancelButton.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        cancelButton.Click += (_, _) => Close();

        importButton.Location = new Point(22, 465);
        exportButton.Location = new Point(142, 465);
        deleteButton.Location = new Point(262, 465);
        saveButton.Location = new Point(505, 465);

        Controls.AddRange(new Control[]
        {
            title, help, _grid, importButton, exportButton, deleteButton, saveButton, cancelButton
        });
    }

    private static Button CreateButton(string text, int x, int width = 110) => new()
    {
        Text = text,
        Location = new Point(x, 0),
        Size = new Size(width, 38),
        FlatStyle = FlatStyle.Flat,
        BackColor = Color.White,
        Cursor = Cursors.Hand
    };

    private void ImportCsv()
    {
        using var dialog = new OpenFileDialog
        {
            Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*",
            Title = "Import supplier mappings"
        };
        if (dialog.ShowDialog(this) != DialogResult.OK)
        {
            return;
        }

        try
        {
            var imported = _mappingService.Import(dialog.FileName);
            foreach (var entry in imported)
            {
                var existing = _rows.FirstOrDefault(row =>
                    row.SupplierName.Equals(entry.SupplierName, StringComparison.OrdinalIgnoreCase));
                if (existing is null)
                {
                    _rows.Add(new MappingRow { SupplierName = entry.SupplierName, SapCode = entry.SapCode });
                }
                else
                {
                    existing.SapCode = entry.SapCode;
                }
            }

            _grid.Refresh();
        }
        catch (Exception exception)
        {
            MessageBox.Show(exception.Message, "Import failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void ExportCsv()
    {
        if (!SaveGrid())
        {
            return;
        }

        using var dialog = new SaveFileDialog
        {
            Filter = "CSV files (*.csv)|*.csv",
            FileName = "supplier_mapping.csv",
            Title = "Export supplier mappings"
        };
        if (dialog.ShowDialog(this) == DialogResult.OK)
        {
            File.Copy(AppPaths.SupplierMappingFile, dialog.FileName, true);
        }
    }

    private void DeleteSelected()
    {
        var selected = _grid.SelectedRows
            .Cast<DataGridViewRow>()
            .Select(row => row.DataBoundItem as MappingRow)
            .Where(row => row is not null)
            .Cast<MappingRow>()
            .ToArray();
        foreach (var row in selected)
        {
            _rows.Remove(row);
        }
    }

    private void SaveAndClose()
    {
        if (!SaveGrid())
        {
            return;
        }

        DialogResult = DialogResult.OK;
        Close();
    }

    private bool SaveGrid()
    {
        _grid.EndEdit();
        var validRows = _rows
            .Where(row => !string.IsNullOrWhiteSpace(row.SupplierName) || !string.IsNullOrWhiteSpace(row.SapCode))
            .ToArray();
        if (validRows.Any(row => string.IsNullOrWhiteSpace(row.SupplierName) || string.IsNullOrWhiteSpace(row.SapCode)))
        {
            MessageBox.Show(
                "Every mapping row needs both Supplier Name and SAP Code.",
                "Invalid mapping",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            return false;
        }

        var duplicate = validRows
            .GroupBy(row => row.SupplierName.Trim(), StringComparer.OrdinalIgnoreCase)
            .FirstOrDefault(group => group.Count() > 1);
        if (duplicate is not null)
        {
            MessageBox.Show(
                $"Duplicate supplier: {duplicate.Key}",
                "Invalid mapping",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            return false;
        }

        _mappingService.Save(validRows.Select(row =>
            new SupplierMappingEntry(row.SupplierName, row.SapCode)));
        return true;
    }

    private sealed class MappingRow : INotifyPropertyChanged
    {
        private string _supplierName = string.Empty;
        private string _sapCode = string.Empty;

        public string SupplierName
        {
            get => _supplierName;
            set
            {
                _supplierName = value ?? string.Empty;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(SupplierName)));
            }
        }

        public string SapCode
        {
            get => _sapCode;
            set
            {
                _sapCode = value ?? string.Empty;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(SapCode)));
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;
    }
}
