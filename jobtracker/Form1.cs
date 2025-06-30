namespace JobTracker;

using OfficeOpenXml;
using System.Data;

public partial class Form1 : Form
{
    // Make sure these fields are not null
    private TextBox? companyNameTextBox;
    private Button? submitButton;
    private TextBox? jobTitleTextBox;
    private CheckBox? easyApplyCheckBox;
    private DateTimePicker? datePicker = null;



    public Form1()
    {
        InitializeComponent();
        this.Text = "Job Application Tracker";
        // License setup for EPPlus
        ExcelPackage.License.SetNonCommercialPersonal("AzmiAyoub");
        InitializeUI();
        this.Shown += (s, e) =>
        {
            // Load the Excel data when the form is shown
            LoadExcelData();
        };
    }

    private DataGridView? dataGridView;
    private Button? loadButton;
    private Button? saveButton;

    private void InitializeUI()
    {
        // Company Name Input
        companyNameTextBox = new TextBox();
        Label companyNameLabel = new Label();
        companyNameLabel.Text = "Company Name:";
        companyNameTextBox.Name = "companyNameTextBox";
        // TODO: Modify the location and size as needed later
        companyNameTextBox.Location = new Point(20, 20);
        companyNameLabel.Location = new Point(20, 0);
        companyNameTextBox.Size = new Size(200, 20);
        this.Controls.Add(companyNameTextBox);
        this.Controls.Add(companyNameLabel);

        // Job Title Input
        jobTitleTextBox = new TextBox();
        Label jobTitleLabel = new Label();
        jobTitleLabel.Text = "Job Title (optional):";
        jobTitleLabel.AutoSize = true;
        jobTitleTextBox.Name = "jobTitleTextBox";
        //TODO: Modify the location and size as needed later
        jobTitleTextBox.Location = new Point(20, 70);
        jobTitleLabel.Location = new Point(20, 50);
        jobTitleTextBox.Size = new Size(200, 20);
        this.Controls.Add(jobTitleTextBox);
        this.Controls.Add(jobTitleLabel);

        // Easy Apply CheckBox
        easyApplyCheckBox = new CheckBox();
        easyApplyCheckBox.Name = "easyApplyCheckBox";
        easyApplyCheckBox.Text = "Easy Apply";
        easyApplyCheckBox.Location = new Point(230, 20);
        this.Controls.Add(easyApplyCheckBox);

        // Submit Button
        submitButton = new Button();
        submitButton.Name = "submitButton";
        submitButton.Text = "Submit Application";
        //TODO: Modify the location and size as needed later
        submitButton.Location = new Point(20, 100);
        submitButton.Size = new Size(200, 30);
        submitButton.Click += SubmitButton_Click!;
        this.Controls.Add(submitButton);

        // Load Data Button
        loadButton = new Button();
        loadButton.Name = "loadButton";
        loadButton.Text = "Refresh";
        loadButton.Location = new Point(230, 100);
        loadButton.Size = new Size(150, 30);
        loadButton.Click += RefreshButton_Click!;
        this.Controls.Add(loadButton);

        // Data Grid View
        dataGridView = new DataGridView();
        dataGridView.Name = "dataGridView";
        dataGridView.Location = new Point(20, 150);
        // dataGridView.Size = new Size(600, 300);
        // dataGridView.ScrollBars = ScrollBars.None;
        dataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
        dataGridView.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
        dataGridView.AllowUserToAddRows = false; // Disable the new row at the bottom
        dataGridView.AllowUserToDeleteRows = true; // Allow user to delete rows
        dataGridView.AllowUserToResizeColumns = true;
        dataGridView.AllowUserToResizeRows = true;
        dataGridView.ReadOnly = false;
        dataGridView.AllowUserToOrderColumns = true;
        dataGridView.MultiSelect = false;
        dataGridView.EditingControlShowing += DataGridView_EditingControlShowing;
        this.Controls.Add(dataGridView);

        // Save Button
        saveButton = new Button();
        saveButton.Name = "saveButton";
        saveButton.Text = "Save Changes";
        saveButton.Location = new Point(400, 100);
        saveButton.Size = new Size(150, 30);
        saveButton.Click += SaveChangesButton_Click!;
        this.Controls.Add(saveButton);
    }

    private void SubmitButton_Click(object sender, EventArgs e)
    {
        // Validate input. Company name must not be empty. Company Text Box is not null at this point hence the !.
        string companyName = companyNameTextBox!.Text;
        string jobTitle = jobTitleTextBox!.Text;
        bool isEasyApply = easyApplyCheckBox?.Checked ?? false;
        if (string.IsNullOrWhiteSpace(companyName))
        {
            MessageBox.Show("Please enter a company name");
            return;
        }
        if (string.IsNullOrWhiteSpace(jobTitle))
        {
            jobTitle = "N/A"; // Default value if job title is not provided
        }
        // TODO: Get user to choose a file path to save the Excel file.
        string filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "job_applications.xlsx");

        try
        {
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets.FirstOrDefault() ??
                               package.Workbook.Worksheets.Add("Applications");

                EnsureWorksheetHasHeaders(worksheet);
                int lastRow = worksheet.Dimension?.End.Row ?? 1;
                worksheet.Cells[lastRow + 1, 1].Value = companyName;
                worksheet.Cells[lastRow + 1, 2].Value = jobTitle;
                worksheet.Cells[lastRow + 1, 3].Value = isEasyApply ? "Yes" : "No";
                worksheet.Cells[lastRow + 1, 4].Value = "Waiting for response.";
                worksheet.Cells[lastRow + 1, 5].Value = DateTime.Today.ToString("dd-MM-yyyy");

                package.Save();
            }

            MessageBox.Show("Application recorded successfully in " + filePath);
            companyNameTextBox.Clear();
            jobTitleTextBox.Clear();
            easyApplyCheckBox!.Checked = false;
            LoadExcelData();
        }
        catch (Exception ex)
        {
            // Handle exceptions such as file not found, access denied, etc.
            MessageBox.Show($"Error saving application: {ex.Message}");
        }
    }
    private void LoadExcelData()
    {
        string filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "job_applications.xlsx");

        if (!File.Exists(filePath))
        {
            MessageBox.Show("No applications file found");
            return;
        }

        try
        {
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                if (worksheet == null || worksheet.Dimension == null)
                {
                    MessageBox.Show("No data found in worksheet");
                    var dataTableEmpty = new DataTable();
                    dataTableEmpty.Columns.Add("Company Name");
                    dataTableEmpty.Columns.Add("Job Title");
                    dataTableEmpty.Columns.Add("Easy Apply");
                    dataTableEmpty.Columns.Add("Status");
                    dataTableEmpty.Columns.Add("Date Applied");
                    dataGridView!.DataSource = dataTableEmpty;
                    ResizeDataGridView();
                    return;
                }

                var dataTable = new DataTable();
                // Add columns
                for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                {
                    dataTable.Columns.Add(worksheet.Cells[1, col].Text);
                }

                // Add rows
                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    var newRow = dataTable.NewRow();
                    for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                    {
                        newRow[col - 1] = worksheet.Cells[row, col].Text;
                    }
                    dataTable.Rows.Add(newRow);
                }

                dataGridView!.DataSource = dataTable;
            }
            // Replace "Easy Apply" text column with ComboBox
            var easyApplyColumn = new DataGridViewComboBoxColumn
            {
                DataPropertyName = "Easy Apply",
                HeaderText = "Easy Apply",
                Name = "Easy Apply",
                DataSource = new string[] { "Yes", "No" }
            };
            ReplaceColumn(dataGridView, "Easy Apply", easyApplyColumn);

            // Replace "Status" text column with ComboBox
            var statusColumn = new DataGridViewComboBoxColumn
            {
                DataPropertyName = "Status",
                HeaderText = "Status",
                Name = "Status",
                DataSource = new string[] { "Waiting for response.", "Rejected", "Accepted", "Interviewing" }
            };
            ReplaceColumn(dataGridView, "Status", statusColumn);
            ResizeDataGridView();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error loading applications: {ex.Message}");
        }
    }

    private void RefreshButton_Click(object sender, EventArgs e)
    {
        LoadExcelData();
    }

    private void EnsureWorksheetHasHeaders(ExcelWorksheet sheet)
    {
        if (sheet.Dimension == null || sheet.Dimension.End.Row < 1)
        {
            sheet.Cells[1, 1].Value = "Company Name";
            sheet.Cells[1, 2].Value = "Job Title";
            sheet.Cells[1, 3].Value = "Easy Apply";
            sheet.Cells[1, 4].Value = "Status";
            sheet.Cells[1, 5].Value = "Date Applied";
        }
    }

    private void ReplaceColumn(DataGridView grid, string columnName, DataGridViewColumn newColumn)
    {
        int index = grid.Columns[columnName].Index;
        grid.Columns.RemoveAt(index);
        grid.Columns.Insert(index, newColumn);
    }

    private void ResizeDataGridView()
    {
        Console.WriteLine("Resizing DataGridView...");
        if (dataGridView != null)
        {
            Console.WriteLine("DataGridView is not null, resizing...");
            // Resize DataGridView to fit content
            dataGridView.PerformLayout();
            dataGridView.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
            int totalWidth = dataGridView.RowHeadersWidth;
            foreach (DataGridViewColumn column in dataGridView.Columns)
            {
                totalWidth += column.Width;
                // Console.WriteLine("Width of column '{0}': {1}", column.HeaderText, column.Width);
            }

            int totalHeight = dataGridView.ColumnHeadersHeight;
            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                totalHeight += row.Height;
            }

            // Set the size of the DataGridView
            dataGridView.Size = new Size(totalWidth, totalHeight + 20);
        }
    }

    private void SaveChangesButton_Click(object sender, EventArgs e)
    {
        string filePath = Environment.GetEnvironmentVariable("JOB_TRACKER_FILE") ??
                          Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "job_applications.xlsx");

        try
        {
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                ExcelWorksheet? worksheet = package.Workbook.Worksheets.FirstOrDefault();

                if (worksheet == null)
                {
                    worksheet = package.Workbook.Worksheets.Add("Applications");
                }

                if (dataGridView == null || dataGridView.Columns.Count == 0)
                {
                    MessageBox.Show("No data to save.");
                    return;
                }

                // Clear existing data
                worksheet.Cells.Clear();

                // Write headers
                for (int col = 0; col < dataGridView.Columns.Count; col++)
                {
                    worksheet.Cells[1, col + 1].Value = dataGridView.Columns[col].HeaderText;
                }

                // Write data rows
                for (int row = 0; row < dataGridView?.Rows.Count; row++)
                {
                    for (int col = 0; col < dataGridView?.Columns.Count; col++)
                    {
                        worksheet.Cells[row + 2, col + 1].Value = dataGridView.Rows[row].Cells[col].Value?.ToString();
                    }
                }

                package.Save();
            }

            MessageBox.Show("Changes saved successfully to Excel!", "Saved", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error saving changes: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

    }

    private void DataGridView_EditingControlShowing(object? sender, DataGridViewEditingControlShowingEventArgs e)
    {
        if (dataGridView == null)
        {
            // Fallback if dataGridView is null, should not happen!
            Console.WriteLine("DataGridView is null, cannot show editing control.");
            return;
        }
        var currentCell = dataGridView.CurrentCell;
        if (currentCell == null || currentCell.OwningColumn.HeaderText != "Date Applied")
            return;

        if (datePicker != null)
        {
            this.Controls.Remove(datePicker);
            datePicker.Dispose();
        }

        datePicker = new DateTimePicker
        {
            Format = DateTimePickerFormat.Short,
            Visible = true
        };

        if (DateTime.TryParse(currentCell.Value?.ToString(), out DateTime currentValue))
            datePicker.Value = currentValue;
        else
            datePicker.Value = DateTime.Today;

        var rect = dataGridView.GetCellDisplayRectangle(currentCell.ColumnIndex, currentCell.RowIndex, true);
        datePicker.Location = new Point(rect.X + dataGridView.Left, rect.Y + dataGridView.Top);
        datePicker.Size = rect.Size;

        datePicker.CloseUp += (s, ev) =>
        {
            currentCell.Value = datePicker.Value.ToShortDateString();
            this.Controls.Remove(datePicker);
            datePicker.Dispose();
            datePicker = null;
        };

        this.Controls.Add(datePicker);
        datePicker.BringToFront();
        datePicker.Focus();
    }


}
