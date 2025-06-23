namespace JobTracker;

using OfficeOpenXml;

public partial class Form1 : Form
{
    // Make sure these fields are not null
    private TextBox? companyNameTextBox;
    private Button? submitButton;
    private TextBox? jobTitleTextBox;
    private CheckBox? easyApplyCheckBox;


    public Form1()
    {
        InitializeComponent();
        this.Text = "Job Application Tracker";
        // License setup for EPPlus
        ExcelPackage.License.SetNonCommercialPersonal("AzmiAyoub");
        InitializeUI();
    }

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
                worksheet.Cells[lastRow + 1, 5].Value = DateTime.Today.ToString("yyyy-MM-dd");

                package.Save();
            }

            MessageBox.Show("Application recorded successfully in " + filePath);
            companyNameTextBox.Clear();
            jobTitleTextBox.Clear();
            easyApplyCheckBox!.Checked = false;
        }
        catch (Exception ex)
        {
            // Handle exceptions such as file not found, access denied, etc.
            MessageBox.Show($"Error saving application: {ex.Message}");
        }
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
}
