using EmployeePortal.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.Mvc.Rendering;
using OfficeOpenXml;

namespace EmployeePortal.Pages.Registration
{
    public class RegistrationModel : PageModel
    {
        private readonly string employeeFilePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "EmployeeData.xlsx");

        [BindProperty]
        public Employee Employee { get; set; }

        public List<SelectListItem> GradeOptions { get; set; }
        public List<SelectListItem> BUOptions { get; set; }
        public List<SelectListItem> ProjectCodeOptions { get; set; }
        public List<SelectListItem> ProjectNameOptions { get; set; }
        public List<SelectListItem> PODNameOptions { get; set; }
        public List<SelectListItem> OffShoreCityOptions { get; set; }
        public List<SelectListItem> TypeOptions { get; set; }
        public List<SelectListItem> TowerOptions { get; set; }

        public List<Employee> Employees { get; set; }
        private Dictionary<string, string> projectCodeToNameMapping = new Dictionary<string, string>();

        // OnGet to load dropdown options and initialize the form
        public IActionResult OnGet()
        {
            // Load dropdown options from the dropdown file
            LoadDropdownOptions();

            // Initialize an empty Employee object for the form
            Employee = new Employee();
            Employees = new List<Employee>();

            return Page();
        }

        // OnPost to save a new employee record
        public async Task<IActionResult> OnPost()
        {
            if (!ModelState.IsValid)
            {
                // Reload dropdown options if validation fails
                LoadDropdownOptions();
                return Page();
            }

            if (Employee.IsActiveInProject == null)
            {
                ModelState.AddModelError("Employee.IsActiveInProject", "Please select if the employee is active in a project.");
                LoadDropdownOptions();
                return Page();
            }

            // Ensure the employee Excel file exists, or create it if not
            var isNewFile = !System.IO.File.Exists(employeeFilePath);
            var package = new ExcelPackage(new FileInfo(employeeFilePath));

            var worksheet = package.Workbook.Worksheets["Employees"];
            if (worksheet == null)
            {
                worksheet = package.Workbook.Worksheets.Add("Employees");

                // Define headers in a single array
                var headers = new[]
                {
                    "Type","Tower","ABLGBL","TLName" ,"GBL_Lead","ProjectCode","ProjectName","PONumber"
                    ,"PODName","AltriaPODOwner","ALCSDirector","GGID","EmpId","Resource","Email","Grade",
                    "IsActiveInProject","Gender","Location","OffshoreCity","OffshoreBackup","SL","New",
                    "Transition","BU","DateOfHire","StartDate","EndDate","COR","Group","MonthlyPrice","AltriaEXP",
                    "RoleinPOD","OverallExp","Skills","Certificates"
                };

                // Populate the header row
                for (int i = 0; i < headers.Length; i++)
                {
                    worksheet.Cells[1, i + 1].Value = headers[i];
                }
            }

            var rowCount = worksheet.Dimension?.Rows ?? 1; // Get current row count

            // Add the employee data dynamically using reflection
            var employeeData = new Dictionary<string, object>
            {
                { "Type", Employee.Type },
                { "Tower", Employee.Tower },
                { "ABLGBL", Employee.ABLGBL },
                { "TLName", Employee.TLName },
                { "GBL_Lead", Employee.GBL_Lead },
                { "ProjectCode", Employee.ProjectCode },
                { "ProjectName", Employee.ProjectName },
                { "PONumber", Employee.PONumber },
                { "PODName", Employee.PODName },
                { "AltriaPODOwner", Employee.AltriaPODOwner },
                { "ALCSDirector", Employee.ALCSDirector },
                { "GGID", Employee.GGID },
                { "EmpId", Employee.EmpId },
                { "Resource", Employee.Resource },
                { "Email", Employee.Email },
                { "Grade", Employee.Grade },
                { "IsActiveInProject", Employee.IsActiveInProject },
                { "Gender", Employee.Gender },
                { "Location", Employee.Location },
                { "OffshoreCity", Employee.OffshoreCity },
                { "OffshoreBackup", Employee.OffshoreBackup },
                { "SL", Employee.SL },
                { "New", Employee.New },
                { "Transition", Employee.Transition },
                { "BU", Employee.BU },
                { "DateOfHire", Employee.DateOfHire.ToString("yyyy-MM-dd") },
                { "StartDate", Employee.StartDate.ToString("yyyy-MM-dd") },
                { "EndDate", Employee.EndDate.ToString("yyyy-MM-dd") },
                { "COR", Employee.COR },
                { "Group", Employee.Group },
                { "MonthlyPrice", Employee.MonthlyPrice.ToString() }, // Convert to string
                { "AltriaEXP", Employee.AltriaEXP.ToString() }, // Convert to string
                { "RoleinPOD", Employee.RoleinPOD },
                { "OverallExp", Employee.OverallExp.ToString() }, // Convert to string
                { "Skills", Employee.Skills },
                { "Certificates", Employee.Certificates }
            };

            // Add dynamic month data validation and assignment
            var months = new[]
            {
                "January", "February", "March", "April", "May", "June",
                "July", "August", "September", "October", "November", "December"
            };

            foreach (var month in months)
            {
                var value = (decimal)typeof(Employee).GetProperty(month)?.GetValue(Employee);
                ValidateMonthValue(value, month);
                employeeData[month] = value.ToString(); // Ensure value is converted to string
            }

            // Write all employee data to the Excel row
            int column = 1;
            foreach (var data in employeeData)
            {
                worksheet.Cells[rowCount + 1, column].Value = data.Value.ToString(); // Ensure all values are strings
                column++;
            }

            // Save the file
            await package.SaveAsync();

            // Redirect to the employee list page after saving
            return RedirectToPage("/Registration/EmployeeList");
        }

        private void ValidateMonthValue(decimal value, string monthName)
        {
            if (value < 0 || value > 1)
            {
                throw new ArgumentException($"The value for {monthName} must be between 0 and 1 (inclusive).");
            }
        }

        // Method to load dropdown options from another Excel file
        private void LoadDropdownOptions()
        {
            GradeOptions = new List<SelectListItem>();
            BUOptions = new List<SelectListItem>();
            ProjectCodeOptions = new List<SelectListItem>();
            ProjectNameOptions = new List<SelectListItem>();
            PODNameOptions = new List<SelectListItem>();
            OffShoreCityOptions = new List<SelectListItem>();
            TypeOptions = new List<SelectListItem>();
            TowerOptions = new List<SelectListItem>();

            if (System.IO.File.Exists(employeeFilePath))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                var package = new ExcelPackage(new FileInfo(employeeFilePath));

                var worksheet = package.Workbook.Worksheets["Dropdown"]; // Ensure this matches your worksheet name
                if (worksheet != null)
                {
                    var rowCount = worksheet.Dimension.Rows;

                    for (int row = 2; row <= rowCount; row++)
                    {
                        var grade = worksheet.Cells[row, 1]?.Text?.Trim();
                        var bu = worksheet.Cells[row, 2]?.Text?.Trim();
                        var projectcode = worksheet.Cells[row, 3]?.Text?.Trim();
                        var projectname = worksheet.Cells[row, 4]?.Text?.Trim();
                        var PODname = worksheet.Cells[row, 5]?.Text?.Trim();
                        var Offshore = worksheet.Cells[row, 6]?.Text?.Trim();
                        var type = worksheet.Cells[row, 7]?.Text?.Trim();
                        var tower = worksheet.Cells[row, 8]?.Text?.Trim();

                        if (!string.IsNullOrWhiteSpace(grade))
                            GradeOptions.Add(new SelectListItem { Value = grade, Text = grade });

                        if (!string.IsNullOrWhiteSpace(bu))
                            BUOptions.Add(new SelectListItem { Value = bu, Text = bu });

                        if (!string.IsNullOrWhiteSpace(projectcode) && !string.IsNullOrWhiteSpace(projectname))
                        {
                            ProjectCodeOptions.Add(new SelectListItem { Value = projectcode, Text = projectcode });
                            projectCodeToNameMapping.Add(projectcode, projectname);
                        }
                        if (!string.IsNullOrWhiteSpace(PODname))
                        {
                            PODNameOptions.Add(new SelectListItem { Value = PODname, Text = PODname });
                        }
                        if (!string.IsNullOrWhiteSpace(Offshore))
                        {
                            OffShoreCityOptions.Add(new SelectListItem { Value = Offshore, Text = Offshore });
                        }
                        if (!string.IsNullOrWhiteSpace(type))
                        {
                            TypeOptions.Add(new SelectListItem { Value = type, Text = type });
                        }
                        if (!string.IsNullOrWhiteSpace(tower))
                        {
                            TowerOptions.Add(new SelectListItem { Value = tower, Text = tower });
                        }
                    }
                }
                else
                {
                    ModelState.AddModelError("", "Worksheet 'Dropdown' not found in the dropdown file.");
                }
            }
            else
            {
                ModelState.AddModelError("", $"Dropdown file not found at {employeeFilePath}.");
            }
        }

        // Endpoint to fetch project name by code
        [HttpGet]
        public IActionResult OnGetProjectName(string projectCode)
        {
            LoadDropdownOptions();
            if (string.IsNullOrWhiteSpace(projectCode))
                return new JsonResult("Invalid Project Code");

            if (projectCodeToNameMapping.TryGetValue(projectCode, out var projectName))
                return new JsonResult(projectName);

            return new JsonResult("Project Code not found");
        }
    }
}
