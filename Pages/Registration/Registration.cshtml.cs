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

        // OnGet to load dropdown options and initialize the form
        public IActionResult OnGet()
        {
            // Load dropdown options from the dropdown file
            LoadDropdownOptions();

            // Initialize an empty Employee object for the form
            Employee = new Employee();

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

            // Ensure the employee Excel file exists, or create it if not
            var isNewFile = !System.IO.File.Exists(employeeFilePath);
            using (var package = new ExcelPackage(new FileInfo(employeeFilePath)))
            {
                var worksheet = package.Workbook.Worksheets["Employees"];
                if (worksheet == null)
                {
                    worksheet = package.Workbook.Worksheets.Add("Employees");
                    // Create header row if it's a new file
                    worksheet.Cells[1, 1].Value = "EmployeeId";
                    worksheet.Cells[1, 2].Value = "FirstName";
                    worksheet.Cells[1, 3].Value = "LastName";
                    worksheet.Cells[1, 4].Value = "Email";
                    worksheet.Cells[1, 5].Value = "Phone";
                    worksheet.Cells[1, 6].Value = "Grade";
                    worksheet.Cells[1, 7].Value = "BU";
                    worksheet.Cells[1, 8].Value = "DateOfHire";
                    worksheet.Cells[1, 9].Value = "ProjectCode";
                    worksheet.Cells[1, 10].Value = "ProjectName";
                    worksheet.Cells[1, 11].Value = "PODName";
                    worksheet.Cells[1, 12].Value = "SatrtDate";
                    worksheet.Cells[1, 13].Value = "EndDate";
                }

                var rowCount = worksheet.Dimension?.Rows ?? 1; // Get current row count
                var newEmployeeId = rowCount; // Generate a new EmployeeId (incremental)

                // Add the new employee data
                worksheet.Cells[rowCount + 1, 1].Value = newEmployeeId;
                worksheet.Cells[rowCount + 1, 2].Value = Employee.FirstName;
                worksheet.Cells[rowCount + 1, 3].Value = Employee.LastName;
                worksheet.Cells[rowCount + 1, 4].Value = Employee.Email;
                worksheet.Cells[rowCount + 1, 5].Value = Employee.Phone;
                worksheet.Cells[rowCount + 1, 6].Value = Employee.Grade;
                worksheet.Cells[rowCount + 1, 7].Value = Employee.BU;
                worksheet.Cells[rowCount + 1, 8].Value = Employee.DateOfHire.ToString("yyyy-MM-dd");
                worksheet.Cells[rowCount + 1, 9].Value = Employee.ProjectCode;
                worksheet.Cells[rowCount + 1, 10].Value = Employee.ProjectName;
                worksheet.Cells[rowCount + 1, 11].Value = Employee.PODName;
                worksheet.Cells[rowCount + 1, 12].Value = Employee.StartDate.ToString("yyyy-MM-dd");
                worksheet.Cells[rowCount + 1, 13].Value = Employee.EndDate.ToString("yyyy-MM-dd");
                

                // Save the file
                await package.SaveAsync();
            }

            // Redirect to the employee list page after saving
            return RedirectToPage("/Registration/EmployeeList");
        }

        // Method to load dropdown options from another Excel file
        private void LoadDropdownOptions()
        {
            GradeOptions = new List<SelectListItem>();
            BUOptions = new List<SelectListItem>();
            ProjectCodeOptions = new List<SelectListItem>();
            ProjectNameOptions = new List<SelectListItem>();
            PODNameOptions = new List<SelectListItem>();

            if (System.IO.File.Exists(employeeFilePath))
            {
                using (var package = new ExcelPackage(new FileInfo(employeeFilePath)))
                {
                    var worksheet = package.Workbook.Worksheets["Dropdown"]; // Ensure this matches your worksheet name
                    if (worksheet != null)
                    {
                        var rowCount = worksheet.Dimension.Rows;
                        
                            if (rowCount < 2)
                            {
                                ModelState.AddModelError("", "Dropdown data file is empty.");
                                return;
                            }

                        for (int row = 2; row <= rowCount; row++)
                        {
                            var grade = worksheet.Cells[row, 1]?.Text?.Trim();
                            var bu = worksheet.Cells[row, 2]?.Text?.Trim();
                            var projectcode = worksheet.Cells[row, 3]?.Text?.Trim();
                            var projectname = worksheet.Cells[row, 4]?.Text?.Trim();
                            var PODname = worksheet.Cells[row, 5]?.Text?.Trim();

                            if (!string.IsNullOrWhiteSpace(grade))
                            {
                                GradeOptions.Add(new SelectListItem { Value = grade, Text = grade });
                            }

                            if (!string.IsNullOrWhiteSpace(bu))
                            {
                                BUOptions.Add(new SelectListItem { Value = bu, Text = bu });
                            }
                            if (!string.IsNullOrWhiteSpace(projectcode))
                            {
                                ProjectCodeOptions.Add(new SelectListItem { Value = projectcode, Text = projectcode });
                            }
                            if (!string.IsNullOrWhiteSpace(projectname))
                            {
                                ProjectNameOptions.Add(new SelectListItem { Value = projectname, Text = projectname });
                            }
                            if (!string.IsNullOrWhiteSpace(PODname))
                            {
                                PODNameOptions.Add(new SelectListItem { Value = PODname, Text = PODname });
                            }
                        }
                    }
                    else
                    {
                        ModelState.AddModelError("", "Worksheet 'Data' not found in the dropdown file.");
                    }
                }
            }
            else
            {
                ModelState.AddModelError("", $"Dropdown file not found at {employeeFilePath}.");
            }
        }
    }
}