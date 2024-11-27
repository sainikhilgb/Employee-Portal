using EmployeePortal.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using OfficeOpenXml;


namespace EmployeePortal.Pages.Registration
{
    public class EmployeeListModel : PageModel
    {
        private readonly string employeeFilePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "EmployeeData.xlsx");

        // Search properties with BindProperty to automatically bind to the query string parameters
        [BindProperty(SupportsGet = true)]
        public string SearchTerm { get; set; }  // The term entered in the search field

        [BindProperty(SupportsGet = true)]
        public string SearchBy { get; set; }  // The field to search by (EmpId, FirstName, LastName, Email)

        public List<Employee> Employees { get; set; } = new List<Employee>();

        // OnGet to load and filter employees based on the search criteria
        public IActionResult OnGet()
        {
            // If SearchBy is not specified, default to 'EmpId'
            if (string.IsNullOrWhiteSpace(SearchBy))
            {
                SearchBy = "EmpId";  // Default search field
            }

            // Load employees from the Excel file
            LoadEmployees();

            // If SearchTerm and SearchBy are provided, filter the list of employees
            if (!string.IsNullOrWhiteSpace(SearchTerm))
            {
                Employees = Employees.Where(e =>
                {
                    if (SearchBy == "EmpId")
                    {
                        return e.EmpId.ToString().Contains(SearchTerm);
                    }
                    else if (SearchBy == "Resource")
                    {
                        return e.Resource.Contains(SearchTerm, StringComparison.OrdinalIgnoreCase);
                    }
                    else if (SearchBy == "Email")
                    {
                        return e.Email.Contains(SearchTerm, StringComparison.OrdinalIgnoreCase);
                    }
                    else
                    {
                        return true;  // If no valid search field, return all employees
                    }
                }).ToList();
            }

            return Page();
        }

        // Method to load employees from the Excel file
        private void LoadEmployees()
        {
            if (System.IO.File.Exists(employeeFilePath))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                var package = new ExcelPackage(new FileInfo(employeeFilePath));

                var worksheet = package.Workbook.Worksheets["Employees"];
                if (worksheet != null)
                {
                    var rowCount = worksheet.Dimension.Rows;

                    for (int row = 2; row <= rowCount; row++)
                    {
                        var employee = new Employee
                        {
                            Type = worksheet.Cells[row, 1].Text,
                            Tower = worksheet.Cells[row, 2].Text,
                            ABLGBL = worksheet.Cells[row, 3].Text,
                            TLName = worksheet.Cells[row, 4].Text,
                            GBL_Lead = worksheet.Cells[row, 5].Text,
                            ProjectCode = ParseInt(worksheet.Cells[row, 6].Text),
                            ProjectName = worksheet.Cells[row, 7].Text,
                            PONumber = ParseInt(worksheet.Cells[row, 8].Text),
                            PODName = worksheet.Cells[row, 9].Text,
                            AltriaPODOwner = worksheet.Cells[row, 10].Text,
                            ALCSDirector = worksheet.Cells[row, 11].Text,
                            GGID = ParseInt(worksheet.Cells[row, 12].Text),
                            EmpId = ParseInt(worksheet.Cells[row, 13].Text),
                            Resource = worksheet.Cells[row, 14].Text,
                            Email = worksheet.Cells[row, 15].Text,
                            Grade = worksheet.Cells[row, 16].Text,
                            IsActiveInProject = worksheet.Cells[row, 17].Text,
                            Gender = worksheet.Cells[row, 18].Text,
                            Location = worksheet.Cells[row, 19].Text,
                            OffshoreCity = worksheet.Cells[row, 20].Text,
                            OffshoreBackup = worksheet.Cells[row, 21].Text,
                            SL = worksheet.Cells[row, 22].Text,
                            New = worksheet.Cells[row, 23].Text,
                            Transition = worksheet.Cells[row, 24].Text,
                            BU = worksheet.Cells[row, 25].Text,
                            DateOfHire = ParseDate(worksheet.Cells[row, 26].Text),
                            StartDate = ParseDate(worksheet.Cells[row, 27].Text),
                            EndDate = ParseDate(worksheet.Cells[row, 28].Text),
                            COR = worksheet.Cells[row, 29].Text,
                            Group = worksheet.Cells[row, 30].Text,
                            MonthlyPrice = ParseDecemal(worksheet.Cells[row, 31].Text),
                            AltriaEXP = ParseDecemal(worksheet.Cells[row, 32].Text),
                            RoleinPOD = worksheet.Cells[row, 33].Text,
                            OverallExp = worksheet.Cells[row, 34].Text,
                            Skills = worksheet.Cells[row, 35].Text,
                            Certificates = worksheet.Cells[row, 36].Text

                        };

                        Employees.Add(employee);
                    }
                }
            }
        }

                 private DateTime ParseDate(string dateString)
                {
                    if (DateTime.TryParse(dateString, out var date))
                    {
                        return date;
                    }
                    return DateTime.MinValue; // Default value for invalid or missing dates
                }

                private int ParseInt(string numberString)
                {
                    if (int.TryParse(numberString, out var number))
                    {
                        return number;
                    }
                    return 0; // Default value for invalid or missing numbers
                }
                private decimal ParseDecemal(string numberString)
                {
                    if (decimal.TryParse(numberString, out var number))
                    {
                        return number;
                    }
                    return 0; // Default value for invalid or missing numbers
                }
    }
}
