using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using EmployeePortal.Models;
using Microsoft.AspNetCore.Mvc.Rendering;
using OfficeOpenXml;
using System.Linq;



namespace EmployeePortal.Pages.Registration
{
    public class EditModel : PageModel
    {
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
        private Dictionary<string, string> projectCodeToNameMapping = new Dictionary<string, string>();
        

        private readonly string employeeFilePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "EmployeeData.xlsx");
    


        public void OnGet(int id)
        {
            LoadDropdownOptions();
            // Fetch the employee to be edited using the ID
            Employee = GetEmployeeById(id);
        }

        public IActionResult OnPost(int id)
        {
            if (!ModelState.IsValid)
            {
                LoadDropdownOptions();
                return Page();
            }

            var package = new ExcelPackage(new FileInfo(employeeFilePath));
            
                var worksheet = package.Workbook.Worksheets["Employees"];

                var row = FindEmployeeRow(worksheet, id);
                if (row != -1)
                {
                    worksheet.Cells[row, 1].Value = Employee.Type;
                    worksheet.Cells[row, 2].Value = Employee.Tower;
                    worksheet.Cells[row, 3].Value = Employee.ABLGBL;
                    worksheet.Cells[row, 4].Value = Employee.TLName;
                    worksheet.Cells[row, 5].Value = Employee.GBL_Lead;
                    worksheet.Cells[row, 6].Value = Employee.ProjectCode;
                    worksheet.Cells[row, 7].Value = Employee.ProjectName;
                    worksheet.Cells[row, 8].Value = Employee.PONumber;
                    worksheet.Cells[row, 9].Value = Employee.PODName;
                    worksheet.Cells[row, 10].Value = Employee.AltriaPODOwner;
                    worksheet.Cells[row, 11].Value = Employee.ALCSDirector;
                    worksheet.Cells[row, 12].Value = Employee.GGID;
                    worksheet.Cells[row, 13].Value = Employee.EmpId;
                    worksheet.Cells[row, 14].Value = Employee.Resource;
                    worksheet.Cells[row, 15].Value = Employee.Email;
                    worksheet.Cells[row, 16].Value = Employee.Grade;
                    worksheet.Cells[row, 17].Value = Employee.IsActiveInProject;
                    worksheet.Cells[row, 18].Value = Employee.Gender;
                    worksheet.Cells[row, 19].Value = Employee.Location;
                    worksheet.Cells[row, 20].Value = Employee.OffshoreCity;
                    worksheet.Cells[row, 21].Value = Employee.OffshoreBackup;
                    worksheet.Cells[row, 22].Value = Employee.SL;
                    worksheet.Cells[row, 23].Value = Employee.New;
                    worksheet.Cells[row, 24].Value = Employee.Transition;
                    worksheet.Cells[row, 25].Value = Employee.BU;
                    worksheet.Cells[row, 26].Value = Employee.DateOfHire.ToString("yyyy-MM-dd");
                    worksheet.Cells[row, 27].Value = Employee.StartDate.ToString("yyyy-MM-dd");
                    worksheet.Cells[row, 28].Value = Employee.EndDate.ToString("yyyy-MM-dd");
                    worksheet.Cells[row, 29].Value = Employee.COR;
                    worksheet.Cells[row, 30].Value = Employee.Group;
                    worksheet.Cells[row, 31].Value = Employee.January;
                    worksheet.Cells[row, 32].Value = Employee.February;
                    worksheet.Cells[row, 33].Value = Employee.March;
                    worksheet.Cells[row, 34].Value = Employee.April;
                    worksheet.Cells[row, 35].Value = Employee.May;
                    worksheet.Cells[row, 36].Value = Employee.June;
                    worksheet.Cells[row, 37].Value = Employee.July;
                    worksheet.Cells[row, 38].Value = Employee.August;
                    worksheet.Cells[row, 39].Value = Employee.September;
                    worksheet.Cells[row, 40].Value = Employee.October;
                    worksheet.Cells[row, 41].Value = Employee.November;
                    worksheet.Cells[row, 42].Value = Employee.December;
                    worksheet.Cells[row, 43].Value = Employee.MonthlyPrice;
                    worksheet.Cells[row, 44].Value = Employee.AltriaEXP;
                    worksheet.Cells[row, 45].Value = Employee.RoleinPOD;
                    worksheet.Cells[row, 46].Value = Employee.OverallExp;
                    worksheet.Cells[row, 47].Value = Employee.Skills;
                    worksheet.Cells[row, 48].Value = Employee.Certificates;

                }

                package.Save();
            

            return RedirectToPage("EmployeeList");
        }

        private Employee GetEmployeeById(int id)
        {
            // Fetch employee data from the Excel sheet based on the ID
            var package = new ExcelPackage(new FileInfo(employeeFilePath));
            
                var worksheet = package.Workbook.Worksheets["Employees"];
                var row = FindEmployeeRow(worksheet, id);
                if (row != -1)
                {
                    return new Employee
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
                        StartDate = ParseDate(worksheet.Cells[row, 27]?.Text),
                        EndDate = ParseDate(worksheet.Cells[row, 28]?.Text),
                        COR = worksheet.Cells[row, 29].Text,
                        Group = worksheet.Cells[row, 30].Text,
                        January = ParseDecemal(worksheet.Cells[row, 31].Text),
                        February = ParseDecemal(worksheet.Cells[row, 32].Text),
                        March = ParseDecemal(worksheet.Cells[row, 33].Text),
                        April = ParseDecemal(worksheet.Cells[row, 34].Text),
                        May = ParseDecemal(worksheet.Cells[row, 35].Text),
                        June = ParseDecemal(worksheet.Cells[row, 36].Text),
                        July = ParseDecemal(worksheet.Cells[row, 37].Text),
                        August = ParseDecemal(worksheet.Cells[row, 38].Text),
                        September = ParseDecemal(worksheet.Cells[row, 39].Text),
                        October = ParseDecemal(worksheet.Cells[row, 40].Text),
                        November = ParseDecemal(worksheet.Cells[row, 41].Text),
                        December = ParseDecemal(worksheet.Cells[row, 42].Text),
                        MonthlyPrice = ParseDecemal(worksheet.Cells[row, 43].Text),
                        AltriaEXP = ParseDecemal(worksheet.Cells[row, 44].Text),
                        RoleinPOD = worksheet.Cells[row, 45].Text,
                        OverallExp = worksheet.Cells[row, 46].Text,
                        Skills = worksheet.Cells[row, 47].Text,
                        Certificates = worksheet.Cells[row, 48].Text

                        

                    };
                }
            
            return null;
        }

        private int FindEmployeeRow(ExcelWorksheet worksheet, int id)
        {
            for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
            {
                if (worksheet.Cells[row, 1].Text == id.ToString())
                {
                    return row;
                }
            }
            return -1;
        }

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
                var package = new ExcelPackage(new FileInfo(employeeFilePath));
                
                    var worksheet = package.Workbook.Worksheets["Dropdown"]; // Ensure this matches your worksheet name
                    if (worksheet != null)
                    {
                        var rows = worksheet.Dimension.Rows;

                        for (int row = 2; row <= rows; row++)
                        {
                            var grade = worksheet.Cells[row, 1]?.Text?.Trim();
                            var bu = worksheet.Cells[row, 2]?.Text?.Trim();
                            var projectcode = worksheet.Cells[row, 3]?.Text?.Trim();
                            var projectname = worksheet.Cells[row, 4]?.Text?.Trim();
                            var PODname = worksheet.Cells[row, 5]?.Text?.Trim();
                            var offshore = worksheet.Cells[row, 6]?.Text?.Trim();
                            var type = worksheet.Cells[row, 7]?.Text?.Trim();
                            var tower = worksheet.Cells[row, 8]?.Text?.Trim();

                            if (!string.IsNullOrWhiteSpace(grade))
                            {
                                GradeOptions.Add(new SelectListItem { Value = grade, Text = grade });
                            }

                            if (!string.IsNullOrWhiteSpace(bu))
                            {
                                BUOptions.Add(new SelectListItem { Value = bu, Text = bu });
                            }
                            if (!string.IsNullOrWhiteSpace(projectcode) && !string.IsNullOrWhiteSpace(projectname))
                            {
                                ProjectCodeOptions.Add(new SelectListItem { Value = projectcode, Text = projectcode });
                                projectCodeToNameMapping.Add(projectcode,projectname);
                            }
                            if (!string.IsNullOrWhiteSpace(PODname))
                            {
                                PODNameOptions.Add(new SelectListItem { Value = PODname, Text = PODname });
                            }
                            if (!string.IsNullOrWhiteSpace(offshore))
                            {
                                OffShoreCityOptions.Add(new SelectListItem { Value = offshore, Text = offshore });
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
                        ModelState.AddModelError("", "Worksheet 'Data' not found in the dropdown file.");
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
