using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using EmployeePortal.Models;
using Microsoft.AspNetCore.Mvc.Rendering;
using OfficeOpenXml;
using System.Numerics;


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
        public List<SelectListItem> OffshoreCityOptions { get; set; }

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

            using (var package = new ExcelPackage(new FileInfo(employeeFilePath)))
            {
                var worksheet = package.Workbook.Worksheets["Employees"];

                var row = FindEmployeeRow(worksheet, id);
                if (row != -1)
                {
                    worksheet.Cells[row, 2].Value = Employee.FirstName;
                    worksheet.Cells[row, 3].Value = Employee.LastName;
                    worksheet.Cells[row, 4].Value = Employee.Email;
                    worksheet.Cells[row, 5].Value = Employee.Phone;
                    worksheet.Cells[row, 6].Value = Employee.Grade;
                    worksheet.Cells[row, 7].Value = Employee.BU;
                    worksheet.Cells[row, 8].Value = Employee.DateOfHire.ToShortDateString();
                    worksheet.Cells[row + 1, 9].Value = Employee.ProjectCode;
                    worksheet.Cells[row + 1, 10].Value = Employee.ProjectName;
                    worksheet.Cells[row + 1, 11].Value = Employee.PODName;
                    worksheet.Cells[row + 1, 12].Value = Employee.StartDate.ToString("yyyy-MM-dd");
                    worksheet.Cells[row + 1, 13].Value = Employee.EndDate.ToString("yyyy-MM-dd");
                    worksheet.Cells[row + 1, 14].Value = Employee.Location;
                    worksheet.Cells[row + 1, 15].Value = Employee.OffshoreCity;
                }

                package.Save();
            }

            return RedirectToPage("EmployeeList");
        }

        private Employee GetEmployeeById(int id)
        {
            // Fetch employee data from the Excel sheet based on the ID
            using (var package = new ExcelPackage(new FileInfo(employeeFilePath)))
            {
                var worksheet = package.Workbook.Worksheets["Employees"];
                var row = FindEmployeeRow(worksheet, id);
                if (row != -1)
                {
                    return new Employee
                    {
                        EmployeeId = id,
                        FirstName = worksheet.Cells[row, 2].Text,
                        LastName = worksheet.Cells[row, 3].Text,
                        Email = worksheet.Cells[row, 4].Text,
                        Phone = worksheet.Cells[row, 5].Text,
                        Grade = worksheet.Cells[row, 6].Text,
                        BU = worksheet.Cells[row, 7].Text,
                        DateOfHire = DateTime.Parse(worksheet.Cells[row, 8].Text),
                        ProjectCode =int.Parse( worksheet.Cells[row, 9].Text),
                        ProjectName = worksheet.Cells[row, 10].Text,
                        PODName = worksheet.Cells[row, 11].Text,
                        StartDate = DateTime.Parse(worksheet.Cells[row, 12].Text),
                        EndDate = DateTime.Parse(worksheet.Cells[row, 13].Text),
                        Location = worksheet.Cells[row, 14].Text,
                        OffshoreCity = worksheet.Cells[row, 15].Text,

                    };
                }
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
            OffshoreCityOptions = new List<SelectListItem>();

            if (System.IO.File.Exists(employeeFilePath))
            {
                using (var package = new ExcelPackage(new FileInfo(employeeFilePath)))
                {
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
                            if (!string.IsNullOrWhiteSpace(offshore))
                            {
                                OffshoreCityOptions.Add(new SelectListItem { Value = offshore, Text = offshore });
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
