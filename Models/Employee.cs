using System.ComponentModel;
using System.ComponentModel.DataAnnotations;

namespace EmployeePortal.Models
{

    public class Employee
    {
        [Key]       
        public int EmployeeId {get; set;}
        [Required(ErrorMessage = "Please enter your first name")]
        [StringLength(50)]
        [DisplayName("First name")]
        public string FirstName {get; set;} 
         [Required(ErrorMessage = "Please enter your Last name")]
        [StringLength(50)]
        [DisplayName("Last name")]
        public string LastName {get; set;}
        [Required(ErrorMessage = "Please enter your email address")]
        [EmailAddress]  
        public string Email {get; set;}
        [Required(ErrorMessage = "Please enter your phone number")]
        [Phone]
        public string Phone {get; set;}
        public string Grade {get; set;}
        public string BU {get; set;}
        [Required(ErrorMessage = "Please enter the date of hire")]
        [DataType(DataType.Date)]
        [DisplayName("Date Of Hire")]
         public DateTime DateOfHire {get; set;}
         [DisplayName("Project Code")]
         public int ProjectCode {get; set;}

         [Required(ErrorMessage = "Please Select the Project name")]
        [DisplayName("Project name")]
        public string ProjectName {get; set;}

         [Required(ErrorMessage = "Please Select the POD name")]
        [DisplayName("POD name")]
        public string PODName {get; set;}

        [Required(ErrorMessage = "Please enter the Project Start date")]
        [DataType(DataType.Date)]
        [DisplayName("Start date")]
         public DateTime StartDate {get; set;}

         [Required(ErrorMessage = "Please enter the Project end date")]
        [DataType(DataType.Date)]
        [DisplayName("End date")]
         public DateTime EndDate {get; set;}
         [Required(ErrorMessage = "Please select Onshore or Offshore")]
         public string Location {get; set;}
         [DisplayName("Offshore City")]
         public string OffshoreCity {get; set;}
         public int Jan {get; set;}
    }

}