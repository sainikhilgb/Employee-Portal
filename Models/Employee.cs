using System.ComponentModel;
using System.ComponentModel.DataAnnotations;

namespace EmployeePortal.Models
{

    public class Employee
    {
        public string Type {get;set;}
        public string Tower {get;set;}
        public string ABLGBL {get;set;}
        public string TLName {get;set;}

        public string GBL_Lead {get;set;}
        [DisplayName("Project Code")]
         public int ProjectCode {get; set;}

         [Required(ErrorMessage = "Please Select the Project name")]
        [DisplayName("Project name")]
        public string ProjectName {get; set;}
        [MaxLength(10)]
        public int PONumber {get;set;}
        [Required(ErrorMessage = "Please Select the POD name")]
        [DisplayName("POD name")]
        public string PODName {get; set;}
        public string AltriaPODOwner {get; set;}
        public string ALCSDirector {get; set;}
        [Required(ErrorMessage = "Please enter GGID")]
        public int GGID {get; set;}

        [Required(ErrorMessage = "Please enter EmployeeId")]
        public int EmpId {get; set;}

        [Required(ErrorMessage = "Please enter your first name")]
        [StringLength(50)]
        public string Resource {get; set;}

        [Required(ErrorMessage = "Please enter your email address")]
        [EmailAddress]  
        public string Email {get; set;}

        public string Grade {get; set;}
        public string GlobalGrade {get; set;}

    
        [Required(ErrorMessage = "Select Yes for tagging/working in project")]
        public string IsActiveInProject { get; set; }

        public string Gender { get; set; }
        [Required(ErrorMessage = "Please select Onshore or Offshore")]
        public string Location {get; set;}
        [DisplayName("Offshore City")]
        public string OffshoreCity {get; set;}
        [DisplayName("Offshore Backup")]
        public string OffshoreBackup {get; set;}

        public string SL {get; set;}

        public string New {get; set;}

        public string Transition {get; set;}
        
        public string BU {get; set;}
        [Required(ErrorMessage = "Please enter the date of hire")]
        [DataType(DataType.Date)]
        [DisplayName("Date Of Hire")]
         public DateTime DateOfHire {get; set;}

        [Required(ErrorMessage = "Please enter the Project Start date")]
        [DataType(DataType.Date)]
        [DisplayName("Start date")]
         public DateTime StartDate {get; set;}

        [Required(ErrorMessage = "Please enter the Project end date")]
        [DataType(DataType.Date)]
        [DisplayName("End date")]
         public DateTime EndDate {get; set;}
         [Range(0,1,ErrorMessage ="Value Should be in between 0-1")]
         public decimal January {get; set;}
         [Range(0,1,ErrorMessage ="Value Should be in between 0-1")]
         public decimal February {get; set;}
         [Range(0,1,ErrorMessage ="Value Should be in between 0-1")]
         public decimal March {get; set;}
         [Range(0,1,ErrorMessage ="Value Should be in between 0-1")]
         public decimal April {get; set;}
         [Range(0,1,ErrorMessage ="Value Should be in between 0-1")]
         public decimal May {get; set;}
         [Range(0,1,ErrorMessage ="Value Should be in between 0-1")]
         public decimal June {get; set;}
         [Range(0,1,ErrorMessage ="Value Should be in between 0-1")]
         public decimal July {get; set;}
         [Range(0,1,ErrorMessage ="Value Should be in between 0-1")]
         public decimal August {get; set;}
         [Range(0,1,ErrorMessage ="Value Should be in between 0-1")]
         public decimal September {get; set;}
         [Range(0,1,ErrorMessage ="Value Should be in between 0-1")]
         public decimal October {get; set;}
         [Range(0,1,ErrorMessage ="Value Should be in between 0-1")]
         public decimal November {get; set;}
         [Range(0,1,ErrorMessage ="Value Should be in between 0-1")]
         public decimal December {get; set;}

         public string COR {get; set;}

         public string Group {get; set;}
         [DisplayName("Monthly Price")]
        public decimal MonthlyPrice {get; set;} 
       
        [DisplayName("Altria EXP")]
        public decimal AltriaEXP {get; set;}
        [DisplayName("Role in POD")]
        public string RoleinPOD {get; set;}
        [DisplayName("Overall Exp")]
        public string OverallExp {get; set;}
        public string Skills {get; set;}

        public string Certificates {get; set;}

    }

}