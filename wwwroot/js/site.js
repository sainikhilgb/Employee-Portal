// // Please see documentation at https://learn.microsoft.com/aspnet/core/client-side/bundling-and-minification
// // for details on configuring this project to bundle and minify static web assets.

// // Write your JavaScript code.
function getDate() {
    // Get the current date
    const today = new Date();
    // Format the date as YYYY-MM-DD
    const formattedDate = today.toISOString().split('T')[0];
     // Select all elements with the class 'date-field'
     const dateFields = document.querySelectorAll('.date-field');
     // Loop through each date field and set the value
     dateFields.forEach(field => {
         field.value = formattedDate;
     });
}
function fetchProjectName() {
    var selectedProjectCode = $("#projectCode").val();
    console.log(selectedProjectCode)

  $.ajax({
    url: `Registration?handler=ProjectName&projectCode=${selectedProjectCode}`,
    type: "GET",
    success: function(projectName) {
      $("#projectName").val(projectName);
      console.log(projectName);
    },
    error: function() {
      console.error("Error fetching project name.");
    }
  });
}
window.onload = getDate;

 