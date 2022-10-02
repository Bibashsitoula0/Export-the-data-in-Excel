using Api.Models;
using AutoMapper;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using UGC.DAL.CommonRepository;
using UGC.DAL;
using UGC.Service.AccountService;
using UGC.Service.CurrentUserService;
using UGC.Service.StudentService;
using UGC.Model;
using System.Data;
using Api.Helpers;
using OfficeOpenXml;

namespace Api.Controllers
{

    [Route("api/[controller]")]
    [ApiController]
    public class StudentController : ControllerBase
    {
       private IStudentService _studentService;

        public StudentController(IStudentService studentService)
        {
            _studentService = studentService;
          
        }
        [HttpGet]
        [Route("Export")]
        public async Task<ActionResult> StudentRecordExports()
        {
            //get data from student table
            var data = await _studentService.GetStudentList();
            
           //Use the LicenseContext property on the ExcelPackage class
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage pck = new ExcelPackage();
            
            // create the row in excell 
            DataTable datatable = new DataTable();
            datatable.Columns.Add("Id");
            datatable.Columns.Add("Neb Id");
            datatable.Columns.Add("Student Name");
            datatable.Columns.Add("Mobile");
            datatable.Columns.Add("Email");
            for (int i = 0; i < data.Count; i++)
            {
                datatable.Rows.Add(
                     data[i].id,
                       data[i].neb_id,
                       data[i].student_name,
                       data[i].mobile_no,
                       data[i].email
                    );
             }
            var heading = "UGC (Student List)";
            var heading1 = "UGC (STUDENT)";
            var heading2 = "Fiscal Year :" + DateTime.Now.ToString("y");
            var heading3 = "Generated on :" + DateTime.Now.ToString("M/d/yyyy");



            byte[] filecontent = ExcelExportHelper.ExportExcel(datatable, heading, heading1, heading2, heading3, true);

            return File(filecontent, ExcelExportHelper.ExcelContentType, "Student Record.xlsx");
        }

     }
}
