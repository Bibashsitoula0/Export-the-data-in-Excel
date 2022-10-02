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
        [Route("GetAllStudent")]
        public async Task<IActionResult>GetAllStudent(int? pageNumber,int? pageSize)
        {
            try
            {
                if (pageNumber == null || pageSize==null )

                {
                     pageNumber = 1;
                     pageSize = 2;
                    var data = await _studentService.GetStudent(pageNumber, pageSize);
                   
                    return Ok(new ResponseModel
                    {
                        success = true,
                        message = "success",
                        data = data
                    });
                }
                else
                {
                    var data = await _studentService.GetStudent(pageNumber, pageSize);
                    return Ok(new ResponseModel
                    {
                        success = true,
                        message = "success",
                        data = data
                    });
                }                
            }
            catch (Exception ex)
            {
                return Ok(new ResponseModel { success = false, message = ex.Message });
            }
        }
   
        [HttpGet]
        [Route("GetStudentById")]
        public async Task<IActionResult> GetStudentById(int Id)
        {
            try
            {
                var data = await _studentService.GetStudentById(Id);
                return Ok(new ResponseModel
                {
                    statuscode=1,
                    success = true,
                    message = "success",
                    data = data
                });
            }
            catch (Exception ex)
            {
                return Ok(new ResponseModel { success = false, message = ex.Message });
            }
        }

        [HttpPost]
        [Route("CreateStudent")]
        public async Task<IActionResult> CreateStudent(StudentVm model)
        {
            try
            {
                var data = new StudentNeb()
                {
                    mobile_no = model.mobile_no,
                    neb_id = model.neb_id,
                    student_name = model.student_name

                };
                var result = await _studentService.Create(data);
                return Ok(new ResponseModel
                {
                    statuscode = 1,
                    success = true,
                    message = "New Student Created",
                    data = data
                });
            }
            catch (Exception ex)
            {
                return Ok(new ResponseModel { success = false, message = ex.Message });
            }
        }

        [HttpPut]
        [Route("UpdateStudent")]
        public async Task<IActionResult> UpdateStudent(StudentNeb model)
        {
            try
            {
                var result = await _studentService.UpdateStudent(model);
                return Ok(new ResponseModel
                {
                    statuscode = 1,
                    success = true,
                    message = "Student Data Updated",
                    data = result
                });
            }
            catch (Exception ex)
            {
                return Ok(new ResponseModel { success = false, message = ex.Message });
            }
        }

        [HttpDelete]
        [Route("DeleteStudent")]
        public async Task<IActionResult> DeleteStudent(StudentNeb model)
        {
            try
            {
                 var result = await _studentService.DeleteStudent(model);
                return Ok(new ResponseModel
                {
                    statuscode = 1,
                    success = true,
                    message = $"The {model.student_name} student is deleted ",
                   data = result
                });
            }
            catch (Exception ex)
            {
                return Ok(new ResponseModel { success = false, message = ex.Message });
            }
        }

        [HttpGet]
        [Route("Export")]
        public async Task<ActionResult> StudentRecordExports()
        {
            var data = await _studentService.GetStudentListInExcel();
            DataTable datatable = new DataTable();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage pck = new ExcelPackage();
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

            return File(filecontent, ExcelExportHelper.ExcelContentType, "ApprenticeShip Training Record.xlsx");
        }

     }
}
