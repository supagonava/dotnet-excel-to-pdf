using System;
using System.IO;
using Microsoft.AspNetCore.Mvc;
using Syncfusion.XlsIO;

namespace API.Controllers
{
    public class ExcelToPdfRequest
    {
        public string excelPath { get; set; }
        public string outputAsExcelPath { get; set; }
    }

    [ApiController]
    [Route("convert-by-path")]
    public class ExcelToPdfByPathController : ControllerBase
    {
        [HttpPost]
        [Route("excel-to-pdf")]
        public IActionResult ExcelToPdf([FromBody] ExcelToPdfRequest request)
        {
            try
            {
                if (!System.IO.File.Exists(request.excelPath))
                {
                    return BadRequest("Input Excel file not found.");
                }

                using (var excelEngine = new ExcelEngine())
                {
                    using (
                        FileStream fileStream = new FileStream(
                            request.excelPath,
                            FileMode.Open,
                            FileAccess.Read
                        )
                    )
                    {
                        IApplication application = excelEngine.Excel;
                        IWorkbook workbook = application.Workbooks.Open(fileStream);
                        IWorksheet worksheet = workbook.Worksheets[0];

                        using (MemoryStream stream = new MemoryStream())
                        {
                            workbook.SaveAs(stream);
                            stream.Position = 0;

                            // Convert to PDF
                            String pdfStream = ConvertExcelToPdf(stream, request.outputAsExcelPath);

                            // Clean up
                            workbook.Close();
                        }
                    }
                }
                return Ok("{\"status\":true}");
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"An error occurred: {ex.Message}\n{ex.StackTrace}");
            }
        }

        public String ConvertExcelToPdf(MemoryStream stream, string outputAsExcelPath)
        {
            // MemoryStream pdfStream = new MemoryStream();

            Spire.Xls.Workbook workbook = new Spire.Xls.Workbook();
            workbook.LoadFromStream(stream);

            // Remove "Evaluation Warning" sheet if it exists
            var sheet = workbook.Worksheets["Evaluation Warning"];
            if (sheet != null)
            {
                workbook.Worksheets["Evaluation Warning"].Remove();
            }

            workbook.ConverterSetting.SheetFitToWidth = true;

            // workbook.SaveToStream(pdfStream, Spire.Xls.FileFormat.PDF);
            workbook.SaveToFile(outputAsExcelPath);

            // pdfStream.Position = 0;
            // return pdfStream;
            return outputAsExcelPath;
        }
    }
}
