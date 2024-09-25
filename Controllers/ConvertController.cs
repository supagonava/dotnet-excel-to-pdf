using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Syncfusion.XlsIO;

namespace API.Controllers
{
    public class IFormExcelToPdf
    {
        public List<IFormFile> excel_file { get; set; }
    };

    [ApiController]
    [Route("convert")]
    public class ConvertController : ControllerBase
    {
        [HttpPost]
        [Route("excel-to-pdf")]
        public MemoryStream ExcelToPdf([FromForm] IFormExcelToPdf form)
        {
            var excelEngine = new ExcelEngine();
            bool exists = System.IO.Directory.Exists(
                Path.Combine(Directory.GetCurrentDirectory(), "temp")
            );
            if (!exists)
            {
                System.IO.Directory.CreateDirectory(
                    Path.Combine(Directory.GetCurrentDirectory(), "temp")
                );
            }
            var tempPath = Path.Combine(
                Directory.GetCurrentDirectory(),
                "temp",
                Guid.NewGuid().ToString() + ".xlsx"
            );
            using (var ms = new MemoryStream())
            {
                form.excel_file[0].CopyTo(ms);
                System.IO.File.WriteAllBytes(tempPath, ms.ToArray());
                ms.Close();
            }

            FileStream fileStream = new FileStream(tempPath, FileMode.Open, FileAccess.Read);

            IApplication application = excelEngine.Excel;
            IWorkbook workbook = application.Workbooks.Open(fileStream);
            IWorksheet worksheet = workbook.Worksheets[0];
            MemoryStream stream = new MemoryStream();
            workbook.SaveAs(stream);
            stream.Position = 0;

            excelEngine.Dispose();
            workbook.Close();
            fileStream.Close();

            FileInfo tempFile = new FileInfo(tempPath);
            tempFile.Delete();

            return ConvertExcelToPdf(stream);
        }

        public MemoryStream ConvertExcelToPdf(MemoryStream stream)
        {
            string fileName = Guid.NewGuid().ToString();
            string path = "template/" + fileName + ".pdf";
            Spire.Xls.Workbook workbook = new Spire.Xls.Workbook();

            workbook.LoadFromStream(stream);

            workbook.Worksheets["Evaluation Warning"].Remove();
            // workbook.ConverterSetting.SheetFitToPage = true;
            workbook.ConverterSetting.SheetFitToWidth = true;
            workbook.SaveToStream(stream, Spire.Xls.FileFormat.PDF);
            // System.IO.File.WriteAllBytes("test.pdf", stream.ToArray());
            return stream;
        }
    }
}
