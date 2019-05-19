using GemBox.Document;
using GemBox.Spreadsheet;
using System;
using System.IO;
using System.Web.Http;

namespace ossOffice.Controllers
{
    public class OfficeController : ApiController
    {
        [HttpPost]
        public ByteArrayResult ExcelToPdf([FromBody] OfficeParam par)
        {
            var result = new ByteArrayResult();

            try
            {
                if (par == null)
                    throw new ArgumentNullException(nameof(par));
                if (par.Bytes == null)
                    throw new ArgumentNullException(nameof(par.Bytes));
                if (par.Ext == null)
                    throw new ArgumentNullException(nameof(par.Ext));

                SpreadsheetInfo.SetLicense("ERDD-TN5J-YKX9-H1KX");
                ExcelFile excel;

                using (var msExcel = new MemoryStream())
                {
                    msExcel.Write(par.Bytes, 0, par.Bytes.Length);
                    msExcel.Position = 0;
                    if (par.Ext.ToLower() == ".xls")
                        excel = ExcelFile.Load(msExcel, GemBox.Spreadsheet.LoadOptions.XlsDefault);
                    else
                        excel = ExcelFile.Load(msExcel, GemBox.Spreadsheet.LoadOptions.XlsxDefault);
                }
                using (var smPdf = new MemoryStream())
                {
                    excel.Save(smPdf, GemBox.Spreadsheet.SaveOptions.PdfDefault);
                    result.Result = smPdf.ToArray();
                }
            }
            catch (Exception ex)
            {
                result.Error = ex.InmostMessage();
            }

            return result;
        }

        [HttpPost]
        public ByteArrayResult WordToPdf([FromBody] OfficeParam par)
        {
            var result = new ByteArrayResult();

            try
            {
                if (par == null)
                    throw new ArgumentNullException(nameof(par));
                if (par.Bytes == null)
                    throw new ArgumentNullException(nameof(par.Bytes));
                if (par.Ext == null)
                    throw new ArgumentNullException(nameof(par.Ext));

                ComponentInfo.SetLicense("D873-P6FI-T8I0-4WLW");
                DocumentModel word;

                using (var msWord = new MemoryStream())
                {
                    msWord.Write(par.Bytes, 0, par.Bytes.Length);
                    msWord.Position = 0;
                    if (par.Ext.ToLower() == ".doc")
                        word = DocumentModel.Load(msWord, GemBox.Document.LoadOptions.DocDefault);
                    else
                        word = DocumentModel.Load(msWord, GemBox.Document.LoadOptions.DocxDefault);
                }
                using (var smPdf = new MemoryStream())
                {
                    word.Save(smPdf, GemBox.Document.SaveOptions.PdfDefault);
                    result.Result = smPdf.ToArray();
                }
            }
            catch (Exception ex)
            {
                result.Error = ex.InmostMessage();
            }

            return result;
        }
    }
}
