using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace WebApplication3
{
    public partial class _Default : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
        }

        private byte[] CreateExcelStream() 
        {
            System.IO.MemoryStream ms = new System.IO.MemoryStream();
            using (System.IO.FileStream fs = System.IO.File.OpenRead(System.AppDomain.CurrentDomain.BaseDirectory.ToString() + "../PlanilhaDeEntrada.xlsx"))
            {
                using (OfficeOpenXml.ExcelPackage excelPackage = new OfficeOpenXml.ExcelPackage(fs))
                {
                    OfficeOpenXml.ExcelWorkbook excelWorkBook = excelPackage.Workbook;
                    OfficeOpenXml.ExcelWorksheet excelWorksheet = excelWorkBook.Worksheets.First();
                    excelWorksheet.Cells[15,2].Value = "Ou 'B1'";
                    excelWorksheet.Cells[15,3].Value = "Ou 'B2'";
                    excelWorksheet.Cells[15,4].Value = "Ou 'B3'";
                    excelWorksheet.Cells[15,5].Value = "Ou 'B4'";
                    excelWorksheet.Cells[15,6].Value = "Ou 'B5'";
                    excelWorksheet.Cells[15,7].Value = "Ou 'B6'";

                    excelPackage.SaveAs(ms); // This is the important part.
                    return ms.ToArray();
                }
            }
        }

        protected void Button1_Click1(object sender, EventArgs e)
        {
            Response.Clear();
            Response.Buffer = true;
            Response.Charset = "";
            HttpResponse response = System.Web.HttpContext.Current.Response;
            Byte[] bytes = CreateExcelStream();
            response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            response.AddHeader("Content-Disposition", "attachment;filename=@PlanilhaDeSaida.xlsx");
            response.AddHeader("Content-Length", bytes.Length.ToString());
            response.BinaryWrite(bytes);
            Response.End();
        }


    }
}