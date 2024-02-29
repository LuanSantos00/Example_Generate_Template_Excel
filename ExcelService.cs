using ClosedXML.Excel;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeOpenXml;
using System.Xml.Linq;

namespace TesteExcel
{
    public sealed class ExcelService
    {
        public Task Handle()
        {
            GenerateFile("Despesas.xlsx");
            return Task.CompletedTask;
        }

        private void GenerateFile(string filename)
        {
            var filePathName = System.IO.Directory.GetCurrentDirectory() + "\\" + filename;

            if (File.Exists(filePathName))
                File.Delete(filePathName);

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("tabelaExemplo");
                var line = 1;
                GenerateHeader(worksheet);
                line++;

                workbook.SaveAs(filePathName);
            }
        }

        private void GenerateHeader(IXLWorksheet planilha)
        {
            planilha.Cell("A1").Value = "Document";
            planilha.Cell("B1").Value = "Name";
            planilha.Cell("C1").Value = "Email";
            planilha.Cell("D1").Value = "SurName";
            planilha.Cell("E1").Value = "CredentialName";
            planilha.Cell("F1").Value = "CompanyDocument";
            planilha.Cell("G1").Value = "TradingName";
        }
    }
}
