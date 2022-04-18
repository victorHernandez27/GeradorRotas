using ClosedXML.Excel;
using GeradorRotas.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace GeradorRotas.Controllers
{
    public class ListarCamposPlanilhaController : Controller
    {
        [HttpGet]
        public IActionResult Index()
        {
            var colunas = new List<SelectListItem>();

            var xls = new XLWorkbook(@"Data\Planilhas\GeradorRotas.xlsx");
            //var planilha = xls.Worksheets[0];
            var planilha = xls.Worksheets.First(w => w.Name == "Planilha1");
            var totalColunas = planilha.Columns().Count();

            for (int i = 1; i <= totalColunas; i++)
            {
                SelectListItem item = new SelectListItem()
                {
                    Text = planilha.Cell(1, i).Value.ToString(),
                    Value = planilha.Cell(1, i).Value.ToString(),
                    Selected = false
                };
                colunas.Add(item);
            };

            ColunaExcel colunaExcelViewModel = new ColunaExcel();
            colunaExcelViewModel.Colunas = colunas;
            return View(colunaExcelViewModel);
        }

        [HttpPost]
        public string Index(IEnumerable<string> colunasSelecionada)
        {
            if (colunasSelecionada == null)
            {
                return "É necessário selecionar um ou mais campos";
            }
            else
            {
                //StringBuilder sb = new StringBuilder();
                //sb.Append("Selected employeeID " + string.Join(" ,", colunasSelecionada));
                //return sb.ToString();

                var xls = new XLWorkbook(@"Data\Planilhas\GeradorRotas.xlsx");
                //var planilha = xls.Worksheets[0];
                var planilha = xls.Worksheets.First(w => w.Name == "Planilha1");
                var totalLinhas = planilha.Rows().Count();

                var colunasSelecionadaList = colunasSelecionada.ToList();

                var range = planilha.RangeUsed();
                var table = range.AsTable();

                // Create a document by supplying the filepath.
                using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(@"Data\Documentos\word.docx", WordprocessingDocumentType.Document))
                {
                    // Add a main document part.
                    var mainPart = wordDocument.AddMainDocumentPart();

                    // Create the document structure and add some text.
                    mainPart.Document = new Document();
                    var body = mainPart.Document.AppendChild(new Body());
                                       

                    for (int l = 2; l <= totalLinhas; l++)
                    {
                        var paragrafo = body.AppendChild(new Paragraph());
                        var run = paragrafo.AppendChild(new Run());

                        for (int idxColuna = 0; idxColuna < colunasSelecionada.Count(); idxColuna++)
                        {
                            var cell = table.FindColumn(c => c.FirstCell().Value.ToString() == colunasSelecionadaList[idxColuna]);
                            var columnLetter = cell.RangeAddress.FirstAddress.ColumnLetter;
                            var textoCelulaExcel = planilha.Cell(l, columnLetter).Value.ToString();
                            run.AppendChild(new Text($"{textoCelulaExcel}  |  "));
                        }
                    }
                }

                return "Gerou";
            }
        }
    }
}
