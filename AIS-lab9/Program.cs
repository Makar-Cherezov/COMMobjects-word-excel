using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;

namespace AIS_lab9
{
    internal class Program
    {
        static void Replace(string toFind, string toReplace, Word.Document wordDocument)
        {
            Word.Range range = wordDocument.StoryRanges[Word.WdStoryType.wdMainTextStory];
            range.Find.ClearFormatting();
            range.Find.Execute(FindText: toFind, ReplaceWith: toReplace);
        }
        static void Main()
        {
            // Редактирование документа Word

            Word.Application wordWorker = new Word.Application();
            wordWorker.Visible = true;
            // Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            object wordFile = @"C:\Учёба\Архитектура ИС\Лр 9\AIS-lab9\AIS-lab9\example.docx";
            Word.Document wordDocument = wordWorker.Documents.Add(ref wordFile, 
                false, 
                Word.WdNewDocumentType.wdNewBlankDocument, 
                true);
            
            Random random = new Random();
            Replace("{n}", (random.Next() % 10).ToString(), wordDocument);
            Replace("{m}", (random.Next() % 30).ToString(), wordDocument);
            Replace("{название}", "The Rookie", wordDocument);

            wordDocument.Bookmarks["here"].Range.Text = "тут";

            try
            {
                wordDocument.SaveAs2(@"C:\Учёба\Архитектура ИС\Лр 9\AIS-lab9\AIS-lab9\altered_example.docx");
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }

            wordWorker.Quit(Word.WdSaveOptions.wdPromptToSaveChanges);

            // Редактирование таблицы Excel

            Excel.Application excelWorker = new Excel.Application();
            Excel.Workbook exbook = excelWorker.Workbooks.Add();
            Excel.Worksheet ws = exbook.ActiveSheet;

            for (int i = 1, num = 61; i < 10; i++, num--)
            {
                ws.Cells[1, i].Value = i;
                ws.Cells[2, i].Value = num;
                ws.Cells[1, i].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                ws.Cells[2, i].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.AliceBlue);
            }

            Excel.Range cell = ws.Cells[3, 1];
            cell.Formula = "=SUM(A1:J1)";
            cell.FormulaHidden = false;
            cell.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            Excel.ChartObject myChart = ((Excel.ChartObjects)ws.ChartObjects(Type.Missing)).Add(50, 100, 400, 300);
            Excel.Chart chart = myChart.Chart;
            chart.ChartType = Excel.XlChartType.xlXYScatterSmooth;
            Excel.Series series = ((Excel.SeriesCollection)chart.SeriesCollection(Type.Missing)).NewSeries();
            series.XValues = ws.Range["A1:J1"];
            chart.SetSourceData(ws.Range["A2:J2"]);
            chart.HasTitle = true;
            chart.ChartTitle.Text = "График из C#";
            chart.HasLegend = true;
            series.Name = "Выборка";

            try
            {
                ws.SaveAs2(@"C:\Учёба\Архитектура ИС\Лр 9\AIS-lab9\AIS-lab9\excel-exampl");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            excelWorker.Quit();

        }
    }
}