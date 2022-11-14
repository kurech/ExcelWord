using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Text.Json;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace ExportImportExcel
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private const int _sheetsCount = 3;
        public MainWindow()
        {
            InitializeComponent();
        }

        private string GetHashString(string s)
        {
            byte[] bytes = Encoding.Unicode.GetBytes(s);

            MD5CryptoServiceProvider CSP = new MD5CryptoServiceProvider();
            byte[] byteHash = CSP.ComputeHash(bytes);
            string hash = "";
            foreach (byte b in byteHash)
            {
                hash += string.Format("{0:x2}", b);
            }
            return hash;
        }

        private void BnImport_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
                Title = "Выберите файл базы данных"
            };
            if(!(ofd.ShowDialog() == true))
            {
                return;
            }

            string[,] list; //for data in excel
            Excel.Application ObjWorkExcel = new Excel.Application();
            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(ofd.FileName);
            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];
            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            int _columns = (int)lastCell.Column;
            int _rows = (int)lastCell.Row;
            list = new string[_rows, _columns];
            for (int j = 0; j < _columns; j++)
            {
                for (int i = 0; i < _rows; i++)
                {
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;
                }
            }
            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();

            using (Entities entities = new Entities())
            {
                for(int i = 1; i < _rows; i++)
                {
                    entities.import.Add(new import() { role = list[i, 0], fio = list[i, 1], login = list[i, 2], password = list[i, 3] });
                }
                MessageBox.Show("Успешно!");
                entities.SaveChanges();
            }
        }

        private void BnExport_Click(object sender, RoutedEventArgs e)
        {
            List<import> staffs;

            using (Entities entities = new Entities())
            {
                staffs = entities.import.ToList();
            }

            List<string[]> RoleCategories = new List<string[]>() { //for sheets name
                new string[]{ "Администратор" },
                new string[]{ "Менеджер" },
                new string[]{ "Клиент" },
            };

            var app = new Microsoft.Office.Interop.Excel.Application();
            app.SheetsInNewWorkbook = _sheetsCount;
            Microsoft.Office.Interop.Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);

            for (int i = 0; i < _sheetsCount; i++)
            {
                int startRowIndex = 1;
                Microsoft.Office.Interop.Excel.Worksheet worksheet = app.Worksheets.Item[i + 1];
                worksheet.Name = $"Категория - {RoleCategories[i][0]}";

                Microsoft.Office.Interop.Excel.Range headerRange = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[2][1]];
                headerRange.Merge();
                headerRange.Value = $"Категория - {RoleCategories[i][0]}";
                headerRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                headerRange.Font.Bold = true;
                startRowIndex++;

                worksheet.Cells[1][startRowIndex] = "Логин";
                worksheet.Cells[2][startRowIndex] = "Пароль";

                startRowIndex++;

                foreach (import import in staffs)
                {
                    if (import.role == RoleCategories[i][0])
                    {
                        worksheet.Cells[1][startRowIndex] = import.login;
                        worksheet.Cells[2][startRowIndex] = this.GetHashString(import.password);
                        startRowIndex++;
                    }
                }

                Microsoft.Office.Interop.Excel.Range rangeBorders = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[2][startRowIndex - 1]];
                rangeBorders.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                rangeBorders.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft].LineStyle =
                rangeBorders.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeTop].LineStyle =
                rangeBorders.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                rangeBorders.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideHorizontal].LineStyle =
                rangeBorders.Borders[Microsoft.Office.Interop.Excel.XlBordersIndex.xlInsideVertical].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                worksheet.Columns.AutoFit();
            }
            app.Visible = true;      
        }

        class ServiceJson
        {
            public int Id { get; set; }
            public string NameServices { get; set; }
            public string TypeOfService { get; set; }
            public string CodeService { get; set; }
            public int Cost { get; set; }
        }

        private void BnImportJson_Click(object sender, RoutedEventArgs e)
        {
            string json = File.ReadAllText(@"C:\Users\ranel\OneDrive\Документы\1.json");
            var services = JsonSerializer.Deserialize<List<ServiceJson>>(json);

            using (Entities entities = new Entities())
            {
                foreach (ServiceJson serviceJson in services)
                {
                    try
                    {
                        entities.Service.Add(new Service()
                        {
                            NameServices = serviceJson.NameServices,
                            TypeOfService = serviceJson.TypeOfService,
                            CodeService = serviceJson.CodeService,
                            Cost = Convert.ToInt32(serviceJson.Cost)
                        });
                    }
                    catch (Exception ex) 
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
                MessageBox.Show("Успешно!");
                entities.SaveChanges();
            }
        }

        private void BnExportWord_Click(object sender, RoutedEventArgs e)
        {
            List<Service> services;

            using (Entities entities = new Entities())
            {
                services = entities.Service.ToList();
            }


            var app = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document document = app.Documents.Add();

            for (int i = 0; i < _sheetsCount; i++)
            {
                Microsoft.Office.Interop.Word.Paragraph paragraph = document.Paragraphs.Add();
                Microsoft.Office.Interop.Word.Range range = paragraph.Range;

                List<string[]> CostCategories = new List<string[]>() {
                    new string[]{ "от 0 до 350" },
                    new string[]{ "от 250 до 800" },
                    new string[]{ "от 800" },
                };

                range.Text = Convert.ToString($"Категория - {CostCategories[i][0]}");
                paragraph.set_Style("Заголовок");

                range.InsertParagraphAfter();

                var data = i == 0 ? services.Where(o => o.Cost >= 0 && o.Cost <= 350)
                        : i == 1 ? services.Where(o => o.Cost >= 250 && o.Cost <= 800)
                        : i == 2 ? services.Where(o => o.Cost >= 800) : services; //sort for task
                List<Service> currentServices = data.ToList();
                int countServiceInCategory = currentServices.Count();

                Microsoft.Office.Interop.Word.Paragraph tableParagraph = document.Paragraphs.Add();
                Microsoft.Office.Interop.Word.Range tableRange = tableParagraph.Range;
                Microsoft.Office.Interop.Word.Table servicessTable = document.Tables.Add(tableRange, countServiceInCategory + 1, 5);
                servicessTable.Borders.InsideLineStyle =
                servicessTable.Borders.OutsideLineStyle =
                Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
                servicessTable.Range.Cells.VerticalAlignment =
                Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                Microsoft.Office.Interop.Word.Range cellRange = servicessTable.Cell(1, 1).Range;
                cellRange.Text = "Id услуги";
                cellRange = servicessTable.Cell(1, 2).Range;
                cellRange.Text = "Наименование";
                cellRange = servicessTable.Cell(1, 3).Range;
                cellRange.Text = "Тип";
                cellRange = servicessTable.Cell(1, 4).Range;
                cellRange.Text = "Код";
                cellRange = servicessTable.Cell(1, 5).Range;
                cellRange.Text = "Стоимость";
                servicessTable.Rows[1].Range.Bold = 1;
                servicessTable.Rows[1].Range.ParagraphFormat.Alignment =
                Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;

                int j = 1;
                foreach (var currentService in currentServices)
                {
                    cellRange = servicessTable.Cell(j + 1, 1).Range;
                    cellRange.Text = $"{currentService.IdServices}";
                    cellRange.ParagraphFormat.Alignment =
                    Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    cellRange = servicessTable.Cell(j + 1, 2).Range;
                    cellRange.Text = currentService.NameServices;
                    cellRange.ParagraphFormat.Alignment =
                    Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    cellRange = servicessTable.Cell(j + 1, 3).Range;
                    cellRange.Text = currentService.TypeOfService;
                    cellRange.ParagraphFormat.Alignment =
                    Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    cellRange = servicessTable.Cell(j + 1, 4).Range;
                    cellRange.Text = currentService.CodeService;
                    cellRange.ParagraphFormat.Alignment =
                    Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    cellRange = servicessTable.Cell(j + 1, 5).Range;
                    cellRange.Text = $"{currentService.Cost}";
                    cellRange.ParagraphFormat.Alignment =
                    Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    j++;
                }
            }

            document.Words.Last.InsertBreak(Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak);
            app.Visible = true;
        }
    }
}
