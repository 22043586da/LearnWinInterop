using LearnWinInterop.Entities;
using Microsoft.Office.Interop.Word;
using System;
using System.Security.Cryptography;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using Paragraph = Microsoft.Office.Interop.Word.Paragraph;

using Excel = Microsoft.Office.Interop.Excel;

namespace LearnWinInterop
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        List<Order> _orders = new List<Order>();
        public MainWindow()
        {
            InitializeComponent();
            _orders = Entities.LaboratoryEntities.GetContext().Order.Where(x=> x.Order_Service.FirstOrDefault().Accepted == true).ToList();
        }

        private void _WordSavebtn__Click(object sender, RoutedEventArgs e)
        {
            var application = new Microsoft.Office.Interop.Word.Application();

            Document document = application.Documents.Add();

            Microsoft.Office.Interop.Word.Paragraph Zagolovock = document.Paragraphs.Add();
            Range ZagolovockRange = Zagolovock.Range;
            ZagolovockRange.Text = "Отчет по принятым анализам";
            Zagolovock.set_Style("Заголовок");
            ZagolovockRange.InsertParagraphAfter();

            Microsoft.Office.Interop.Word.Paragraph tableParagraph = document.Paragraphs.Add();
            Microsoft.Office.Interop.Word.Range tablerange = tableParagraph.Range;
            Microsoft.Office.Interop.Word.Table orderTable = document.Tables.Add(tablerange, _orders.Count() + 1, 3);

            orderTable.Borders.InsideLineStyle = orderTable.Borders.OutsideLineStyle
                = WdLineStyle.wdLineStyleSingle;
            orderTable.Range.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;

            Range cellRange;

            cellRange = orderTable.Cell(1, 1).Range;
            cellRange.Text = "Номер пробирки";
            cellRange = orderTable.Cell(1, 2).Range;
            cellRange.Text = "Дата поступления";
            cellRange = orderTable.Cell(1, 3).Range;
            cellRange.Text = "Пронализировано (да/нет)";

            orderTable.Rows[1].Range.Bold = 1;
            orderTable.Rows[1].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;

            for (int i = 0; i < _orders.Count(); i++)
            {
                var currentOrder = _orders[i];

                cellRange = orderTable.Cell(i + 2, 1).Range;
                cellRange.Text = currentOrder.ID.ToString();

                cellRange = orderTable.Cell(i + 2, 2).Range;
                cellRange.Text = currentOrder.Date_Created.ToString("dd.MM.yyyy");

                cellRange = orderTable.Cell(i + 2, 3).Range;
                cellRange.Text = ((Order_Service)currentOrder.Order_Service.First()).Finished_Date == null ? "Нет" : "Да";
            }

            Paragraph AllOrdersParagraph = document.Paragraphs.Add();
            Range AllOrdersParagraphRange = AllOrdersParagraph.Range;
            AllOrdersParagraphRange.Text = $"Количество всех заказов: {_orders.Count}";
            AllOrdersParagraphRange.InsertParagraphAfter();

            int countUnfinished = _orders.Count(s => s.Order_Service.First().Finished_Date == null);

            Paragraph UnfinishedParagraph = document.Paragraphs.Add();
            Range UnfinishedParagraphRange = UnfinishedParagraph.Range;
            UnfinishedParagraphRange.Text = $"Количество не проанализированных заказов: {countUnfinished}";
            UnfinishedParagraphRange.Font.Color = WdColor.wdColorDarkRed;
            UnfinishedParagraphRange.InsertParagraphAfter();

            int countFinished = _orders.Count() - countUnfinished;

            Paragraph FinishedParagraph = document.Paragraphs.Add();
            Range FinishedParagraphRange = FinishedParagraph.Range;
            FinishedParagraphRange.Text = $"Количество проанализированных заказов: {countFinished}";
            FinishedParagraphRange.Font.Color = WdColor.wdColorDarkGreen;
            FinishedParagraphRange.InsertParagraphAfter();

            application.Visible = true;

            var path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            document.SaveAs2(Path.Combine(path,"OrderReport.docx"));
            document.SaveAs2(Path.Combine(path,"OrdersReport.pdf"), WdExportFormat.wdExportFormatPDF);

        }

        private void _ExcelSavebtn__Click(object sender, RoutedEventArgs e)
        {
            Dictionary<DateTime, List<Order>> UnFinishedOrdersinDate = new Dictionary<DateTime, List<Order>>();
            Dictionary<DateTime, List<Order>> FinishedOrdersinDate = new Dictionary<DateTime, List<Order>>();

            Dictionary<DateTime, List<Order>> AllOrdersinDate = new Dictionary<DateTime, List<Order>>();

            var datesNewOrders = _orders.Where(x => x.Order_Service.FirstOrDefault().Finished_Date == null ).Select(s => s.Date_Created).Distinct().OrderByDescending( s => s.Date ).ToList();
            var datesFinishedOrders = _orders.Where(x => x.Order_Service.FirstOrDefault().Finished_Date != null ).Select(s => s.Order_Service.FirstOrDefault().Finished_Date).Distinct().OrderByDescending(s => s.Value.Date).ToList();

            //var allDateCreated = _orders.Select(s => s.Date_Created).Distinct().OrderByDescending(s => s.Date).ToList();

            var TempAllDateCreated = _orders.Select(s => s.Date_Created).ToList();

            var allDateCreated = new List<DateTime>();

            for (int i = 0; i < length; i++)
            {

            }


            var application = new Excel.Application();
            application.SheetsInNewWorkbook = 3;

        }
    }
}
