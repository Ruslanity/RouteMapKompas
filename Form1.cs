using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using KompasAPI7;
using System.Runtime.InteropServices;
using System.IO;

namespace RouteMapKompas
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            IApplication application = (IApplication)Marshal.GetActiveObject("Kompas.Application.7");
            IKompasDocument3D document3D = (IKompasDocument3D)application.ActiveDocument;
            IPart7 part = document3D.TopPart;

            string tempname = part.FileName;
            string PathName = document3D.Path;


            Workbook wb = new Workbook();
            Worksheet sheet = wb.Worksheets[0];
            sheet.Name = "МК";
            Cell cell = sheet.Cells["A1"];

            //задаю ширину колонок
            Worksheet worksheet = wb.Worksheets[0];
            Cells cells = worksheet.Cells;
            cells.SetColumnWidth(0, 4.14);
            cells.SetColumnWidth(1, 12.29);
            cells.SetColumnWidth(2, 15.29);
            cells.SetColumnWidth(3, 7.57);
            cells.SetColumnWidth(4, 5);
            cells.SetColumnWidth(5, 8.71);
            cells.SetColumnWidth(6, 12.43);
            cells.SetColumnWidth(7, 14.86);
            cells.SetColumnWidth(8, 8);


            cells.SetRowHeight(0, 27.75);
            cells.SetRowHeight(1, 18.75);
            cells.SetRowHeight(2, 10.5);
            cells.SetRowHeight(3, 14.25);
            cells.SetRowHeight(4, 12);
            cells.SetRowHeight(5, 18.75);
            cells.SetRowHeight(6, 21);
            cells.SetRowHeight(7, 15.75);
            cells.SetRowHeight(8, 9.75);
            cells.SetRowHeight(9, 12.75);
            cells.SetRowHeight(10, 31.5);

            //Вношу запись "ЛСУ-Трейд"
            cells.Merge(0, 0, 1, 3);
            Style style = cell.GetStyle();
            style.Font.Name = "Times New Roman";
            style.Font.IsBold = true;
            style.Font.IsItalic = true;
            style.Font.Size = 26;
            style.HorizontalAlignment = TextAlignmentType.Center;
            style.VerticalAlignment = TextAlignmentType.Center;

            worksheet.Cells[0, 0].PutValue("ЛСУ-Трейд");
            worksheet.Cells[0, 0].SetStyle(style);

            //Выполняю обводку ячейки
            Aspose.Cells.Range _range;
            //worksheet.Cells[1, 0].PutValue("Hair Lines");
            _range = worksheet.Cells.CreateRange("A1", "C1");
            _range.SetOutlineBorders(CellBorderType.Thick, Color.Black);

            //Вношу запись "МК"
            _range.SetOutlineBorders(CellBorderType.None, Color.Black);
            worksheet.Cells[0, 8].PutValue("МК");
            Style style2 = cell.GetStyle();
            style2.Font.Name = "Arial Cyr";
            style2.Font.Size = 12;
            style2.Font.IsBold = true;
            style2.Font.IsItalic = false;
            style2.HorizontalAlignment = TextAlignmentType.Right;
            style2.VerticalAlignment = TextAlignmentType.Center;
            worksheet.Cells[0, 8].SetStyle(style2);

            //Вношу запись обоначение детали
            cells.Merge(10, 0, 1, 3);
            
            Style style3 = cell.GetStyle();
            style3.Font.Name = "Arial Cyr";
            style3.Font.Size = 12;
            style3.Font.IsBold = false;
            style3.Font.IsItalic = false;
            style3.HorizontalAlignment = TextAlignmentType.Center;
            style3.VerticalAlignment = TextAlignmentType.Center;
            worksheet.Cells[10, 0].SetStyle(style3);
            worksheet.Cells[10, 0].PutValue(part.Marking);
            Range _range2;
            _range2 = worksheet.Cells.CreateRange("A11", "C11");
            _range2.SetOutlineBorders(CellBorderType.Thin, Color.Black);


            //Вношу запись наименование детали
            cells.Merge(10, 3, 1, 3);
            
            style3.HorizontalAlignment = TextAlignmentType.Center;
            style3.VerticalAlignment = TextAlignmentType.Center;
            worksheet.Cells[10, 3].SetStyle(style3);
            worksheet.Cells[10, 3].PutValue(part.Name);
            Range _range3;
            _range3 = worksheet.Cells.CreateRange("D11", "F11");
            _range3.SetOutlineBorders(CellBorderType.Thin, Color.Black);


            //cell.PutValue("Hello World!");
            wb.Save(PathName+"ExcelTest.xlsx", SaveFormat.Xlsx);
        }
    }
}
