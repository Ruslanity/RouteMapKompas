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
using Kompas6Constants;
using System.Runtime.InteropServices;
using System.IO;
using System.Data.SqlClient;

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
            IPropertyMng propertyMng = (IPropertyMng)application;
            var properties = propertyMng.GetProperties(document3D);
            IPropertyKeeper propertyKeeper = (IPropertyKeeper)part;

            string partName = "";
            string partDesignation = "";
            string partMaterial = "";
            double partMass = 0;

            #region Вытаскиваем свойства
            foreach (IProperty item in properties)
            {
                if (item.Name == "Наименование")
                {
                    dynamic info;
                    bool source;
                    propertyKeeper.GetPropertyValue((_Property)item, out info, false, out source);
                    partName = info;
                    //SqlCommand cmd3 = new SqlCommand($"UPDATE [Detail] SET [FileName] = N'{info}' WHERE [Name] = N'{part.Name}'", SqlConnection);
                    //cmd3.ExecuteNonQuery();
                }
                if (item.Name == "Обозначение")
                {
                    dynamic info;
                    bool source;
                    propertyKeeper.GetPropertyValue((_Property)item, out info, false, out source);
                    partDesignation = info;
                    //SqlCommand cmd3 = new SqlCommand($"UPDATE [Detail] SET [Designation] = N'{info}' WHERE [Name] = N'{part.Name}'", SqlConnection);
                    //cmd3.ExecuteNonQuery();
                }
                if (item.Name == "Материал")
                {
                    dynamic info;
                    bool source;
                    propertyKeeper.GetPropertyValue((_Property)item, out info, false, out source);
                    partMaterial = info;
                    //SqlCommand cmd3 = new SqlCommand($"UPDATE [Detail] SET [Material] = N'{info}' WHERE [Name] = N'{part.Name}'", SqlConnection);
                    //cmd3.ExecuteNonQuery();
                }
                if (item.Name == "Масса")
                {
                    item.SignificantDigitsCount = 2;
                    dynamic info;
                    bool source;
                    propertyKeeper.GetPropertyValue((_Property)item, out info, false, out source);
                    partMass = info;
                    //SqlCommand cmd3 = new SqlCommand($"UPDATE [Detail] SET [Mass] = N'{info}' WHERE [Name] = N'{part.Name}'", SqlConnection);
                    //cmd3.ExecuteNonQuery();
                }
                if (item.Name == "Раздел спецификации")
                {
                    dynamic info;
                    bool source;
                    propertyKeeper.GetPropertyValue((_Property)item, out info, false, out source);
                    //SqlCommand cmd3 = new SqlCommand($"UPDATE [Detail] SET [Раздел спецификации] = N'{info}' WHERE [Name] = N'{part.Name}'", SqlConnection);
                    //cmd3.ExecuteNonQuery();
                }
            }
            #endregion


            string tempname = part.FileName;
            string PathName = document3D.Path;
                        
            Workbook wb = new Workbook();
            Worksheet sheet = wb.Worksheets[0];
            sheet.Name = "МК";

            Aspose.Cells.Range _range;
            Cell cell = sheet.Cells["A1"];

            #region Форматируем таблицу
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
            #endregion
            
            #region "МК"            
            worksheet.Cells[0, 8].PutValue("МК");
            Style style2 = cell.GetStyle();
            style2.Font.Name = "Arial Cyr";
            style2.Font.Size = 12;
            style2.Font.IsBold = true;
            style2.Font.IsItalic = false;
            style2.HorizontalAlignment = TextAlignmentType.Right;
            style2.VerticalAlignment = TextAlignmentType.Center;
            worksheet.Cells[0, 8].SetStyle(style2);
            _range = worksheet.Cells.CreateRange("I1");
            _range.SetOutlineBorders(CellBorderType.None, Color.Black);
            #endregion

            #region Проект№
            cells.Merge(1, 2, 1, 2);
            _range = worksheet.Cells.CreateRange("C2","D2");
            _range.SetOutlineBorders(CellBorderType.None, Color.Black);
            worksheet.Cells[1, 2].PutValue("Проект№");
            Style style4 = cell.GetStyle();
            style4.Font.Name = "Arial Cyr";
            style4.Font.Size = 10;
            style4.Font.IsBold = true;
            style4.Font.IsItalic = true;
            style4.HorizontalAlignment = TextAlignmentType.Right;
            style4.VerticalAlignment = TextAlignmentType.Center;
            worksheet.Cells[1, 2].SetStyle(style4);
            #endregion

            #region AlexLift            
            cells.Merge(0, 0, 1, 3);
            Style style = cell.GetStyle();
            style.Font.Name = "Times New Roman";
            style.Font.IsBold = true;
            style.Font.IsItalic = false;
            style.Font.Size = 26;
            style.HorizontalAlignment = TextAlignmentType.Center;
            style.VerticalAlignment = TextAlignmentType.Center;
            worksheet.Cells[0, 0].PutValue("ALEXLIFT");
            worksheet.Cells[0, 0].SetStyle(style);
            _range = worksheet.Cells.CreateRange("A1", "C1");
            _range.SetOutlineBorders(CellBorderType.Thick, Color.Black);
            #endregion

            #region Алекс-Лифт
            cell = sheet.Cells["E2"];
            cells.Merge(1, 4, 1, 2);
            Style style5 = cell.GetStyle();
            style5.Font.Name = "Arial Cyr";
            style5.Font.IsBold = true;
            style5.Font.IsItalic = true;
            style5.Font.Size = 10;
            style5.HorizontalAlignment = TextAlignmentType.Center;
            style5.VerticalAlignment = TextAlignmentType.Center;

            worksheet.Cells[1, 4].PutValue("Алекс-Лифт");
            worksheet.Cells[1, 4].SetStyle(style5);
            Aspose.Cells.Range _range4;
            _range4 = worksheet.Cells.CreateRange("E2", "F2");
            _range4.SetOutlineBorders(CellBorderType.Thin, Color.Black);
            #endregion

            #region Заносим обозначение детали
            Cell cell3 = sheet.Cells["E2"];
            cells.Merge(10, 0, 1, 3);
            Style style3 = cell3.GetStyle();
            style3.Font.Name = "Arial Cyr";
            style3.Font.Size = 12;
            style3.Font.IsBold = false;
            style3.Font.IsItalic = false;
            style3.IsTextWrapped = true;
            style3.HorizontalAlignment = TextAlignmentType.Center;
            style3.VerticalAlignment = TextAlignmentType.Center;
            worksheet.Cells[10, 0].SetStyle(style3);
            worksheet.Cells[10, 0].PutValue(partDesignation);
            _range = worksheet.Cells.CreateRange("A11", "C11");
            _range.SetOutlineBorders(CellBorderType.Thin, Color.Black);
            #endregion

            #region Заношу наименование

            cells.Merge(10, 3, 1, 3);
            style3.HorizontalAlignment = TextAlignmentType.Center;
            style3.VerticalAlignment = TextAlignmentType.Center;
            worksheet.Cells[10, 3].SetStyle(style3);
            worksheet.Cells[10, 3].PutValue(partName);
            _range = worksheet.Cells.CreateRange("D11", "F11");
            _range.SetOutlineBorders(CellBorderType.Thin, Color.Black);

            #endregion

            #region Заношу материал

            cells.Merge(10, 6, 1, 2);
            worksheet.Cells[10, 6].SetStyle(style3);
            worksheet.Cells[10, 6].PutValue(partMaterial);
            _range = worksheet.Cells.CreateRange("G11", "H11");
            _range.SetOutlineBorders(CellBorderType.Thin, Color.Black);

            #endregion

            #region Заношу массу

            worksheet.Cells[10, 8].SetStyle(style3);
            worksheet.Cells[10, 8].PutValue(Math.Round(partMass, 1));
            _range = worksheet.Cells.CreateRange("I11");
            _range.SetOutlineBorders(CellBorderType.Thin, Color.Black);

            #endregion

            #region Заношу строку детали если сборка            
            if (document3D.DocumentType == Kompas6Constants.DocumentTypeEnum.ksDocumentAssembly)
            {
                cells.Merge(12, 0, 1, 9);
                _range = worksheet.Cells.CreateRange("A13", "I13");
                _range.SetOutlineBorders(CellBorderType.Thin, Color.Black);
                cell = sheet.Cells["A13"];
                Style style6 = cell.GetStyle();
                style6.Font.IsBold = true;
                style6.Font.IsItalic = false;
                style6.HorizontalAlignment = TextAlignmentType.Center;
                style6.VerticalAlignment = TextAlignmentType.Center;
                worksheet.Cells[12, 0].SetStyle(style6);
                worksheet.Cells[12, 0].PutValue("Детали, входящие в сборку");
            }
            #endregion
            
            #region Заношу программу
            if (document3D.DocumentType == Kompas6Constants.DocumentTypeEnum.ksDocumentPart)
            {
                worksheet.Cells[17, 1].PutValue("Программа");
                IKompasDocument2D document2D = (IKompasDocument2D)application.ActiveDocument;
                IVariable7 variable7 = (IVariable7)document2D.Vari;
                //var variable = variable7.Cell
                ISheetMetalContainer sheetMetalContainer = (ISheetMetalContainer)part;                
                ISheetMetalBody sheetMetalBody = (ISheetMetalBody)sheetMetalContainer.SheetMetalBodies;
                double T = sheetMetalBody.Thickness;
                string NameProgramm = T + "mm_" + partDesignation.Remove(0, 3);
                worksheet.Cells[17, 2].PutValue(NameProgramm);
            }
            #endregion
                

            #region Обработка если сборка
            //if (document3D.DocumentType == Kompas6Constants.DocumentTypeEnum.ksDocumentAssembly)
            //{
            //    IPart7 part7 = document3D.TopPart;
            //    IParts7 collection = part7.Parts;
            //    foreach (IPart7 item in collection)
            //    {
            //        documents.Open(item.FileName, true, false);
            //        CheckDetails(app);
            //    }
            //}
            #endregion

            //cell.PutValue("Hello World!");
            wb.Save(PathName+"ExcelTest.xlsx", SaveFormat.Xlsx);
        }
    }
}
