using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using static OpenXML_Driver.PrintExcel;

namespace OpenXML_Driver
{
    public class ReportRead
    {
        public List<string> Merges = new List<string> { };
        public ExcelCellColection ExcelCells = new ExcelCellColection();
        public ExcelFontColection ExcelFonts = new ExcelFontColection();
        public ExcelFillColection ExcelFills = new ExcelFillColection();
        public ExcelBorderColection ExcelBorders = new ExcelBorderColection();
        public ExcelCellFormatColection ExcelFormats = new ExcelCellFormatColection();

        void GetFonts(Stylesheet styleSheet)
        {
            Fonts fonts = styleSheet.Fonts;
            foreach (var f in fonts)
            {
                FontSize sz = null;
                FontName fn = null;
                Color rgb = null;
                Bold b = null;
                Italic i = null;
                Underline u = null;
                Strike s = null;

                try { sz = f.Elements<FontSize>().First(); } catch { }
                try { fn = f.Elements<FontName>().First(); } catch { }
                try { rgb = f.Elements<Color>().First(); } catch { }
                try { b = f.Elements<Bold>().First(); } catch { }
                try { i = f.Elements<Italic>().First(); } catch { }
                try { u = f.Elements<Underline>().First(); } catch { }
                try { s = f.Elements<Strike>().First(); } catch { }

                ExcelFont item = new ExcelFont();
                if (sz != null && sz.Val != null)
                    item.size = sz.Val.Value;
                if (fn != null && fn.Val != null)
                    item.fontname = fn.Val.Value;
                if (rgb != null && rgb.Rgb != null)
                    item.hexColor = rgb.Rgb.Value;

                List<FontStyle> fSt = new List<FontStyle>();

                if (b != null)
                    fSt.Add(FontStyle.Bold);
                if (i != null)
                    fSt.Add(FontStyle.Italic);
                if (u != null)
                    fSt.Add(FontStyle.Underline);
                if (s != null)
                    fSt.Add(FontStyle.Strike);

                item.Styl = fSt.ToArray();
                ExcelFonts.Add(item);
            }
        }
        void GetFill(Stylesheet styleSheet)
        {
            Fills fills = styleSheet.Fills;

            foreach (var fill in fills)
            {
                try
                {
                    var pattern = fill.Elements<PatternFill>().First();

                    ForegroundColor rgb = null;
                    try { rgb = pattern.Elements<ForegroundColor>().First(); } catch { }

                    ExcelFill item = new ExcelFill();
                    if (rgb != null && rgb.Rgb != null)
                        item.hexColor = rgb.Rgb.Value;

                    ExcelFills.Add(item);
                }
                catch
                { }
            }
        }
        void GetBorder(Stylesheet styleSheet)
        {
            Borders borders = styleSheet.Borders;
            foreach (var border in borders)
            {
                LeftBorder LeftBorder = null;
                RightBorder RightBorder = null;
                TopBorder TopBorder = null;
                BottomBorder BottomBorder = null;

                try { LeftBorder = border.Elements<LeftBorder>().First(); } catch { }
                try { RightBorder = border.Elements<RightBorder>().First(); } catch { }
                try { TopBorder = border.Elements<TopBorder>().First(); } catch { }
                try { BottomBorder = border.Elements<BottomBorder>().First(); } catch { }


                string color = "";
                Thickness thickness = new Thickness(0, 0, 0, 0);
                BorderStyleValues styl = new BorderStyleValues();
                BorderStyl Styl = new BorderStyl();

                if (LeftBorder != null)
                {
                    thickness.l = 1;
                    if (LeftBorder.Color != null && LeftBorder.Color.Rgb != null) color = LeftBorder.Color.Rgb.Value;
                    if (LeftBorder.Style != null) styl = LeftBorder.Style;
                }

                if (RightBorder != null)
                {
                    thickness.r = 1;
                    if (RightBorder.Color != null && RightBorder.Color.Rgb != null) color = RightBorder.Color.Rgb.Value;
                    if (RightBorder.Style != null) styl = RightBorder.Style;
                }

                if (TopBorder != null)
                {
                    thickness.t = 1;
                    if (TopBorder.Color != null && TopBorder.Color.Rgb != null) color = TopBorder.Color.Rgb.Value;
                    if (TopBorder.Style != null) styl = TopBorder.Style;
                }

                if (BottomBorder != null)
                {
                    thickness.b = 1;
                    if (BottomBorder.Color != null && BottomBorder.Color.Rgb != null) color = BottomBorder.Color.Rgb.Value;
                    if (BottomBorder.Style != null) styl = BottomBorder.Style;
                }

                switch (styl)
                {
                    case BorderStyleValues.DashDot:
                        Styl = BorderStyl.DashDot;
                        break;
                    case BorderStyleValues.DashDotDot:
                        Styl = BorderStyl.DashDotDot;
                        break;
                    case BorderStyleValues.Dashed:
                        Styl = BorderStyl.Dashed;
                        break;
                    case BorderStyleValues.Dotted:
                        Styl = BorderStyl.Dotted;
                        break;
                    case BorderStyleValues.Double:
                        Styl = BorderStyl.Double;
                        break;
                    case BorderStyleValues.Hair:
                        Styl = BorderStyl.Hair;
                        break;
                    case BorderStyleValues.Medium:
                        Styl = BorderStyl.Medium;
                        break;
                    case BorderStyleValues.MediumDashDot:
                        Styl = BorderStyl.MediumDashDot;
                        break;
                    case BorderStyleValues.MediumDashDotDot:
                        Styl = BorderStyl.MediumDashDotDot;
                        break;
                    case BorderStyleValues.MediumDashed:
                        Styl = BorderStyl.MediumDashed;
                        break;
                    case BorderStyleValues.None:
                        Styl = BorderStyl.None;
                        break;
                    case BorderStyleValues.SlantDashDot:
                        Styl = BorderStyl.SlantDashDot;
                        break;
                    case BorderStyleValues.Thick:
                        Styl = BorderStyl.Thick;
                        break;
                    case BorderStyleValues.Thin:
                        Styl = BorderStyl.Thin;
                        break;
                }


                ExcelBorder item = new ExcelBorder();
                item.borderhexColor = color;
                item.BorderThickness = thickness;
                item.Styl = Styl;

                ExcelBorders.Add(item);
            }
        }
        void GetStyleFormat(Stylesheet styleSheet)
        {
            CellFormats formats = styleSheet.CellFormats;
            foreach (CellFormat cell in formats)
            {
                ExcelCellFormat item = new ExcelCellFormat();
                item.borderID = cell.BorderId;
                item.fillID = cell.FillId;
                item.fontID = cell.FontId;

                ExcelAlignment Alignment_ = new ExcelAlignment();
                Alignment alignment = cell.Alignment;

                if (alignment != null)
                {
                    if (alignment.Horizontal != null)
                    {
                        switch (alignment.Horizontal.Value)
                        {
                            case HorizontalAlignmentValues.Center:
                                Alignment_.HAligment = HorizontalAligment.Center;
                                break;
                            case HorizontalAlignmentValues.CenterContinuous:
                                Alignment_.HAligment = HorizontalAligment.CenterContinuous;
                                break;
                            case HorizontalAlignmentValues.Distributed:
                                Alignment_.HAligment = HorizontalAligment.Distributed;
                                break;
                            case HorizontalAlignmentValues.Fill:
                                Alignment_.HAligment = HorizontalAligment.Fill;
                                break;
                            case HorizontalAlignmentValues.General:
                                Alignment_.HAligment = HorizontalAligment.General;
                                break;
                            case HorizontalAlignmentValues.Justify:
                                Alignment_.HAligment = HorizontalAligment.Justify;
                                break;
                            case HorizontalAlignmentValues.Left:
                                Alignment_.HAligment = HorizontalAligment.Left;
                                break;
                            case HorizontalAlignmentValues.Right:
                                Alignment_.HAligment = HorizontalAligment.Right;
                                break;
                        }
                    }

                    if (alignment.Vertical != null)
                        switch (alignment.Vertical.Value)
                        {
                            case VerticalAlignmentValues.Bottom:
                                Alignment_.VAligment = VerticalAligment.Bottom;
                                break;
                            case VerticalAlignmentValues.Center:
                                Alignment_.VAligment = VerticalAligment.Center;
                                break;
                            case VerticalAlignmentValues.Distributed:
                                Alignment_.VAligment = VerticalAligment.Distributed;
                                break;
                            case VerticalAlignmentValues.Justify:
                                Alignment_.VAligment = VerticalAligment.Justify;
                                break;
                            case VerticalAlignmentValues.Top:
                                Alignment_.VAligment = VerticalAligment.Top;
                                break;
                        }

                    if (alignment.WrapText != null)
                    {
                        Alignment_.Wrap = alignment.WrapText.Value;
                    }
                }

                item.Alignment = Alignment_;
                ExcelFormats.Add(item);
            }

        }
        void GetCell(Cell c, SharedStringTable sst, bool kind)
        {
            ExcelCell Cell = new ExcelCell();

            var find = Merges.Where(x => x.Contains(c.CellReference.Value)).ToList();

            if (find.Any())
            {
                Cell.Zakres = find.First();
                Merges.Remove(find.First());
            }
            else
                Cell.Zakres = c.CellReference.Value;


            if (Cell.Zakres.Contains(':'))
            {
                string[] Ranges = Cell.Zakres.Split(':');
                Cell.Coordinates.Add(ReportHelper.getxy(Ranges[0]));
                Cell.Coordinates.Add(ReportHelper.getxy(Ranges[1]));
            }
            else
            {
                Cell.Coordinates.Add(ReportHelper.getxy(Cell.Zakres));
            }

            if (c.CellFormula != null && c.CellFormula.Text != null)
                Cell.Formula = c.CellFormula.Text;

            if (kind)
            {
                if (c.CellValue != null && c.CellValue.Text != null)
                {
                    int ssid = int.Parse(c.CellValue.Text);
                    string str = sst.ChildElements[ssid].InnerText;

                    Cell.Value = str;
                }
            }
            else
            {
                if (c.CellValue != null && c.CellValue.Text != null)
                    Cell.Value = c.CellValue.Text;
            }

            if (c.DataType != null && c.DataType.Value != null)
                switch (c.DataType.Value)
                {
                    case CellValues.Date:
                        Cell.Typ = Arkusz.CellType.Date;
                        break;
                    case CellValues.Number:
                        Cell.Typ = Arkusz.CellType.Number;
                        break;
                    case CellValues.String:
                        Cell.Typ = Arkusz.CellType.String;
                        break;
                }

            if (c.StyleIndex != null)
                Cell.Styl = c.StyleIndex;
            ExcelCells.Add(Cell);
        }

        void Excel03to07(string fileName)
        {
            if (File.Exists(Path.GetTempPath() + "EMMA2\\PEARR\\temporary.xlsm"))
                File.Delete(Path.GetTempPath() + "EMMA2\\PEARR\\temporary.xlsm");

            if (!Directory.Exists(Path.GetTempPath() + "EMMA2"))
                Directory.CreateDirectory(Path.GetTempPath() + "EMMA2");
            if (!Directory.Exists(Path.GetTempPath() + "EMMA2\\PEARR"))
                Directory.CreateDirectory(Path.GetTempPath() + "EMMA2\\PEARR");

            string svfileName = Path.GetTempPath() + "EMMA2\\PEARR\\temporary.xlsm";

            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false;
            excelApp.DisplayAlerts = false;

            Microsoft.Office.Interop.Excel.Workbook eWorkbook = excelApp.Workbooks.Open(fileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            eWorkbook.SaveAs(svfileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            eWorkbook.Close(false, Type.Missing, Type.Missing);

            excelApp.Quit(); 
        }

        WorksheetPart GetWorksheetFromSheetName(WorkbookPart workbookPart, string sheetName)
        {
            Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == sheetName);
            if (sheet == null) throw new Exception(string.Format("Could not find sheet with name {0}", sheetName));
            else return workbookPart.GetPartById(sheet.Id) as WorksheetPart;
        }

        public void GetExcelWorkSheet(string filename, string WorksheetName)
        {
            string fileName = filename;

            if (Path.GetExtension(filename) == ".xls")
            {
                Excel03to07(filename);
                fileName = Path.GetTempPath() + "EMMA2\\PEARR\\temporary.xlsm";
            }
            
            using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(fs, false))
                {
                    WorkbookPart workbookPart = doc.WorkbookPart;
                    SharedStringTablePart sstpart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
                    SharedStringTable sst = sstpart.SharedStringTable;

                    var styleSheet = doc.WorkbookPart.WorkbookStylesPart.Stylesheet;

                    GetFonts(styleSheet);
                    GetFill(styleSheet);
                    GetBorder(styleSheet);
                    GetStyleFormat(styleSheet);

                    WorksheetPart worksheetPart = GetWorksheetFromSheetName(workbookPart, WorksheetName);

                    Worksheet sheet = worksheetPart.Worksheet;

                    var cells = sheet.Descendants<Cell>();
                    var rows = sheet.Descendants<Row>();

                    //Debug.WriteLine("Row count = {0}", rows.LongCount());
                    //Debug.WriteLine("Cell count = {0}", cells.LongCount());
                    /*
                    // One way: go through each cell in the sheet
                    foreach (Cell cell in cells)
                    {
                        if ((cell.DataType != null) && (cell.DataType == CellValues.SharedString))
                        {
                            int ssid = int.Parse(cell.CellValue.Text);
                            string str = sst.ChildElements[ssid].InnerText;
                            Debug.WriteLine(cell.OuterXml);
                        }
                        else if (cell.CellValue != null)
                        {
                            
                            //Debug.WriteLine("Cell contents: {0}", cell.CellValue.Text);
                        }
                    }
                    
                    // Or... via each row
                    */
                    try
                    {
                        MergeCells mergedCells = worksheetPart.Worksheet.Elements<MergeCells>().First();
                        foreach (MergeCell mc in mergedCells)
                        {
                            Merges.Add(mc.Reference.Value);
                        }
                    }
                    catch { }

                    //try
                    {
                        foreach (Row row in rows)
                        {
                            foreach (Cell c in row.Elements<Cell>())
                            {
                                if ((c.DataType != null) && (c.DataType == CellValues.SharedString))
                                {
                                    GetCell(c, sst, true);
                                }
                                else if (c.CellValue != null)
                                {
                                    GetCell(c, sst, false);
                                }
                            }
                        }
                    }
                    //catch { }
                }
            }
            if (File.Exists(Path.GetTempPath() + "EMMA2\\PEARR\\temporary.xlsm"))
                File.Delete(Path.GetTempPath() + "EMMA2\\PEARR\\temporary.xlsm");
        }

        void Check()
        {
            Regex rx = new Regex(@"([A-Z]{1,}[0-9]{1,})|(\$[A-Z]{1,}\$[0-9]{1,})");
            var col = rx.Matches("IF(B48>\" \",IF(LEFT(B48,3)=\"bl.\",MID(B48,4,2)/1000*7850,IF(LEFT(B48,3)=\"Br.\",MID(B48,4,2)/1000*7850,INDEX([1]PROFILE!$I$1:$I$15000,MATCH(B48,[1]PROFILE!$A$1:$A$15000,0),1))),\"\")");
            foreach (var c in col)
            {
                Debug.WriteLine(c);
            }
        }
    }

    public class ExcelCell
    {
        public ExcelCell()
        {
            Coordinates = new List<int[]> { };
        }
        public string Zakres { get; set; }
        public List<int[]> Coordinates { get; set; }
        public uint Styl { get; set; }
        public Arkusz.CellType Typ { get; set; }
        public object Value { get; set; }
        public string Formula { get; set; }
        public override string ToString()
        {
            string corrd = "";
            foreach(var v in Coordinates)
            {
                corrd += "[" + v[0] + ";" + v[1] + "] " ;
            }
            string styl = "";
            string typ = "";
            string value = "";
            string formula = "";

            if (Styl != null) styl = Styl.ToString();
            if (Typ != null) typ = Typ.ToString();
            if (Value != null) value = Value.ToString();
            if (Formula != null) formula = Formula;

            return string.Format("Cell: {0}; Coordinates: {1}, Styl: {2}; Typ: {3}, Value: {4}, Formula: {5}", Zakres, corrd, styl, typ, value, formula);
        }
    }
    public class ExcelFont
    {
        public double size { get; set; }
        public string fontname { get; set; }
        public FontStyle[] Styl { get; set; }
        public string hexColor { get; set; }

        public override string ToString()
        {
            string s = "";
            foreach (var v in Styl)
            {
                s += v + ", ";
            }
            return string.Format("size:{0}; fontname:{1}; FontStyle:{2}; hexColor:{3}", size, fontname, s, hexColor);
        }
    }
    public class ExcelFill
    {
        public string hexColor { get; set; }
        public override string ToString()
        {
            return string.Format("hexColor:{0}", hexColor);
        }
    }
    public class ExcelBorder
    {
        public Thickness BorderThickness { get; set; }
        public BorderStyl Styl { get; set; }
        public string borderhexColor { get; set; }
        public override string ToString()
        {
            return string.Format("BorderThickness:{0}; Styl:{1}; borderhexColor:{2}", BorderThickness, Styl, borderhexColor);
        }
    }
    public class ExcelCellFormat
    {
        public uint fontID { get; set; }
        public uint fillID { get; set; }
        public uint borderID { get; set; }
        public ExcelAlignment Alignment { get; set; }
        public override string ToString()
        {
            if (fontID == null)
                fontID = 1000000;
            if (fillID == null)
                fillID = 1000000;
            if (borderID == null)
                borderID = 1000000;
            if (Alignment == null)
                Alignment = new ExcelAlignment();

            return string.Format("fontID:{0}; fillID:{1}; borderID:{2}; Alignment:[{3}]", fontID, fillID, borderID, Alignment.ToString());
        }
    }
    public class ExcelAlignment
    {
        public HorizontalAligment HAligment { get; set; }
        public VerticalAligment VAligment { get; set; }
        public bool Wrap { get; set; }
        public override string ToString()
        {
            return string.Format("HAligment:{0}; VAligment:{1}; Wrap:{2}", HAligment, VAligment, Wrap);
        }
    }
    public class ExcelFillColection : Collection<ExcelFill> { }
    public class ExcelFontColection : Collection<ExcelFont> { }
    public class ExcelBorderColection : Collection<ExcelBorder> { }
    public class ExcelCellFormatColection : Collection<ExcelCellFormat> { }
    public class ExcelCellColection : Collection<ExcelCell> { }
}
