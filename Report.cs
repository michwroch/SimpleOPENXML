using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace OpenXML_Driver
{
    [StructLayout(LayoutKind.Explicit)]
    public struct Thickness
    {
        [FieldOffset(0)]
        public double l;
        [FieldOffset(0)]
        public double r;
        [FieldOffset(0)]
        public double t;
        [FieldOffset(0)]
        public double b;

        public Thickness(double left)
        {
            l = left;
            r = 0;
            t = 0;
            b = 0;
        }

        public Thickness(double left, double right)
        {
            l = left;
            r = right;
            t = 0;
            b = 0;
        }

        public Thickness(double left, double right, double top)
        {
            l = left;
            r = right;
            t = top;
            b = 0;
        }

        public Thickness(double left, double right, double top, double buttom)
        {
            l = left;
            r = right;
            t = top;
            b = buttom;
        }
    }

    public class PrintExcel
    {
        WorkbookPart workbookPart;
        Sheets sheets;
        //
        //Enumerator Styli czcionki
        public enum FontStyle
        {
            Normal,
            Bold,
            Italic,
            Underline,
            Strike
        }

        //Enumerator ramki
        public enum BorderStyl
        {
            None,
            Thin,
            Medium,
            Dashed,
            Dotted,
            Thick,
            Double,
            Hair,
            MediumDashed,
            DashDot,
            MediumDashDot,
            DashDotDot,
            MediumDashDotDot,
            SlantDashDot
        }

        //Enumearator Vertical
        public enum VerticalAligment
        {
            None,
            Top,
            Bottom,
            Center,
            Justify,
            Distributed
        }

        //Enumearator Horizontal
        public enum HorizontalAligment
        {
            None,
            Left,
            Right,
            Center,
            Fill,
            General,
            Justify,
            Distributed,
            CenterContinuous
        }

        //Stałe
        Fonts fonts = new Fonts();
        Fills fills = new Fills();
        Borders borders = new Borders();
        CellFormats cellFormats = new CellFormats();
        uint nr_sheet = 1;

        //Inicjacja klasy
        public PrintExcel()
        {
            AddFont();
            AddBorders();
            AddFill();
            AddStyle();
        }

        //Zwolniij Ram
        public void Dispose()
        {
            workbookPart = null;
            sheets = null;
            fonts = null;
            fills = null;
            borders = null;
            cellFormats = null;
        }

        //Utwórz plik Excela
        public void CreateExcelDoc(string fileName, Action Body)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook))
            {
                workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                WorkbookStylesPart stylePart = workbookPart.AddNewPart<WorkbookStylesPart>();
                stylePart.Stylesheet = new Stylesheet(fonts, fills, borders, cellFormats);
                stylePart.Stylesheet.Save();

                sheets = workbookPart.Workbook.AppendChild(new Sheets());

                Body();

                workbookPart.Workbook.Save();

                stylePart = null;
            }
        }

        //Dodaj czcionkę
        public void AddFont()
        {
            Font font = new Font();
            fonts.Append(font);
        }

        public void AddFont(double size, FontStyle[] Styl, string fontname = "Calibri", string hexColor = "000000")
        {
            Font font = new Font(
                    new FontSize() { Val = size },
                    new FontName() { Val = fontname },
                    new Color() { Rgb = new HexBinaryValue() { Value = hexColor } }
                );

            foreach (var styl in Styl)
            {
                if (styl == FontStyle.Bold)
                    font.Bold = new Bold();
                if (styl == FontStyle.Italic)
                    font.Italic = new Italic();
                if (styl == FontStyle.Underline)
                    font.Underline = new Underline();
                if (styl == FontStyle.Strike)
                    font.Strike = new Strike();
            }

            fonts.Append(font);
        }

        //Dodaj wypełnienie
        public void AddFill()
        {
            fills.Append(
            new Fill(new PatternFill() { PatternType = PatternValues.Solid }));
        }

        public void AddFill(string hexColor = "000000")
        {
            // BackgroundColor = new BackgroundColor() { Rgb = new HexBinaryValue() { Value = hexColor } },
            fills.Append(
            new Fill(new PatternFill() { PatternType = PatternValues.Solid, ForegroundColor=new ForegroundColor() { Rgb = new HexBinaryValue() { Value = hexColor } } }));
        }

        //Dodaj Ramki
        public void AddBorders()
        {
            borders.Append(new Border());
        }

        public void AddBorders(Thickness BorderThickness, BorderStyl style = BorderStyl.Thin, string borderhexColor = "000000")
        {
            Border border = new Border();

            if (BorderThickness.l > 0)
            {
                border.LeftBorder = new LeftBorder(new Color() { Rgb = borderhexColor }) { Style = (BorderStyleValues)style };
            }
            if (BorderThickness.r > 0)
            {
                border.RightBorder = new RightBorder(new Color() { Rgb = borderhexColor }) { Style = (BorderStyleValues)style };
            }
            if (BorderThickness.t > 0)
            {
                border.TopBorder = new TopBorder(new Color() { Rgb = borderhexColor }) { Style = (BorderStyleValues)style };
            }
            if (BorderThickness.b > 0)
            {
                border.BottomBorder = new BottomBorder(new Color() { Rgb = borderhexColor }) { Style = (BorderStyleValues)style };
            }
            borders.Append(border);
        }

        //Dodaj wyśrodkowanie
        public Alignment GetAlignent(VerticalAligment Vertical = VerticalAligment.None, HorizontalAligment Horizontal = HorizontalAligment.None, bool Wrap = false)
        {
            Alignment a1 = new Alignment();


            if (Vertical != VerticalAligment.None)
            {
                VerticalAlignmentValues agl = VerticalAlignmentValues.Bottom;
                if (Vertical == VerticalAligment.Bottom) agl = VerticalAlignmentValues.Bottom;
                if (Vertical == VerticalAligment.Center) agl = VerticalAlignmentValues.Center;
                if (Vertical == VerticalAligment.Justify) agl = VerticalAlignmentValues.Justify;
                if (Vertical == VerticalAligment.Distributed) agl = VerticalAlignmentValues.Distributed;
                if (Vertical == VerticalAligment.Top) agl = VerticalAlignmentValues.Top;

                a1.Vertical = new EnumValue<VerticalAlignmentValues>(agl);
            }
            if (Horizontal != HorizontalAligment.None)
            {
                HorizontalAlignmentValues agl = HorizontalAlignmentValues.Left;
                if (Horizontal == HorizontalAligment.Left) agl = HorizontalAlignmentValues.Left;
                if (Horizontal == HorizontalAligment.Right) agl = HorizontalAlignmentValues.Right;
                if (Horizontal == HorizontalAligment.Center) agl = HorizontalAlignmentValues.Center;
                if (Horizontal == HorizontalAligment.Fill) agl = HorizontalAlignmentValues.Fill;
                if (Horizontal == HorizontalAligment.General) agl = HorizontalAlignmentValues.General;
                if (Horizontal == HorizontalAligment.Justify) agl = HorizontalAlignmentValues.Justify;
                if (Horizontal == HorizontalAligment.Distributed) agl = HorizontalAlignmentValues.Distributed;
                if (Horizontal == HorizontalAligment.CenterContinuous) agl = HorizontalAlignmentValues.CenterContinuous;

                a1.Horizontal = new EnumValue<HorizontalAlignmentValues>(agl);
            }
            a1.WrapText = Wrap;

            return a1;
        }

        //Dodaj Styl
        public void AddStyle(uint fontID = 0, uint fillID = 0, uint borderID = 0)
        {
            CellFormat CellFormatt = new CellFormat();
            if (fonts.ChildElements.Count > fontID)
                CellFormatt.FontId = new UInt32Value(fontID);
            if (fills.ChildElements.Count > fillID)
                CellFormatt.FillId = new UInt32Value(fillID);
            if (borders.ChildElements.Count > borderID)
                CellFormatt.BorderId = new UInt32Value(borderID);

            cellFormats.Append(CellFormatt);
        }

        public void AddStyle(Alignment alignment, uint fontID = 0, uint fillID = 0, uint borderID = 0)
        {
            CellFormat CellFormatt = new CellFormat();

            CellFormatt.Alignment = alignment;

            if (fonts.ChildElements.Count > fontID)
                CellFormatt.FontId = new UInt32Value(fontID);
            if (fills.ChildElements.Count > fillID)
                CellFormatt.FillId = new UInt32Value(fillID);
            if (borders.ChildElements.Count > borderID)
                CellFormatt.BorderId = new UInt32Value(borderID);

            cellFormats.Append(CellFormatt);
        }

        //Ustaw pole zadruku

        List<string> PagePr = new List<string>();
        public void PrintPage(string NazwaArkusz, string Zakres)
        {
            string[] zakres = Zakres.Split(':');

            List<string[]> Explode = new List<string[]> { };
            foreach (string s in zakres)
            {
                Regex Rgx = new Regex(@"([A-Z]{1,})");
                string[] Range = Rgx.Split(s);
                List<string> range = Range.ToList();
                range.RemoveAll(x => x == string.Empty);
                Range = range.ToArray();
                Explode.Add(Range);
            }

            PagePr.Add(string.Format("'{0}'!${1}${2}:${3}${4}", NazwaArkusz, Explode[0][0], Explode[0][1], Explode[1][0], Explode[1][1]));
        }


        public void Save()
        {
            DefinedNames definedNames = new DefinedNames();
            uint u = 0;
            foreach (var s in PagePr)
            {
                DefinedName printAreaDefName = new DefinedName() { Name = "_xlnm.Print_Area", LocalSheetId = (UInt32Value)u };
                printAreaDefName.Text = s;
                definedNames.Append(printAreaDefName);
                u++;
            }
            workbookPart.Workbook.Append(definedNames);
        }
        //Dodaj Arkusz
        public void AddSheet(Arkusz arkusz, string Nazwa)
        {
            arkusz.SaveSheet(workbookPart, sheets, ref nr_sheet, Nazwa);
        }
    }

    public class Arkusz
    {
        //Stałe
        List<string> RangeMerge = new List<string> { };
        public string Nazwa { get; set; }
        List<Kolumna> Kolumny = new List<Kolumna> { };
        public List<Row> Wiersze = new List<Row> { };

        //Zapisz Arkusz
        public void SaveSheet(WorkbookPart workbookPart, Sheets sheets, ref uint id, string nazwa)
        {
            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet();
               
            worksheetPart.Worksheet.AppendChild(GetColumns());

            Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = id, Name = nazwa };
            sheets.Append(sheet);

            SheetData sheetData = worksheetPart.Worksheet.AppendChild(new SheetData());

            foreach (var row in Wiersze)
            {
                sheetData.AppendChild(row);
            }

            worksheetPart.Worksheet.InsertAfter(GetMerge(), worksheetPart.Worksheet.Elements<SheetData>().First());


            PageMargins pageMargins1;
            PageSetup pageSetup;

            if(margins.Length == 0)
            {
                margins = new double[] { 0.4, 0.4, 0.4, 0.4, 0.2, 0.2 };
            }

            pageMargins1 = new PageMargins();
            pageMargins1.Left = margins[0];
            pageMargins1.Right = margins[1];
            pageMargins1.Top = margins[2];
            pageMargins1.Bottom = margins[3];
            pageMargins1.Header = margins[4];
            pageMargins1.Footer = margins[5];

            pageSetup = new PageSetup();
            pageSetup.Orientation = OrientationValues.Landscape;
            pageSetup.FitToHeight = 2;
            pageSetup.HorizontalDpi = 200;
            pageSetup.VerticalDpi = 200;

            worksheetPart.Worksheet.AppendChild(pageMargins1);
            //worksheetPart.Worksheet.AppendChild(pageSetup);


            worksheetPart.Worksheet.Save();
            id++;
            Nazwa = nazwa;
        }

        //Pobierz scalenia
        MergeCells GetMerge()
        {
            MergeCells mergeCells = new MergeCells();

            foreach (string s in RangeMerge)
            {
                mergeCells.Append(new MergeCell() { Reference = new StringValue(s) });
            }

            return mergeCells;
        }

        double[] margins = new double[] { };
        public void PageSetup(double[] margins_)
        {
            margins = margins_;
        }


        private List<int[]> getrange(ref string Zakres)
        {
            List<int[]> Range = new List<int[]> { };

            if (Zakres.Contains(':'))
            {
                string[] Ranges = Zakres.Split(':');

                Regex Rgx1 = new Regex(@"([A-Z]{1,})");
                string[] Range1 = Rgx1.Split(Ranges[0]);
                List<string> range1 = Range1.ToList();
                range1.RemoveAll(x => x == string.Empty);
                Range1 = range1.ToArray();
                Regex Rgx2 = new Regex(@"([A-Z]{1,})");
                string[] Range2 = Rgx2.Split(Ranges[1]);
                List<string> range2 = Range2.ToList();
                range2.RemoveAll(x => x == string.Empty);
                Range2 = range2.ToArray();

                if (int.Parse(Range1[1]) > int.Parse(Range2[1]))
                {
                    Zakres = Range1[0] + Range2[1] + ":" + Range2[0] + Range1[1];
                }

                if (!RangeMerge.Contains(Zakres))
                    RangeMerge.Add(Zakres);

                Ranges = Zakres.Split(':');

                Range.Add(ReportHelper.getxy(Ranges[0]));
                Range.Add(ReportHelper.getxy(Ranges[1]));
            }
            else
            {
                Range.Add(ReportHelper.getxy(Zakres));
            }

            return Range;
        }

        //Dodaj kolumy 
        void AddCols()
        {
            Kolumny.Add(new Kolumna());
        }

        //Ustaw szerokość kolumny
        public void ColumnWidth(string NazwaKolumny, double width)
        {
            int[] ob_ = ReportHelper.getxy(NazwaKolumny + "1");

            if (Kolumny.Count < ob_[1])
            {
                if (Kolumny.Count == 0)
                {
                    for (int i = 0; i < ob_[1]; i++)
                        AddCols();
                }
                else
                {
                    foreach (Row row_ in Wiersze)
                    {
                        for (int i = Kolumny.Count - 1; i < ob_[1]; i++)
                        {
                            row_.Append(new Cell());
                        }
                    }

                    for (int i = Kolumny.Count - 1; i < ob_[1]; i++)
                        AddCols();
                }
            }

            for (int i = 0; i < Kolumny.Count; i++)
                Kolumny[i].ID = new UInt32Value((uint)(i + 1));

            Kolumny[ob_[1] - 1].Width = new DoubleValue(width);
        }

        //Wysokość wiersza
        public void RowHeigth(int Number, double height)
        {
            if (Number - 1 > Wiersze.Count)
            {
                Row row = new Row();
                for (int i = 0; i < Kolumny.Count; i++)
                    row.Append(new Cell());

                Wiersze.Add(row);
            }

            Wiersze[Number - 1].Height = new DoubleValue(height);
            Wiersze[Number - 1].CustomHeight = true;
        }

        //Pobierz komórkę
        public Cell[] GetCell(string Zakres)
        {
            List<int[]> xy = getrange(ref Zakres);

            foreach (object ob in xy)
            {
                int[] ob_ = ob as int[];

                if (Kolumny.Count < ob_[1])
                {
                    if (Kolumny.Count == 0)
                    {
                        for (int i = 0; i < ob_[1]; i++)
                            AddCols();
                    }
                    else
                    {
                        foreach (Row row_ in Wiersze)
                        {
                            for (int i = Kolumny.Count - 1; i < ob_[1]; i++)
                            {
                                row_.Append(new Cell());
                            }
                        }

                        for (int i = Kolumny.Count - 1; i < ob_[1]; i++)
                            AddCols();
                    }
                }

                if (Wiersze.Count < ob_[0])
                {
                    for (int i = Wiersze.Count - 1; i < ob_[0]; i++)
                    {
                        Row row_ = new Row();
                        for (int j = 0; j < Kolumny.Count; j++)
                            row_.Append(new Cell());
                        Wiersze.Add(row_);
                    }
                }
            }
            for (int i = 0; i < Kolumny.Count; i++)
                Kolumny[i].ID = new UInt32Value((uint)(i + 1));

            int x = xy[0][0] - 1;
            int y = xy[0][1] - 1;
            Row row = Wiersze[x];
            List<Cell> outp = new List<Cell> { };

            if (row == null)
                return null;

            List<Cell> Cells1 = row.Elements<Cell>().ToList();

            if (xy.Count > 1)
            {
                int xend = xy[1][0] - 1;
                int yend = xy[1][1] - 1;
                Row rowend = Wiersze[xend];


                List<Cell> Cells2 = rowend.Elements<Cell>().ToList();

                for (int i = y; i <= yend; i++)
                {
                    outp.Add(Cells1[i]);
                    outp.Add(Cells2[i]);
                }

            }
            else
            {
                outp.Add(Cells1[y]);
            }

            return outp.ToArray();
        }

        //Pobierz kolumny
        public Columns GetColumns()
        {
            Columns col = new Columns();

            foreach (Kolumna kol in Kolumny)
            {
                BooleanValue widthflag = false;
                if (kol.Width != null)
                {
                    col.Append(
                            new Column
                            {
                                Min = kol.ID,
                                Max = kol.ID,
                                Width = kol.Width,
                                CustomWidth = true
                            }
                        );
                }
                else
                {
                    col.Append(
                                new Column
                                {
                                    Min = kol.ID,
                                    Max = kol.ID,
                                    Width = new DoubleValue(9.2d),
                                    CustomWidth = true
                                }
                            );
                }
            }

            return col;
        }

        //Ustaw komórkę
        public void CellEdit(string Zakres, uint styl = 0, CellType typ = CellType.String, object text = null, string formula = "")
        {
            CellValues cv = CellValues.String;
            if (typ == CellType.Date) cv = CellValues.Date;
            if (typ == CellType.Number) cv = CellValues.Number;
            if (typ == CellType.String) cv = CellValues.String;

            Cell[] cells = GetCell(Zakres);

            foreach (Cell cell in cells)
            {
                cell.StyleIndex = new UInt32Value(styl);
                cell.DataType = new EnumValue<CellValues>(cv);
            }

            if (text != null)
            {
                if (typ == CellType.Number)
                    cells.First().CellFormula = new CellFormula(text.ToString().Replace(",", "."));
                else
                    cells.First().CellValue = new CellValue(text.ToString());
            }

            if (formula != "")
            {
                cells.First().CellFormula = new CellFormula(formula);
            }
        }

        public string GetRange(params int[] arg)
        {
            string o = "";

            if(arg.Length == 2)
            {
                o += GetExcelColumnName(arg[0]) + arg[1];
            }
            if (arg.Length == 4)
            {
                o += GetExcelColumnName(arg[0]) + arg[1] + ":" + GetExcelColumnName(arg[2]) + arg[3];
            }
            return o;
        }

        public string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }

        public enum CellType
        {
            Date,
            Number,
            String
        }
    }

    public class Kolumna
    {
        public UInt32Value ID { get; set; }
        public DoubleValue Width { get; set; }
    }

    public static class ReportHelper
    {
        static char[] colsrangename = new char[] {'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L',
        'M', 'N', 'O', 'P','Q', 'R', 'S','T', 'U','V','W','X','Y','Z'};

        //pobierz  indeksy komórek
        public static int[] getxy(string Zakres)
        {
            int[] ret = new int[2];
            int retx = 0;
            int rety = 0;

            Regex Rgx = new Regex(@"([A-Z]{1,})");
            string[] Range = Rgx.Split(Zakres);
            List<string> range = Range.ToList();
            range.RemoveAll(x => x == string.Empty);
            Range = range.ToArray();

            retx = int.Parse(Range[1]);

            for (int i = Range[0].Length - 1; i > -1; i--)
            {
                if (i == Range[0].Length - 1)
                {
                    int letter = colsrangename.ToList().FindIndex(x => x == Range[0][i]);
                    rety += letter + 1;
                }
                else
                {
                    int row = Range[0].Length - i - 1;
                    int letter = colsrangename.ToList().FindIndex(x => x == Range[0][i]) + 1;
                    rety += (int)(Math.Pow(26, (double)row) * (double)letter);
                }
            }
            ret[0] = retx;
            ret[1] = rety;

            return ret;
        }
    }
}
