//Create Excel Files 

PrintExcel report = new PrintExcel();
            report.CreateExcelDoc(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)+@"\Report.xlsx",
            () =>
            {
                
                report.AddFont(15, new FontStyle[] { FontStyle.Bold, FontStyle.Italic }, "Arial");
                report.AddFill();
                report.AddFill("3ABFCD");
                report.AddBorders(new Thickness(1, 1, 1, 1), BorderStyl.Medium);
                Alignment align = report.GetAlignent(VerticalAligment.Center, HorizontalAligment.Center, true);
                report.AddStyle(align, 1, 2, 1);
                
                Arkusz arkusz = new Arkusz();
                arkusz.CellEdit("A4:B4", 1, Arkusz.CellType.Number, 3.6);
                arkusz.CellEdit("C3", 0, Arkusz.CellType.String, "Test");
                arkusz.CellEdit("E2:G1", 1, Arkusz.CellType.String, "Test Test Test Test Test Test Test Test Test");

                arkusz.ColumnWidth("A", 10);
                arkusz.ColumnWidth("B", 10);
                arkusz.ColumnWidth("C", 10);

                arkusz.PageSetup(new double[] { 0.4, 0.4, 0.4, 0.4, 0.2, 0.2 });
                report.AddSheet(arkusz, "NAZWA Arkusza");
                report.PrintPage("NAZWA Arkusza", "A1:D8");
 
                
                Arkusz arkusz2 = new Arkusz();
                arkusz2.CellEdit("D4:E4", 1, Arkusz.CellType.Number, 3.6);
                arkusz2.CellEdit("K3", 0, Arkusz.CellType.String, "Test");
                arkusz2.CellEdit("A2:B1", 1, Arkusz.CellType.String, "Test Test Test Test Test Test Test Test Test");

                arkusz2.ColumnWidth("A", 10);
                arkusz2.ColumnWidth("B", 10);
                arkusz2.ColumnWidth("C", 10);

                report.AddSheet(arkusz2, "NAZWA Arkusza2");
                arkusz2.PageSetup(new double[] { 0.8, 0.8, 0.8, 0.8, 0.5, 0.5 });
                report.PrintPage("NAZWA Arkusza2", "A1:G8");

                report.Save();
            }
            );
            report.Dispose();
            report = null;

            Console.WriteLine("Excel file has created!");
        }
