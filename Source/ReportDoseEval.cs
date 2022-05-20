// This program tests MigraDoc
using System;
using System.Collections.Generic;

using PdfSharp;
using PdfSharp.Pdf;

using MigraDoc;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Shapes;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.Rendering;

namespace ReportPrinter
{
  class DoseEval
  {
    static void Main(string[] args)
    {
      // Fill in blank test data
      VMS.TPS.Script.DoseEvalReportData ReportData =
        new VMS.TPS.Script.DoseEvalReportData();
      ReportData.LogoFileName = "test_logo_lunar_eclipse.jpg";
      ReportData.DepartmentName = "Department Name";
      ReportData.PatientName = "Last, First";
      ReportData.PatientID = "E123456";
      ReportData.SummaryText =
        "For Last, First, the following treatment planning constraints for "+
        "target coverage and critical organ coverage were requested. After "+
        "completing the planning process the plan constraints were assessed "+
        "according to the list. After reviewing the plan constraints with the "+
        "dosimetrist and physicist, the constraints have been approved.";

      VMS.TPS.Script.PlanComparison tempResult1 =
        new VMS.TPS.Script.PlanComparison();
      tempResult1.OrganName = "PTV";
      tempResult1.StructureName = "PTV";
      tempResult1.ConstraintName = "D95% (>)";
      tempResult1.ConstraintValue = "70 Gy";
      tempResult1.PlanValue = "70.3 Gy";
      tempResult1.PassFail = "Pass";
      tempResult1.ConstraintNote = "Physician target coverage";
      ReportData.ComparisonResults.Add(tempResult1);

      VMS.TPS.Script.PlanComparison tempResult2 =
        new VMS.TPS.Script.PlanComparison();
      tempResult2.OrganName = "Spinal Cord";
      tempResult2.StructureName = "Spinal Cord";
      tempResult2.ConstraintName = "Max Dose (<)";
      tempResult2.ConstraintValue = "45 Gy";
      tempResult2.PlanValue = "38.5 Gy";
      tempResult2.PassFail = "Pass";
      tempResult2.ConstraintNote = "Normal tissue constraint";
      ReportData.ComparisonResults.Add(tempResult2);

      // Create a MigraDoc document
      string status = CreateDoseEvalReport(ReportData, "TestReport.pdf");

      // Exit
      Console.WriteLine(status);
    }

    public static string CreateDoseEvalReport(
      VMS.TPS.Script.DoseEvalReportData Data, string filename)
    {
      // Create a new MigraDoc document
      Document document = new Document();
      document.UseCmykColor = true;
      const bool unicode = false;

      document.Info.Title = "Test Report";
      document.Info.Subject =
        "Demonstrates an excerpt of the capabilities of MigraDoc.";
      document.Info.Author = "Steven Dolly";

      ////////////
      // Styles //
      ////////////

      // Get the predefined style Normal.
      Style style = document.Styles["Normal"];
      // Because all styles are derived from Normal, the next line changes the
      // font of the whole document. Or, more exactly, it changes the font of
      // all styles and paragraphs that do not redefine the font.
      style.Font.Name = "Times New Roman";
      style.Font.Size = 12;
      style.ParagraphFormat.Alignment = ParagraphAlignment.Justify;
      // Heading1 - Heading9 are predefined styles
      // Heading1
      style = document.Styles["Heading1"];
      style.Font.Name = "Tahoma";
      style.Font.Size = 16;
      style.Font.Bold = true;
      style.Font.Color = Colors.DarkBlue;
      style.ParagraphFormat.Alignment = ParagraphAlignment.Center;
      style.ParagraphFormat.PageBreakBefore = false;
      style.ParagraphFormat.SpaceAfter = 10;
      // Heading2
      style = document.Styles["Heading2"];
      style.Font.Name = "Tahoma";
      style.Font.Size = 14;
      style.Font.Bold = true;
      style.Font.Color = Colors.Black;
      style.ParagraphFormat.Alignment = ParagraphAlignment.Center;
      style.ParagraphFormat.PageBreakBefore = false;
      style.ParagraphFormat.SpaceAfter = 10;
      // Heading3
      style = document.Styles["Heading3"];
      style.Font.Name = "Times New Roman";
      style.Font.Size = 12;
      style.Font.Bold = true;
      style.Font.Color = Colors.Black;
      style.ParagraphFormat.Alignment = ParagraphAlignment.Left;
      style.ParagraphFormat.PageBreakBefore = false;
      style.ParagraphFormat.SpaceAfter = 10;

      //////////////////
      // Main Content //
      //////////////////
      Section section = document.AddSection();
      section.PageSetup = document.DefaultPageSetup.Clone();
      section.PageSetup.PageWidth = Unit.FromPoint(612);
      section.PageSetup.PageHeight = Unit.FromPoint(792);
      section.PageSetup.TopMargin = "2.5cm";
      section.PageSetup.BottomMargin = "2.5cm";
      section.PageSetup.LeftMargin = "2.5cm";
      section.PageSetup.RightMargin = "2.5cm";

      var image = section.AddImage(Data.LogoFileName);
      image.LockAspectRatio = true;
      image.Width = "5cm";
      image.Left = ShapePosition.Center;
      document.LastSection.AddParagraph();

      document.LastSection.AddParagraph(Data.DepartmentName, "Heading1");
      document.LastSection.AddParagraph("Dose Constraint Report", "Heading2");

      var NameString = Data.PatientName + " (" + Data.PatientID + ")";
      Paragraph paragraph = document.LastSection.AddParagraph(
        "Patient Name (ID): ", "Heading3");
      paragraph.AddText(NameString);

      paragraph = document.LastSection.AddParagraph();
      paragraph.AddText(Data.SummaryText);
      paragraph = document.LastSection.AddParagraph();

      ///////////
      // Table //
      ///////////
      Table table = new Table();
      table.Borders.Width = 0.5;
      // Columns
      Column column = table.AddColumn(Unit.FromCentimeter(4));
      column.Format.Alignment = ParagraphAlignment.Center;
      column = table.AddColumn(Unit.FromCentimeter(3));
      column.Format.Alignment = ParagraphAlignment.Center;
      column = table.AddColumn(Unit.FromCentimeter(2));
      column.Format.Alignment = ParagraphAlignment.Center;
      column = table.AddColumn(Unit.FromCentimeter(2.4));
      column.Format.Alignment = ParagraphAlignment.Center;
      column = table.AddColumn(Unit.FromCentimeter(5));
      column.Format.Alignment = ParagraphAlignment.Center;
      column.Shading.Color = Colors.PaleGoldenrod;
      // Header Row 1
      Row row = table.AddRow();
      //row.Shading.Color = Colors.PaleGoldenrod;
      Cell cell = row.Cells[0];
      cell.AddParagraph("Structure");
      cell = row.Cells[1];
      cell.AddParagraph("Constraint");
      row.Cells[1].MergeRight = 1;
      cell = row.Cells[3];
      cell.AddParagraph("Plan Value");
      cell = row.Cells[4];
      cell.AddParagraph("Notes");
      // Data Rows
      foreach(VMS.TPS.Script.PlanComparison c in Data.ComparisonResults)
      {
        row = table.AddRow();
        cell = row.Cells[0];
        cell.AddParagraph(c.OrganName);
        cell = row.Cells[1];
        cell.AddParagraph(c.ConstraintName);
        cell = row.Cells[2];
        cell.AddParagraph(c.ConstraintValue);
        cell = row.Cells[3];
        cell.AddParagraph(c.PlanValue);
        cell = row.Cells[4];
        cell.AddParagraph(c.ConstraintNote);
      }
      // Finalize table and add to section
      table.SetEdge(0, 0, 5, 3, Edge.Box, BorderStyle.Single, 1.0, Colors.Black);
      document.LastSection.Add(table);

      // Render as PDF and save
      PdfDocumentRenderer renderer = new PdfDocumentRenderer(unicode);
      renderer.Document = document;
      renderer.RenderDocument();
      renderer.PdfDocument.Save(filename);

      string statusText = "PDF Report Created!";
      return statusText;
    }
  }
}
