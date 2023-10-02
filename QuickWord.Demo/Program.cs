using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using QuickWord.OpenXml;
using QuickWord.OpenXml.DrawingExtensions;
using QuickWord.OpenXml.Extras;
using QuickWord.OpenXml.Extras.Extensions;
using QuickWord.OpenXml.QuickObjects;
using System.Diagnostics;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

// // // // // // // // // // // // // // // //
// Create document
// // // // // // // // // // // // // // // //

string fileName = "TEST.docx";
using var document = WordprocessingDocument.Create(fileName, WordprocessingDocumentType.Document);
Body body = document.CreateBody()
	.LeftMargin(1.5, MeasuringUnits.Centimeters)
	.RightMargin(1.5, MeasuringUnits.Centimeters)
	.PageWidth(8.2, MeasuringUnits.Inches)
	.PageHeight(11.7, MeasuringUnits.Inches);

// // // // // // // // // // // // // // // //
// Paragraph with 6 different runs
// // // // // // // // // // // // // // // //

Run[] runs = new[]
{
	new Run().Text("This is a")
		.FontSize(16)
		.FontColor("Red")
		.Bold(),

	new Run().Text(" single paragraph ")
		.FontSize(24)
		.FontColor("Green")
		.Italic()
		.Outline()
		.FontFace("Comic Sans MS"),

	new Run().Text("with 5 different")
		.FontSize(18)
		.FontColor("Blue")
		.HighlightColor(HighlightColorValues.Cyan),

	new Run().Text(" Runs in it.")
		.FontSize(14)
		.FontColor("Magenta")
		.SmallCaps()
		.FontFace("Times New Roman"),

	new Run().Text("\nThe 5th Run is on another line."),

	new Run().Text("\nThere is also a hidden 6th run.")
		.Hidden()
};

body.AppendChild(new Paragraph(runs)
	.Justification(JustificationValues.Center)
	.SpacingBefore(2, LineMeasuringUnits.WholeLines));

// // // // // // // // // // // // // // // //
// Simple hard-coded table example
// // // // // // // // // // // // // // // //

StartSection("Simple table with hard-coded data");

body.AppendChild(QTable.Create(new[]
{
	new[] { "1", "John", "Doe" },
	new[] { "2", "Jane", "Doe" },
	new[] { "3", "John", "Smith" },
	new[] { "4", "Jane", "Smith" }
})
.Width(100, WidthUnits.Percentage)
.Borders(1));

// // // // // // // // // // // // // // // //
// Simple hard-coded table with cell merging
// // // // // // // // // // // // // // // //

StartSection("Simple table with merged cells");

Table table = QTable.Create(new[]
{
	new[] { "One", "Two", "Three", "Four" },
	new[] { "One", "Two", "Three", "Four" },
	new[] { "One", "Two", "Three", "Four" },
	new[] { "One", "Two", "Three", "Four" }
})
.Width(100, WidthUnits.Percentage)
.Borders(1);

table.Rows(0)?.Cells(0)?.VerticalMerge(MergedCellValues.Restart)
	.VerticalContentAlignment(TableVerticalAlignmentValues.Center)
	.Paragraphs(0)?.SpacingAfter(0);
table.Rows(1)?.Cells(0)?.VerticalMerge(MergedCellValues.Continue);

table.Rows(0)?.Cells(1)?.HorizontalMerge(MergedCellValues.Restart)
	.Paragraphs(0)?.Justification(JustificationValues.Center);
table.Rows(0)?.Cells(2)?.HorizontalMerge(MergedCellValues.Continue);
table.Rows(0)?.Cells(3)?.HorizontalMerge(MergedCellValues.Continue);

table.Rows(2)?.Cells(2)?.VerticalMerge(MergedCellValues.Restart)
	.VerticalContentAlignment(TableVerticalAlignmentValues.Center)
	.Paragraphs(0)?
		.Justification(JustificationValues.Center)
		.SpacingAfter(0);
table.Rows(3)?.Cells(2)?.VerticalMerge(MergedCellValues.Continue);
table.Rows(2)?.Cells(3)?.VerticalMerge(MergedCellValues.Restart);
table.Rows(3)?.Cells(3)?.VerticalMerge(MergedCellValues.Continue);
table.Rows(2)?.Cells(2)?.HorizontalMerge(MergedCellValues.Restart);
table.Rows(2)?.Cells(3)?.HorizontalMerge(MergedCellValues.Continue);
table.Rows(3)?.Cells(2)?.HorizontalMerge(MergedCellValues.Restart);
table.Rows(3)?.Cells(3)?.HorizontalMerge(MergedCellValues.Continue);

body.AppendChild(table);

// // // // // // // // // // // // // // // //
// Table with array data
// // // // // // // // // // // // // // // //

Person[] exampleData = new[]
{
	new Person(1, "John", "Doe", true),
	new Person(2, "Jane", "Doe", false),
	new Person(3, "John", "Smith", true),
	new Person(4, "Jane", "Smith", true)
};

StartSection("Table with passed array data");

Table table1 = new Table()
	.Width(100, WidthUnits.Percentage)
	.Borders(1);

table1.AppendChild(QTableRow.Create(
	new[] { "ID", "First Name", "Last Name" },
	rowFormatting: new TableRowFormatting { IsHeader = true },
	runFormatting: new RunFormatting { Bold = true }));

foreach (Person person in exampleData)
{
	table1.AppendChild(QTableRow.Create(
		new[] { person.Id.ToString(), person.FirstName, person.LastName }));
}

body.AppendChild(table1);

// // // // // // // // // // // // // // // //
// Table with formatting
// // // // // // // // // // // // // // // //

StartSection("Table with passed array data and custom formatting");
body.AppendChild(QParagraph.Create("(The table header repeats on each new page)",
	new ParagraphFormatting { Justification = JustificationValues.Center },
	new RunFormatting { FontSize = 9 }));

Table table2 = new Table()
	.BottomBorder(1)
	.Width(100, WidthUnits.Percentage);

ParagraphFormatting cellParagraphFormatting = new Paragraph()
	.Justification(JustificationValues.Center)
	.SpacingBefore(6, LineMeasuringUnits.Points)
	.SpacingAfter(6, LineMeasuringUnits.Points)
	.CloneFormatting();

TableCellFormatting headerCellFormatting = new TableCell()
	.VerticalContentAlignment(TableVerticalAlignmentValues.Center)
	.BottomBorder(1)
	.CloneFormatting();

var headerRunFormatting = new RunFormatting { Bold = true };

TableRow headerRow = QTableRow.Create(
	new[] { "ID", "First Name", "Last Name", "Verified?" },
	new TableRowFormatting { IsHeader = true },
	headerCellFormatting, cellParagraphFormatting, headerRunFormatting);

headerRow.Cells(0)?.Width(30, WidthUnits.Points);
table2.AppendChild(headerRow);

for (int i = 0; i < exampleData.Length; i++)
{
	Person person = exampleData[i];

	TableCellFormatting cellFormatting = new TableCell()
		.VerticalContentAlignment(TableVerticalAlignmentValues.Center)
		.FillColor(i % 2 == 0 ? "#E0E0E0" : "White")
		.CloneFormatting();

	TableRow recordRow = QTableRow.Create(
		new[] { person.Id.ToString(), person.FirstName, person.LastName, person.Verified ? "Yes" : "No" },
		null, cellFormatting, cellParagraphFormatting);

	recordRow.Cells(0)?.Paragraphs(0)?.Runs(0)?.FontColor("Red").Bold();
	recordRow.Cells(1)?.Paragraphs(0)?.Runs(0)?.FontColor("Green");
	recordRow.Cells(2)?.Paragraphs(0)?.Runs(0)?.FontColor("Blue");

	table2.AppendChild(recordRow);
}

foreach (TableCell? cell in table2.GetColumnOfCells(3).Skip(1))
{
	if (cell?.Paragraphs(0)?.GetText() == "Yes")
		cell?.FillColor("LightGreen");
	else
		cell?.FillColor("LightPink");
}

body.AppendChild(table2);

// // // // // // // // // // // // // // // //
// Inlined image
// // // // // // // // // // // // // // // //

StartSection("Inlined image");
body.AppendChild(new Paragraph(
	new Run(QDrawing.FromImage(body, "Icon.png", ImagePartType.Png, 32, 32)),
	new Run().Text(" This image is inlined with the text.").VerticalPosition(8)
).Justification(JustificationValues.Center));

// // // // // // // // // // // // // // // //
// Anchored images
// // // // // // // // // // // // // // // //

StartSection("Anchored images");
body.AppendChild(new EmptyLine().SpacingBefore(0, LineMeasuringUnits.Points).SpacingAfter(0, LineMeasuringUnits.Points));

body.AppendChild(new Paragraph(
	new Run(QDrawing.FromImage(body, "Icon.png", ImagePartType.Png, 128, 128)
		.ToAnchoredDrawing()
		.AbsoluteVerticalPosition(0, ImageMeasuringUnits.Pixels, DW.VerticalRelativePositionValues.Paragraph)
		.SquareWrapping(0, 0, 0.5, 0, ImageMeasuringUnits.Centimeters)),

	new Run().Text("This\nimage\nis\naligned\nbefore\nthe text.\n")
));

body.AppendChild(new Paragraph(
	new Run(QDrawing.FromImage(body, "Icon.png", ImagePartType.Png, 128, 128)
		.ToAnchoredDrawing()
		.AbsoluteVerticalPosition(0, ImageMeasuringUnits.Pixels, DW.VerticalRelativePositionValues.Paragraph)
		.AbsoluteHorizontalPosition(-128, ImageMeasuringUnits.Pixels, DW.HorizontalRelativePositionValues.RightMargin)
		.SquareWrapping(0.5, 0, 0, 0, ImageMeasuringUnits.Centimeters)),

	new Run().Text("This\nimage\nis\naligned\nafter\nthe text.\n")
).Justification(JustificationValues.Right));

body.AppendChild(new EmptyLine());

// // // // // // // // // // // // // // // //
// Transformations
// // // // // // // // // // // // // // // //

body.AppendChild(new Paragraph(
	new Run(QDrawing.FromImage(body, "Icon.png", ImagePartType.Png, 128, 128)
		.ToAnchoredDrawing()
		.HorizontalAlignment(DW.HorizontalAlignmentValues.Center, DW.HorizontalRelativePositionValues.Page)
		.VerticalAlignment(DW.VerticalAlignmentValues.Center, DW.VerticalRelativePositionValues.Paragraph)
		.TopAndBottomWrapping(0, 0, ImageMeasuringUnits.Centimeters)
		.Opacity(0.4)),

	new Run().Text("This image is 60% transparent.")
).Justification(JustificationValues.Center));

body.AppendChild(new EmptyLine());

body.AppendChild(new Paragraph(
	new Run(QDrawing.FromImage(body, "Icon.png", ImagePartType.Png, 128, 128)
		.ToAnchoredDrawing()
		.HorizontalAlignment(DW.HorizontalAlignmentValues.Center, DW.HorizontalRelativePositionValues.Page)
		.VerticalAlignment(DW.VerticalAlignmentValues.Center, DW.VerticalRelativePositionValues.Paragraph)
		.TopAndBottomWrapping(0, 1.0, ImageMeasuringUnits.Centimeters)
		.Rotation(45)),

	new Run().Text("This image is rotated by 45 degrees.")
).Justification(JustificationValues.Center));

body.AppendChild(new EmptyLine());

body.AppendChild(new Paragraph(
	new Run(QDrawing.FromImage(body, "Icon.png", ImagePartType.Png, 128, 128)
		.ToAnchoredDrawing()
		.HorizontalAlignment(DW.HorizontalAlignmentValues.Center, DW.HorizontalRelativePositionValues.Page)
		.VerticalAlignment(DW.VerticalAlignmentValues.Center, DW.VerticalRelativePositionValues.Paragraph)
		.TopAndBottomWrapping(0, 0.5, ImageMeasuringUnits.Centimeters)
		.Border(5, "Red")
		.Cropping(0, 0.15, 0.3, 0.15)),

	new Run().Text("This image is cropped by 15% on top & bottom and by 30% on the right.\nIt also has a red border applied to it.")
).Justification(JustificationValues.Center));

// // // // // // // // // // // // // // // //
// Page break
// // // // // // // // // // // // // // // //

body.AppendChild(new PageBreak());
body.AppendChild(QParagraph.Create("This text has a page break before.",
	new ParagraphFormatting { Justification = JustificationValues.Center }));

// // // // // // // // // // // // // // // //
// Horizontal line
// // // // // // // // // // // // // // // //

body.AppendChild(new HorizontalLine());
body.AppendChild(QParagraph.Create("This text has a horizontal line before.",
	new ParagraphFormatting { Justification = JustificationValues.Center }));

// // // // // // // // // // // // // // // //
// Save and run
// // // // // // // // // // // // // // // //

document.Save();

var process = new ProcessStartInfo("WINWORD.exe", fileName) { UseShellExecute = true };
Process.Start(process);

// // // // // // // // // // // // // // // //
// Helping methods
// // // // // // // // // // // // // // // //

void StartSection(string text, bool lineBreakBefore = true)
{
	if (lineBreakBefore)
		body.AppendChild(new EmptyLine());

	body.AppendChild(QParagraph.Create(text,
		new ParagraphFormatting
		{
			OutlineLevel = OutlineLevelValues.Level1,
			Justification = JustificationValues.Center
		},
		new RunFormatting { FontSize = 16 }));
}

record Person(int Id, string FirstName, string LastName, bool Verified);
