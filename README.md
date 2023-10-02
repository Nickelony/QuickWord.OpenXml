# ![Icon32x32](https://github.com/Nickelony/QuickWord.OpenXml/assets/20436882/9fb9f9c8-dc60-46dc-9d04-68a9ff60146f) QuickWord for DocumentFormat.OpenXml
A set of extension methods for `DocumentFormat.OpenXml` which simplifies creating and modifying Word documents (such as .docx).
You can use these methods with your existing `DocumentFormat.OpenXml` code without having to change anything.

# Features
- Super quick and easy Builder-like pattern:
```cs
var run = new Run().Text("This is a simple Run.")
  .FontSize(24)
  .FontColor("Red")
  .Bold()
  .Italic();
```
---
- The ability to create Paragraphs with multiple Runs and multiple formattings:
```cs
Run[] runs = new[]
{
	new Run().Text("This is a").FontSize(16),
	new Run().Text(" single paragraph ").FontFace("Comic Sans MS"),
	new Run().Text("with 3 Runs.").HighlightColor(HighlightColorValues.Cyan)
};

body.AppendChild(new Paragraph(runs))
```
Example with 5 runs:

![image](https://github.com/Nickelony/QuickWord.OpenXml/assets/20436882/85ddb350-834f-41d2-b2a8-e46e54bbe42f)

---
- The ability to create sections:
```cs
// Start a new section
body.AppendChild(QParagraph.Create(text,
	new ParagraphFormatting
	{
		OutlineLevel = OutlineLevelValues.Level1, // Section level 1
		Justification = JustificationValues.Center
	},
	new RunFormatting { FontSize = 16 }));
```
---
- The ability to create tables quickly:
```cs
body.AppendChild(QTable.Create(new[]
{
	new[] { "1", "John", "Doe" },
	new[] { "2", "Jane", "Doe" },
	new[] { "3", "John", "Smith" },
	new[] { "4", "Jane", "Smith" }
}));
```
Example:

![image](https://github.com/Nickelony/QuickWord.OpenXml/assets/20436882/211e54e6-c6d8-4db4-93b0-5a85c7b72591)

---
- The ability quickly merge cells:
```cs
// Merge cells [0,0] and [1,0] (vertically)
table.Rows(0)?.Cells(0)?.VerticalMerge(MergedCellValues.Restart);
table.Rows(1)?.Cells(0)?.VerticalMerge(MergedCellValues.Continue);

// Merge cells [2,0] and [3,0] (vertically)
table.Rows(2)?.Cells(0)?.VerticalMerge(MergedCellValues.Restart);
table.Rows(3)?.Cells(0)?.VerticalMerge(MergedCellValues.Continue);
```
Example:

![image](https://github.com/Nickelony/QuickWord.OpenXml/assets/20436882/ce0cbcd7-e7fc-4f73-bcb2-95d589de3e02)

---
- The ability to quickly create formatted rows and cells:
```cs
table1.AppendChild(QTableRow.Create(
	new[] { "ID", "First Name", "Last Name" },
	rowFormatting: new TableRowFormatting { IsHeader = true },
	runFormatting: new RunFormatting { Bold = true }));
```
- The ability to easily customize existing table cells:
```cs
foreach (TableCell? cell in table2.GetColumnOfCells(3).Skip(1)) // 4th column of cells, skip header cell
{
	if (cell?.Paragraphs(0)?.GetText() == "Yes")
		cell?.FillColor("LightGreen");
	else
		cell?.FillColor("LightPink");
}
```
Example:

![image](https://github.com/Nickelony/QuickWord.OpenXml/assets/20436882/8615e12c-baa7-4af4-8ff8-1df516c8190b)

---
- The ability to create inlined images:
```cs
body.AppendChild(new Paragraph(
	new Run(QDrawing.FromImage(body, "Icon.png", ImagePartType.Png, 32, 32)),
	new Run().Text(" This image is inlined with the text.").VerticalPosition(8)
).Justification(JustificationValues.Center));
```
Example:

![image](https://github.com/Nickelony/QuickWord.OpenXml/assets/20436882/8743289e-91a3-4de7-b296-d765478655d6)

---
- The ability to create anchored images:
```cs
body.AppendChild(new Paragraph(
	new Run(QDrawing.FromImage(body, "Icon.png", ImagePartType.Png, 128, 128)
		.ToAnchoredDrawing()
		.AbsoluteVerticalPosition(0, ImageMeasuringUnits.Pixels, DW.VerticalRelativePositionValues.Paragraph)
		.SquareWrapping(0, 0, 0.5, 0, ImageMeasuringUnits.Centimeters)),

	new Run().Text("This\nimage\nis\naligned\nbefore\nthe text.\n")
));
```
Example:

![image](https://github.com/Nickelony/QuickWord.OpenXml/assets/20436882/a52e8424-f6bb-45ae-8fd2-f83ec21755a4)

With images you can also:
  - Set the transparency
  - Set the rotation
  - Set cropping
  - Set a border
  - Set text wrapping
  - Set the position
  - many more...

![image](https://github.com/Nickelony/QuickWord.OpenXml/assets/20436882/89883b1b-04dd-43b4-b0a3-3ebc7240cb95)

---
- You can also use additional `Paragraph` inherited objects, such as:
  - `EmptyLine`
  - `PageBreak`
  - `HorizontalLine`