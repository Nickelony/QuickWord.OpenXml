using DocumentFormat.OpenXml.Wordprocessing;
using QuickWord.OpenXml.Extras;

namespace QuickWord.OpenXml.QuickObjects;

public static class QTableCell
{
	public static TableCell Create(string runText, TableCellFormatting? cellFormatting = null,
		ParagraphFormatting? paragraphFormatting = null, RunFormatting? runFormatting = null,
		bool parseNewLineChars = true)
	{
		var cell = new TableCell();
		var paragraph = new Paragraph();
		Run run = new Run().Text(runText, parseNewLineChars);

		if (cellFormatting is not null)
			cell.ApplyFormatting(cellFormatting);

		if (paragraphFormatting is not null)
			paragraph.ApplyFormatting(paragraphFormatting);

		if (runFormatting is not null)
			run.ApplyFormatting(runFormatting);

		paragraph.AppendChild(run);
		cell.AppendChild(paragraph);

		return cell;
	}
}
