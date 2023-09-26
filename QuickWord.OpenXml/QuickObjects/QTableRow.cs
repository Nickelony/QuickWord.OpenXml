using DocumentFormat.OpenXml.Wordprocessing;
using QuickWord.OpenXml.Extras;

namespace QuickWord.OpenXml.QuickObjects;

public static class QTableRow
{
	public static TableRow Create(string?[] data,
		TableRowFormatting? rowFormatting = null, TableCellFormatting? cellFormatting = null,
		ParagraphFormatting? paragraphFormatting = null, RunFormatting? runFormatting = null,
		bool parseNewLineChars = true)
	{
		var row = new TableRow();

		if (rowFormatting is not null)
			row.ApplyFormatting(rowFormatting);

		foreach (string? cellData in data)
			row.AppendChild(QTableCell.Create(cellData ?? "",
				cellFormatting, paragraphFormatting, runFormatting, parseNewLineChars));

		return row;
	}
}
