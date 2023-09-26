using DocumentFormat.OpenXml.Wordprocessing;
using QuickWord.OpenXml.Extras;

namespace QuickWord.OpenXml.QuickObjects;

public static class QTable
{
	public static Table Create(string?[][] data,
		TableFormatting? tableFormatting = null, TableRowFormatting? rowFormatting = null,
		TableCellFormatting? cellFormatting = null, ParagraphFormatting? paragraphFormatting = null,
		RunFormatting? runFormatting = null, bool parseNewLineChars = true)
	{
		var table = new Table();

		if (tableFormatting is not null)
			table.ApplyFormatting(tableFormatting);

		foreach (string?[] rowData in data)
			table.AppendChild(QTableRow.Create(rowData, rowFormatting,
				cellFormatting, paragraphFormatting, runFormatting, parseNewLineChars));

		return table;
	}
}
