using DocumentFormat.OpenXml.Wordprocessing;
using QuickWord.OpenXml.Extras;

namespace QuickWord.OpenXml.QuickObjects;

public static class QParagraph
{
	public static Paragraph Create(string runText,
		ParagraphFormatting? paragraphFormatting = null, RunFormatting? runFormatting = null,
		bool parseNewLineChars = true)
	{
		var paragraph = new Paragraph();
		Run run = new Run().Text(runText, parseNewLineChars);

		if (paragraphFormatting is not null)
			paragraph.ApplyFormatting(paragraphFormatting);

		if (runFormatting is not null)
			run.ApplyFormatting(runFormatting);

		paragraph.AppendChild(run);
		return paragraph;
	}
}
