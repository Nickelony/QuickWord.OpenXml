using DocumentFormat.OpenXml.Wordprocessing;

namespace QuickWord.OpenXml.Extras;

public class PageBreak : Paragraph
{
	public PageBreak()
		=> AppendChild(new Run(new Break() { Type = BreakValues.Page }));
}
