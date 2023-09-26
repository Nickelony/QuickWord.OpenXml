using DocumentFormat.OpenXml.Wordprocessing;

namespace QuickWord.OpenXml.Extras;

public class EmptyLine : Paragraph
{
	public EmptyLine(double fontSize = 11, double spacingBefore = 0, double spacingAfter = 8)
	{
		AppendChild(new Run().FontSize(fontSize));

		this.SpacingBefore(spacingBefore);
		this.SpacingAfter(spacingAfter);
	}
}
