using DocumentFormat.OpenXml.Wordprocessing;

namespace QuickWord.OpenXml.Extras;

public class QUnderline
{
	public UnderlineValues Style { get; }
	public string Color { get; }

	public QUnderline(UnderlineValues style = UnderlineValues.Single, string htmlColor = "auto")
	{
		Style = style;
		Color = htmlColor;
	}
}
