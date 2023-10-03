using DocumentFormat.OpenXml.Wordprocessing;

namespace QuickWord.OpenXml.Extras;

public class QBorder
{
	public double Width { get; }
	public BorderValues Border { get; }
	public string Color { get; }
	public uint Spacing { get; }

	public QBorder(double width, BorderValues type = BorderValues.Single, string htmlColor = "auto", uint spacing = 0)
	{
		Width = width;
		Border = type;
		Color = htmlColor;
		Spacing = spacing;
	}
}
