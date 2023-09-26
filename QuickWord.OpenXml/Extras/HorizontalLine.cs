using DocumentFormat.OpenXml.Wordprocessing;
using QuickWord.OpenXml.Measurements;

namespace QuickWord.OpenXml.Extras;

public class HorizontalLine : Paragraph
{
	public HorizontalLine(HorizontalLinePosition linePosition = HorizontalLinePosition.Bottom,
		double size = 1, BorderValues type = BorderValues.Single, string color = "auto")
	{
		var paraProperties = new ParagraphProperties();
		var paraBorders = new ParagraphBorders();

		BorderType border = linePosition switch
		{
			HorizontalLinePosition.Top => new TopBorder()
			{
				Val = type,
				Color = color,
				Size = BorderSize.ToSixth(size)
			},
			_ => new BottomBorder()
			{
				Val = type,
				Color = color,
				Size = BorderSize.ToSixth(size)
			}
		};

		paraBorders.AppendChild(border);
		paraProperties.AppendChild(paraBorders);

		AppendChild(paraProperties);
	}
}
