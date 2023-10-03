namespace QuickWord.OpenXml.Extras;

public class QWidth
{
	public double Width { get; set; }
	public WidthUnits? Units { get; set; }

	public QWidth(double width, WidthUnits? units)
	{
		Width = width;
		Units = units;
	}
}
