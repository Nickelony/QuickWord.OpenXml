namespace QuickWord.OpenXml.Extras;

public class QParagraphSpacing
{
	public double Spacing { get; }
	public LineMeasuringUnits Units { get; }

	public QParagraphSpacing(double spacing, LineMeasuringUnits units)
	{
		Spacing = spacing;
		Units = units;
	}
}
