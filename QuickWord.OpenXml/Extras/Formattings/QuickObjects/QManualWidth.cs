namespace QuickWord.OpenXml.Extras;

public class QManualWidth
{
	public double Width { get; }
	public MeasuringUnits Units { get; }

	public QManualWidth(double width, MeasuringUnits units = MeasuringUnits.Points)
	{
		Width = width;
		Units = units;
	}
}
