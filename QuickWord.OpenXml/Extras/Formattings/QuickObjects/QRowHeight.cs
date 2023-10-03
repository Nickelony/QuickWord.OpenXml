using DocumentFormat.OpenXml.Wordprocessing;

namespace QuickWord.OpenXml.Extras;

public class QRowHeight
{
	public double Height { get; set; }
	public MeasuringUnits Units { get; set; }
	public HeightRuleValues Rule { get; set; }

	public QRowHeight(double height, MeasuringUnits units, HeightRuleValues rule)
	{
		Height = height;
		Units = units;
		Rule = rule;
	}
}
