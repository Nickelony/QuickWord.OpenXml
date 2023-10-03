namespace QuickWord.OpenXml.Extras;

public class QParagraphIndentation
{
	public double Indentation { get; }
	public IndentationUnits Units { get; }

	public QParagraphIndentation(double indentation, IndentationUnits units)
	{
		Indentation = indentation;
		Units = units;
	}
}
