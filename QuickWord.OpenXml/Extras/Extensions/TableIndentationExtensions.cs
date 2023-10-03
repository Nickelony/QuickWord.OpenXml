using DocumentFormat.OpenXml.Wordprocessing;
using System;

namespace QuickWord.OpenXml.Extras;

/// <summary>
/// A set of extension methods for the <see cref="TableIndentation"/> class.
/// </summary>
public static class TableIndentationExtensions
{
	public static double GetExactWidth(this TableIndentation tableIndentation, out WidthUnits? units)
	{
		units = tableIndentation.Type?.Value switch
		{
			TableWidthUnitValues.Pct => WidthUnits.Percentage,
			TableWidthUnitValues.Dxa => WidthUnits.Points,
			_ => WidthUnits.Auto
		};

		double widthValue = double.TryParse(tableIndentation.Width, out double result) ? result : 0;

		return units switch
		{
			WidthUnits.Percentage => widthValue / 50,
			WidthUnits.Points => widthValue / 20,
			_ => widthValue
		};
	}

	public static TableIndentation SetExactWidth(this TableIndentation tableIndentation, double width, WidthUnits? units)
	{
		tableIndentation.Width = (int)(units switch
		{
			WidthUnits.Percentage => Math.Round(width * 50),
			WidthUnits.Points => Math.Round(width * 20),
			_ => width
		});

		tableIndentation.Type = units switch
		{
			WidthUnits.Percentage => TableWidthUnitValues.Pct,
			WidthUnits.Points => TableWidthUnitValues.Dxa,
			WidthUnits.Auto => TableWidthUnitValues.Auto,
			_ => TableWidthUnitValues.Nil
		};

		return tableIndentation;
	}
}
