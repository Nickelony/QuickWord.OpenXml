using DocumentFormat.OpenXml.Wordprocessing;
using System;

namespace QuickWord.OpenXml.Extras;

public static class TableWidthTypeExtensions
{
	public static double GetExactWidth(this TableWidthType tableWidth, out WidthUnits? units)
	{
		units = tableWidth.Type?.Value switch
		{
			TableWidthUnitValues.Pct => WidthUnits.Percentage,
			TableWidthUnitValues.Dxa => WidthUnits.Points,
			_ => WidthUnits.Auto
		};

		double widthValue = double.TryParse(tableWidth.Width, out double result) ? result : 0;

		return units switch
		{
			WidthUnits.Percentage => widthValue / 50,
			WidthUnits.Points => widthValue / 20,
			_ => widthValue
		};
	}

	public static TableWidthType SetExactWidth(this TableWidthType tableWidth, double width, WidthUnits? units)
	{
		tableWidth.Width = (units switch
		{
			WidthUnits.Percentage => Math.Round(width * 50),
			WidthUnits.Points => Math.Round(width * 20),
			_ => width
		}).ToString();

		tableWidth.Type = units switch
		{
			WidthUnits.Percentage => TableWidthUnitValues.Pct,
			WidthUnits.Points => TableWidthUnitValues.Dxa,
			WidthUnits.Auto => TableWidthUnitValues.Auto,
			_ => TableWidthUnitValues.Nil
		};

		return tableWidth;
	}
}
