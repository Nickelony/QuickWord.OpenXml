using System;

namespace QuickWord.OpenXml.Measurements;

/// <summary>
/// Twentieth of a point (0.05pt).
/// </summary>
public static class Twips
{
	public static double ToOther(int twips, MeasuringUnits desiredUnits) => desiredUnits switch
	{
		MeasuringUnits.Centimeters => Math.Round(twips / (1440 / 2.54), 2), // ~567tw => 1cm
		MeasuringUnits.Inches => twips / 1440.0, // 1440tw => 1"
		MeasuringUnits.Points => twips / 20.0, // 20tw => 1pt
		_ => twips // Same unit
	};

	public static double ToOther(int twips, LineMeasuringUnits desiredUnits) => desiredUnits switch
	{
		LineMeasuringUnits.WholeLines => twips / 240.0, // 240tw => 1ln
		LineMeasuringUnits.Points => twips / 20.0, // 20tw => 1pt
		_ => twips // Same unit
	};

	public static double ToOther(int twips, IndentationUnits desiredUnits) => desiredUnits switch
	{
		IndentationUnits.Centimeters => Math.Round(twips / (1440 / 2.54), 2), // ~567tw => 1cm
		IndentationUnits.Inches => twips / 1440.0, // 1440tw => 1"
		IndentationUnits.Points => twips / 20.0, // 20tw => 1pt
		_ => twips
	};

	public static int FromOther(double other, MeasuringUnits fromUnits) => fromUnits switch
	{
		MeasuringUnits.Centimeters => (int)(other * Math.Round(1440 / 2.54)), // 1cm => ~567tw
		MeasuringUnits.Inches => (int)(other * 1440), // 1" => 1440tw
		MeasuringUnits.Points => (int)(other * 20), // 1pt => 20tw
		_ => (int)other // Same unit
	};

	public static int FromOther(double other, LineMeasuringUnits fromUnits) => fromUnits switch
	{
		LineMeasuringUnits.WholeLines => (int)(other * 240), // 1ln => 240tw
		LineMeasuringUnits.Points => (int)(other * 20), // 1pt => 20tw
		_ => (int)other // Same unit
	};

	public static int FromOther(double other, IndentationUnits fromUnits) => fromUnits switch
	{
		IndentationUnits.Centimeters => (int)(other * Math.Round(1440 / 2.54)), // 1cm => ~567tw
		IndentationUnits.Inches => (int)(other * 1440), // 1" => 1440tw
		IndentationUnits.Points => (int)(other * 20), // 1pt => 20tw
		_ => (int)other // Same unit
	};
}
